"""XLS → XLSX conversion using LibreOffice."""

from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path

# LibreOffice executable path — can be overridden via:
#   1. Environment variable: DSM_LIBREOFFICE_PATH
#   2. Directly setting dsm.convert.LIBREOFFICE_PATH
LIBREOFFICE_PATH: str | None = os.environ.get("DSM_LIBREOFFICE_PATH")

_SEARCH_PATHS = [
    "/usr/bin/libreoffice",
    "/usr/bin/soffice",
    "/usr/local/bin/libreoffice",
    "/usr/local/bin/soffice",
    "/snap/bin/libreoffice",
    # macOS
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    # Windows (WSL)
    "/mnt/c/Program Files/LibreOffice/program/soffice.exe",
]


def _find_libreoffice() -> str:
    """Find LibreOffice executable.

    Priority:
        1. LIBREOFFICE_PATH module variable
        2. DSM_LIBREOFFICE_PATH environment variable
        3. 'libreoffice' / 'soffice' on PATH
        4. Common installation paths
    """
    # 1. Module-level setting
    if LIBREOFFICE_PATH:
        p = Path(LIBREOFFICE_PATH)
        if p.exists():
            return str(p)
        raise FileNotFoundError(
            f"LibreOffice not found at configured path: {LIBREOFFICE_PATH}"
        )

    # 2. Environment variable (re-check in case set after module load)
    env_path = os.environ.get("DSM_LIBREOFFICE_PATH")
    if env_path:
        p = Path(env_path)
        if p.exists():
            return str(p)
        raise FileNotFoundError(
            f"LibreOffice not found at DSM_LIBREOFFICE_PATH: {env_path}"
        )

    # 3. On PATH
    for name in ("libreoffice", "soffice"):
        found = shutil.which(name)
        if found:
            return found

    # 4. Common paths
    for candidate in _SEARCH_PATHS:
        if Path(candidate).exists():
            return candidate

    raise FileNotFoundError(
        "LibreOffice not found. Install it or set DSM_LIBREOFFICE_PATH.\n"
        "  export DSM_LIBREOFFICE_PATH=/path/to/libreoffice"
    )


def convert_xls_to_xlsx(
    xls_path: str | Path,
    output_dir: str | Path | None = None,
    libreoffice_path: str | None = None,
    timeout: int = 120,
) -> Path:
    """Convert .xls file to .xlsx using LibreOffice.

    Args:
        xls_path: Path to the .xls file.
        output_dir: Directory for the output .xlsx file.
                    If None, uses a temp directory next to the input file.
        libreoffice_path: Override LibreOffice executable path for this call.
        timeout: Timeout in seconds for the conversion process.

    Returns:
        Path to the converted .xlsx file.

    Raises:
        FileNotFoundError: If LibreOffice is not found.
        RuntimeError: If conversion fails.
    """
    xls_path = Path(xls_path).resolve()
    if not xls_path.exists():
        raise FileNotFoundError(f"Input file not found: {xls_path}")

    lo = libreoffice_path or _find_libreoffice()

    # Use a temp directory for conversion to avoid conflicts
    with tempfile.TemporaryDirectory(prefix="dsm_convert_") as tmpdir:
        cmd = [
            lo,
            "--headless",
            "--convert-to", "xlsx",
            "--outdir", tmpdir,
            str(xls_path),
        ]

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout,
            )
        except subprocess.TimeoutExpired as e:
            raise RuntimeError(
                f"LibreOffice conversion timed out after {timeout}s"
            ) from e

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (exit {result.returncode}):\n"
                f"  stdout: {result.stdout}\n"
                f"  stderr: {result.stderr}"
            )

        # Find the converted file
        converted = Path(tmpdir) / f"{xls_path.stem}.xlsx"
        if not converted.exists():
            # Sometimes LibreOffice changes the name slightly
            xlsx_files = list(Path(tmpdir).glob("*.xlsx"))
            if not xlsx_files:
                raise RuntimeError(
                    f"Conversion produced no .xlsx file.\n"
                    f"  stdout: {result.stdout}\n"
                    f"  stderr: {result.stderr}"
                )
            converted = xlsx_files[0]

        # Move to output location
        if output_dir is None:
            output_dir = xls_path.parent
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        dest = output_dir / converted.name
        shutil.move(str(converted), str(dest))

    return dest


# ── Cached XLS → XLSX ─────────────────────────────────────────────

_CACHE_DIR_NAME = "__dsm_cache__"


def _xls_cache_key(xls_path: Path) -> tuple[str, str]:
    """Generate cache key (hash, mtime_str) from .xls file path and mtime."""
    abs_path = xls_path.resolve()
    mtime = abs_path.stat().st_mtime
    raw = f"{abs_path}_{mtime}"
    h = hashlib.sha256(raw.encode()).hexdigest()[:12]
    mtime_str = datetime.fromtimestamp(mtime).strftime("%Y%m%d_%H%M%S")
    return h, mtime_str


def ensure_xlsx_cached(
    path: Path,
    on_progress=None,
) -> Path:
    """If path is .xls, convert to .xlsx and cache in __dsm_cache__/.

    If path is already .xlsx, return it as-is.
    Cached .xlsx files are reused when the source .xls has not been modified.

    Args:
        path: Path to .xls or .xlsx file.
        on_progress: Optional callback for progress messages.

    Returns:
        Path to the .xlsx file (original or cached conversion).
    """
    if path.suffix.lower() != ".xls":
        return path

    cache_dir = Path.cwd() / _CACHE_DIR_NAME
    cache_dir.mkdir(exist_ok=True)

    h, mtime_str = _xls_cache_key(path)
    cached_xlsx = cache_dir / f"{path.stem}_{h}_{mtime_str}.xlsx"

    if cached_xlsx.exists():
        msg = f"Using cached XLSX for {path.name} ({cached_xlsx.name})"
        if on_progress:
            on_progress(msg)
        else:
            print(msg)
        return cached_xlsx

    msg = f"Converting {path.name} → .xlsx (LibreOffice)..."
    if on_progress:
        on_progress(msg)
    else:
        print(msg)

    xlsx_path = convert_xls_to_xlsx(path, output_dir=cache_dir)

    # Rename to include hash/mtime in filename
    if xlsx_path != cached_xlsx:
        shutil.move(str(xlsx_path), str(cached_xlsx))

    msg = f"Converted: {cached_xlsx.name}"
    if on_progress:
        on_progress(msg)
    else:
        print(msg)

    return cached_xlsx
