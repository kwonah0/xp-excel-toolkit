"""XLS → XLSX conversion using LibreOffice + format validation + cache."""

from __future__ import annotations

import hashlib
import shutil
import subprocess
import warnings
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from xp_excel_toolkit import config

_FALLBACK_CACHE_DIR = ".xltk_cache"

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


# ── LibreOffice detection ────────────────────────────────────────────

def _find_libreoffice() -> str:
    """Find LibreOffice executable.

    Priority:
        1. config.LIBREOFFICE_PATH
        2. 'libreoffice' / 'soffice' on PATH
        3. Common installation paths
    """
    if config.LIBREOFFICE_PATH:
        p = Path(config.LIBREOFFICE_PATH)
        if p.exists():
            return str(p)
        raise FileNotFoundError(
            f"LibreOffice not found at configured path: {config.LIBREOFFICE_PATH}"
        )

    for name in ("libreoffice", "soffice"):
        found = shutil.which(name)
        if found:
            return found

    for candidate in _SEARCH_PATHS:
        if Path(candidate).exists():
            return candidate

    raise FileNotFoundError(
        "LibreOffice not found. Install it or set xp_excel_toolkit.config.LIBREOFFICE_PATH."
    )


# ── Cache helpers ────────────────────────────────────────────────────

def get_cache_dir() -> Path:
    """Return the cache directory, creating it if needed.

    Resolution order:
        1. config.CACHE_DIR (set by downstream callers)
        2. <cwd>/.xltk_cache/  (fallback)
    """
    d = (
        Path(config.CACHE_DIR)
        if config.CACHE_DIR is not None
        else Path.cwd() / _FALLBACK_CACHE_DIR
    )
    d.mkdir(parents=True, exist_ok=True)
    return d


def cache_key(path: Path) -> tuple[str, str]:
    """Generate cache key (hash, mtime_str) from file path and mtime."""
    abs_path = path.resolve()
    mtime = abs_path.stat().st_mtime
    raw = f"{abs_path}_{mtime}"
    h = hashlib.sha256(raw.encode()).hexdigest()[:12]
    mtime_str = datetime.fromtimestamp(mtime).strftime("%Y%m%d_%H%M%S")
    return h, mtime_str


# ── XLS → XLSX conversion ───────────────────────────────────────────

def convert_xls_to_xlsx(
    xls_path: str | Path,
    output_dir: str | Path | None = None,
    timeout: int = 600,
) -> Path:
    """Convert .xls file to .xlsx using LibreOffice.

    Args:
        xls_path: Path to the .xls file.
        output_dir: Directory for the output .xlsx file.
                    If None, uses cache dir.
        timeout: Timeout in seconds for the conversion process.

    Returns:
        Path to the converted .xlsx file.
    """
    xls_path = Path(xls_path).resolve()
    if not xls_path.exists():
        raise FileNotFoundError(f"Input file not found: {xls_path}")

    lo = _find_libreoffice()

    if output_dir is None:
        output_dir = get_cache_dir()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    cmd = [
        lo,
        "--headless",
        "--convert-to", "xlsx",
        "--outdir", str(output_dir),
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

    # Find the converted file — check before raising on returncode,
    # because LibreOffice sometimes exits non-zero but still produces output.
    converted = output_dir / f"{xls_path.stem}.xlsx"
    if not converted.exists():
        xlsx_files = list(output_dir.glob(f"{xls_path.stem}*.xlsx"))
        converted = xlsx_files[0] if xlsx_files else None

    if converted is None:
        raise RuntimeError(
            f"LibreOffice conversion failed (exit {result.returncode}):\n"
            f"  stdout: {result.stdout}\n"
            f"  stderr: {result.stderr}"
        )

    if result.returncode != 0 and result.stderr:
        warnings.warn(
            f"LibreOffice exited with code {result.returncode}, "
            f"but output file was created.",
            RuntimeWarning,
            stacklevel=2,
        )

    return converted


# ── Format validation ────────────────────────────────────────────────

def validate_xlsx_format(path: Path) -> None:
    """Check that a .xlsx file is actually a valid ZIP (OOXML) file.

    Detects the common mistake of renaming a binary .xls file to .xlsx.
    """
    _OLE2_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
    _ZIP_MAGIC = b"PK"

    with open(path, "rb") as f:
        header = f.read(8)

    if header[:8] == _OLE2_MAGIC:
        raise ValueError(
            f"{path.name} is a binary .xls file renamed to .xlsx.\n"
            f"  Use the original .xls extension — toolkit will auto-convert via LibreOffice.\n"
            f"  Or convert manually: libreoffice --headless --convert-to xlsx '{path}'"
        )

    if header[:2] != _ZIP_MAGIC:
        raise ValueError(
            f"{path.name} is not a valid .xlsx file (expected ZIP/OOXML format).\n"
            f"  File header: {header[:4].hex()}"
        )


def ensure_xlsx_cached(
    path: Path,
    on_progress: Callable[[str], None] | None = None,
) -> Path:
    """If path is .xls, convert to .xlsx and cache in the cache dir.

    If path is already .xlsx, validates it is a real OOXML file (not a
    renamed binary .xls).
    Cached .xlsx files are reused when the source .xls has not been modified.
    """
    if path.suffix.lower() != ".xls":
        validate_xlsx_format(path)
        return path

    cache_dir = get_cache_dir()

    h, mtime_str = cache_key(path)
    cached_xlsx = cache_dir / f"{path.stem}_{h}_{mtime_str}.xlsx"

    if cached_xlsx.exists():
        if on_progress:
            on_progress(f"Using cached XLSX for {path.name} ({cached_xlsx.name})")
        return cached_xlsx

    if on_progress:
        on_progress(f"Converting {path.name} → .xlsx (LibreOffice)...")

    xlsx_path = convert_xls_to_xlsx(path, output_dir=cache_dir)

    # Rename to include hash/mtime in filename
    if xlsx_path != cached_xlsx:
        shutil.move(str(xlsx_path), str(cached_xlsx))

    if on_progress:
        on_progress(f"Converted: {cached_xlsx.name}")

    return cached_xlsx
