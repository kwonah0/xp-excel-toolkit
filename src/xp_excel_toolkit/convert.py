"""XLS → XLSX conversion using LibreOffice + import cache resolution."""

from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

# ── Configuration ────────────────────────────────────────────────────
# LibreOffice executable path. Set this if auto-detection fails.
LIBREOFFICE_PATH: str | None = None

_CACHE_DIR_NAME = "__dsm__"

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
        1. LIBREOFFICE_PATH module variable (set above)
        2. 'libreoffice' / 'soffice' on PATH
        3. Common installation paths
    """
    if LIBREOFFICE_PATH:
        p = Path(LIBREOFFICE_PATH)
        if p.exists():
            return str(p)
        raise FileNotFoundError(
            f"LibreOffice not found at configured path: {LIBREOFFICE_PATH}"
        )

    for name in ("libreoffice", "soffice"):
        found = shutil.which(name)
        if found:
            return found

    for candidate in _SEARCH_PATHS:
        if Path(candidate).exists():
            return candidate

    raise FileNotFoundError(
        "LibreOffice not found. Install it or set xp_excel_toolkit.convert.LIBREOFFICE_PATH."
    )


# ── XLS → XLSX conversion ───────────────────────────────────────────

# Env vars consulted for cache override (DSM kept for backward compat with
# the dsm CLI's --cache-dir flag).
_CACHE_DIR_ENVS = ("XLTK_CACHE_DIR", "DSM_CACHE_DIR")


def _get_cache_dir() -> Path:
    """Return the cache directory, creating it if needed.

    Resolution order:
        1. $XLTK_CACHE_DIR / $DSM_CACHE_DIR (set by --cache-dir or shell)
        2. cwd / __dsm__/   (fallback, kept for backward compat)
    """
    for env in _CACHE_DIR_ENVS:
        override = os.environ.get(env)
        if override:
            d = Path(override)
            d.mkdir(parents=True, exist_ok=True)
            return d
    d = Path.cwd() / _CACHE_DIR_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d


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
        output_dir = _get_cache_dir()
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
        import sys
        print(f"Warning: LibreOffice exited with code {result.returncode}, "
              f"but output file was created.", file=sys.stderr)

    return converted


# ── Cache helpers ────────────────────────────────────────────────────

def cache_key(path: Path) -> tuple[str, str]:
    """Generate cache key (hash, mtime_str) from file path and mtime."""
    abs_path = path.resolve()
    mtime = abs_path.stat().st_mtime
    raw = f"{abs_path}_{mtime}"
    h = hashlib.sha256(raw.encode()).hexdigest()[:12]
    mtime_str = datetime.fromtimestamp(mtime).strftime("%Y%m%d_%H%M%S")
    return h, mtime_str


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


# ── Public API ───────────────────────────────────────────────────────

def resolve_db(
    path: Path,
    on_progress=None,
    with_formulas: bool = False,
    *,
    import_fn=None,
) -> tuple[Path, bool]:
    """Resolve any file path to a DB path, auto-importing if needed.

    - .db → return as-is
    - .xlsx/.xls → ensure xlsx, import to <cache>/<stem>_<hash>.db

    Args:
        path: Source file (.db / .xlsx / .xls).
        on_progress: Optional progress callback.
        with_formulas: Whether to load formula strings during import.
        import_fn: Optional callable ``(session, xlsx_path, *, on_progress,
            with_formulas) -> None``. Defaults to a cells-only import via
            :func:`xp_excel_toolkit.xlsx_parser.import_xlsx`. Domain packages can pass
            their own pipeline (parse + build) here so that the cached DB
            includes their domain rows.

    Returns:
        (db_path, is_cached) — is_cached=True if DB is in cache dir.
    """
    if path.suffix == ".db":
        return path, False

    if path.suffix.lower() in (".xlsx", ".xls"):
        xlsx_path = ensure_xlsx_cached(path, on_progress=on_progress)

        cache_dir = _get_cache_dir()

        h, mtime_str = cache_key(path)
        cached_db = cache_dir / f"{path.stem}_{h}_{mtime_str}.db"

        if cached_db.exists():
            msg = f"Using cached DB for {path.name} ({cached_db.name})"
            if on_progress:
                on_progress(msg)
            else:
                print(msg)
            return cached_db, True

        msg = f"Importing {xlsx_path.name} into cache..."
        if on_progress:
            on_progress(msg)
        else:
            print(msg)

        from xp_excel_toolkit.models import init_db
        from xp_excel_toolkit.xlsx_parser import import_xlsx

        Session = init_db(f"sqlite:///{cached_db}")
        with Session() as session:
            if import_fn is None:
                import_xlsx(session, xlsx_path, on_progress=on_progress,
                            with_formulas=with_formulas)
            else:
                import_fn(session, xlsx_path,
                          on_progress=on_progress,
                          with_formulas=with_formulas)
            session.commit()
        return cached_db, True

    raise ValueError(f"Unsupported file type: {path.suffix} (expected .db, .xlsx, or .xls)")


def ensure_xlsx_cached(
    path: Path,
    on_progress=None,
) -> Path:
    """If path is .xls, convert to .xlsx and cache in the cache dir.

    If path is already .xlsx, validates it is a real OOXML file (not a
    renamed binary .xls).
    Cached .xlsx files are reused when the source .xls has not been modified.
    """
    if path.suffix.lower() != ".xls":
        validate_xlsx_format(path)
        return path

    cache_dir = _get_cache_dir()

    h, mtime_str = cache_key(path)
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
