"""Pipeline orchestration — the only module allowed to know every layer.

Resolves a user-supplied file path (.db / .xlsx / .xls) into an imported
SQLite DB, chaining the ingest layer (convert → parse) with the schema
layer (init_db).
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Protocol

from sqlalchemy.orm import Session

from xp_excel_toolkit.ingest.convert import cache_key, ensure_xlsx_cached, get_cache_dir
from xp_excel_toolkit.ingest.xlsx import import_xlsx
from xp_excel_toolkit.models import init_db


class ImportFn(Protocol):
    def __call__(
        self,
        session: Session,
        path: Path,
        *,
        on_progress: Callable[[str], None] | None = None,
        with_formulas: bool = False,
    ) -> object: ...


def resolve_db(
    path: Path,
    on_progress: Callable[[str], None] | None = None,
    with_formulas: bool = False,
    *,
    import_fn: ImportFn | None = None,
) -> tuple[Path, bool]:
    """Resolve any file path to a DB path, auto-importing if needed.

    - .db → return as-is
    - .xlsx/.xls → ensure xlsx, import to <cache>/<stem>_<hash>.db

    Args:
        path: Source file (.db / .xlsx / .xls).
        on_progress: Optional progress callback.
        with_formulas: Whether to load formula strings during import.
        import_fn: Optional callable ``(session, xlsx_path, *, on_progress,
            with_formulas)``. Defaults to a cells-only import via
            :func:`xp_excel_toolkit.ingest.xlsx.import_xlsx`. Domain packages can
            pass their own pipeline (parse + build) here so that the cached
            DB includes their domain rows.

    Returns:
        (db_path, is_cached) — is_cached=True if DB is in cache dir.
    """
    if path.suffix == ".db":
        return path, False

    if path.suffix.lower() in (".xlsx", ".xls"):
        xlsx_path = ensure_xlsx_cached(path, on_progress=on_progress)

        cache_dir = get_cache_dir()

        h, mtime_str = cache_key(path)
        cached_db = cache_dir / f"{path.stem}_{h}_{mtime_str}.db"

        if cached_db.exists():
            if on_progress:
                on_progress(f"Using cached DB for {path.name} ({cached_db.name})")
            return cached_db, True

        if on_progress:
            on_progress(f"Importing {xlsx_path.name} into cache...")

        run_import = import_fn if import_fn is not None else import_xlsx
        session_factory = init_db(f"sqlite:///{cached_db}")
        with session_factory() as session:
            run_import(session, xlsx_path,
                       on_progress=on_progress,
                       with_formulas=with_formulas)
            session.commit()
        return cached_db, True

    raise ValueError(f"Unsupported file type: {path.suffix} (expected .db, .xlsx, or .xls)")
