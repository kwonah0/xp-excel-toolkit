"""xp_excel_toolkit — Excel toolkit.

Generic Excel ↔ SQLite layer. Domain-agnostic: parses xlsx/xls into
ExcelWorkbook/Sheet/Cell/Merge ORM rows, exports cells back to xlsx,
and provides helpers for domain packages that build their own models
on top of the cell data.

Layers (each imports models only; pipeline is the sole orchestrator):
  ingest   — file boundary (convert/validate/cache) + parsers
  models   — 4-table cell schema (the contract)
  query    — read-only DB helpers
  export   — cell tables → xlsx + merge-write policies
  diff     — DB×DB cell diff (separate DiffBase)
  pipeline — resolve_db orchestration
"""

# Import config first so that ingest.convert (and any other submodule
# below) can safely `from xp_excel_toolkit import config` without hitting
# a partially-initialised package.
from xp_excel_toolkit import config
from xp_excel_toolkit.export.merge_policy import (
    MERGE_POLICIES,
    MergePolicy,
    MergeWriteConflict,
    detect_merge_conflicts,
    resolve_conflicts,
)
from xp_excel_toolkit.export.writer import (
    apply_style,
    export_from_cells,
)
from xp_excel_toolkit.ingest.convert import (
    convert_xls_to_xlsx,
    ensure_xlsx_cached,
    validate_xlsx_format,
)
from xp_excel_toolkit.ingest.xls import import_xls
from xp_excel_toolkit.ingest.xlsx import (
    BULK_CHUNK,
    extract_cell_value,
    extract_style,
    find_header_row,
    import_sheet,
    import_xlsx,
)
from xp_excel_toolkit.merge import MergeResolver
from xp_excel_toolkit.models import (
    AUDIT_TARGETS,
    Base,
    ChangeLog,
    ExcelCell,
    ExcelMerge,
    ExcelSheet,
    ExcelWorkbook,
    SheetConfigEntry,
    init_db,
    register_audit_target,
)
from xp_excel_toolkit.pipeline import resolve_db
from xp_excel_toolkit.query import (
    find_header_row_db,
    iter_rows_by_header,
)

__all__ = [
    "AUDIT_TARGETS",
    "BULK_CHUNK",
    "Base",
    "ChangeLog",
    "ExcelCell",
    "ExcelMerge",
    "ExcelSheet",
    "ExcelWorkbook",
    "MergeResolver",
    "SheetConfigEntry",
    "apply_style",
    "config",
    "convert_xls_to_xlsx",
    "ensure_xlsx_cached",
    "export_from_cells",
    "extract_cell_value",
    "extract_style",
    "find_header_row",
    "MERGE_POLICIES",
    "MergePolicy",
    "MergeWriteConflict",
    "detect_merge_conflicts",
    "find_header_row_db",
    "iter_rows_by_header",
    "resolve_conflicts",
    "import_sheet",
    "import_xls",
    "import_xlsx",
    "init_db",
    "register_audit_target",
    "resolve_db",
    "validate_xlsx_format",
]
