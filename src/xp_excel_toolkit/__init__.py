"""xp_excel_toolkit — Excel toolkit.

Generic Excel ↔ SQLite layer. Domain-agnostic: parses xlsx/xls into
ExcelWorkbook/Sheet/Cell/Merge ORM rows, exports cells back to xlsx,
and provides helpers for domain packages that build their own models
on top of the cell data.
"""

from xp_excel_toolkit.convert import (
    convert_xls_to_xlsx,
    ensure_xlsx_cached,
    resolve_db,
    validate_xlsx_format,
)
from xp_excel_toolkit.exporter import (
    apply_style,
    export_from_cells,
)
from xp_excel_toolkit.helpers import (
    MERGE_POLICIES,
    MergePolicy,
    MergeWriteConflict,
    detect_merge_conflicts,
    find_header_row_db,
    iter_rows_by_header,
    resolve_conflicts,
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
from xp_excel_toolkit.xlsx_parser import (
    BULK_CHUNK,
    extract_cell_value,
    extract_style,
    find_header_row,
    import_sheet,
    import_xlsx,
)
from xp_excel_toolkit.xls_parser import import_xls

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
