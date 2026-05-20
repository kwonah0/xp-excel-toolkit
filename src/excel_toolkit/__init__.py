"""excel_toolkit — Excel(.xlsx/.xls) ↔ SQLite round-trip infrastructure.

Persists workbook + sheets + cells (with style/comment/merge/formula) into a
SQLAlchemy schema, with the original binary kept in ``excel_workbook.blob`` so
that exports preserve the source formatting.

Bring your own domain models: subclass :class:`Base`, declare them before
``init_db()``, and pass a ``sheet_configs`` dict to :func:`import_xlsx` to map
sheet-name patterns onto your models.
"""

from excel_toolkit.convert import (
    convert_xls_to_xlsx,
    ensure_xlsx_cached,
    resolve_db,
    validate_xlsx_format,
)
from excel_toolkit.exporter import export_from_cells
from excel_toolkit.merge import MergeResolver
from excel_toolkit.models import (
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
from excel_toolkit.xls_parser import import_xls
from excel_toolkit.xlsx_parser import (
    SheetConfig,
    extract_style,
    find_header_row,
    import_sheet,
    import_xlsx,
    register_domain,
)

__all__ = [
    # Models / DB
    "AUDIT_TARGETS",
    "Base",
    "ChangeLog",
    "ExcelCell",
    "ExcelMerge",
    "ExcelSheet",
    "ExcelWorkbook",
    "SheetConfigEntry",
    "init_db",
    "register_audit_target",
    # Parsing
    "SheetConfig",
    "extract_style",
    "find_header_row",
    "import_sheet",
    "import_xls",
    "import_xlsx",
    "register_domain",
    # Merge
    "MergeResolver",
    # Export
    "export_from_cells",
    # Conversion / cache
    "convert_xls_to_xlsx",
    "ensure_xlsx_cached",
    "resolve_db",
    "validate_xlsx_format",
]
