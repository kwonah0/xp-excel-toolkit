"""Ingest layer — file boundary (convert/validate/cache) + parsers (file → cell tables)."""

from xp_excel_toolkit.ingest.convert import (
    cache_key,
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

__all__ = [
    "BULK_CHUNK",
    "cache_key",
    "convert_xls_to_xlsx",
    "ensure_xlsx_cached",
    "extract_cell_value",
    "extract_style",
    "find_header_row",
    "import_sheet",
    "import_xls",
    "import_xlsx",
    "validate_xlsx_format",
]
