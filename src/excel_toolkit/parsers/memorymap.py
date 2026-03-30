"""Parser for memorymap sheet."""

from __future__ import annotations

from pathlib import Path

from sqlalchemy.orm import Session

from excel_toolkit.domain_models import MEMMAP_FIELD_MAP, MemoryMapEntry
from excel_toolkit.models import ExcelSheet
from excel_toolkit.xlsx_parser import import_xlsx


def parse_memorymap(
    session: Session,
    path: str | Path,
    sheet_name: str = "memorymap",
) -> ExcelSheet:
    """Import the memorymap sheet into the DB.

    Args:
        session: SQLAlchemy session.
        path: Path to the xlsx file.
        sheet_name: Name of the memorymap sheet (default: "memorymap").

    Returns:
        The created ExcelSheet record.
    """
    return import_xlsx(
        session, path,
        sheet_name=sheet_name,
        field_map=MEMMAP_FIELD_MAP,
        domain_cls=MemoryMapEntry,
    )
