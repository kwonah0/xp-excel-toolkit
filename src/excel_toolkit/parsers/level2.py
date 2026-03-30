"""Parser for level2_* sheets (register bit-field specifications)."""

from __future__ import annotations

from pathlib import Path

from sqlalchemy.orm import Session

from excel_toolkit.models import ExcelSheet
from excel_toolkit.domain_models import REGMAP_FIELD_MAP, Register
from excel_toolkit.xlsx_parser import import_xlsx


def parse_level2(
    session: Session,
    path: str | Path,
    sheet_name: str,
) -> ExcelSheet:
    """Import a level2_* sheet into the DB.

    Args:
        session: SQLAlchemy session.
        path: Path to the xlsx file.
        sheet_name: Name of the level2 sheet (e.g. "level2_common").

    Returns:
        The created ExcelSheet record.
    """
    return import_xlsx(
        session, path,
        sheet_name=sheet_name,
        field_map=REGMAP_FIELD_MAP,
        domain_cls=Register,
    )
