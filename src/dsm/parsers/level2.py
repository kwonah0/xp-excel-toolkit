"""Parser for level2_* sheets (register bit-field specifications)."""

from __future__ import annotations

from pathlib import Path

from sqlalchemy.orm import Session

from dsm.models import ExcelSheet
from dsm.domain_models import REGMAP_FIELD_MAP, Register
from dsm.xlsx_parser import import_sheet


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
    return import_sheet(
        session, path,
        sheet_name=sheet_name,
        field_map=REGMAP_FIELD_MAP,
        domain_cls=Register,
    )
