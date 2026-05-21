"""Toy host package that consumes excel_toolkit.

This single file stands in for what would normally be a small package
(``pinmap/models.py``, ``pinmap/importer.py``, ``pinmap/exporter.py``).
Everything a host package needs from excel_toolkit is touched here.
"""

from __future__ import annotations

from pathlib import Path

from sqlalchemy import ForeignKey, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from excel_toolkit import (
    Base,
    ExcelSheet,
    ExportHandler,
    SheetConfig,
    export_domain_xlsx,
    import_xlsx,
    register_audit_target,
)


# ── 1. Domain model — subclass excel_toolkit.Base so the table goes into
#       the same MetaData that init_db() will create_all() over.
class PinEntry(Base):
    __tablename__ = "pin_entry"

    id: Mapped[int] = mapped_column(primary_key=True)
    pin_no:    Mapped[str | None] = mapped_column(Text)
    name:      Mapped[str | None] = mapped_column(Text)
    direction: Mapped[str | None] = mapped_column(Text)

    # Excel-origin tracking — required for round-trip export to know
    # which sheet and which row the value came from.
    sheet_id:  Mapped[int | None] = mapped_column(ForeignKey("excel_sheet.id"))
    excel_row: Mapped[int | None]
    sheet: Mapped[ExcelSheet | None] = relationship()


# ── 2. Field map: Excel header → domain field name.
PIN_FIELD_MAP = {
    "Pin": "pin_no",
    "Name": "name",
    "Dir": "direction",
}


# ── 3. Audit registration — must run before init_db() so the SQLite
#       UPDATE/DELETE triggers are installed for this table.
register_audit_target("pin_entry", list(PIN_FIELD_MAP.values()))


# ── 4. SheetConfig: how to match sheets at import time.
SHEET_CONFIGS = {
    "Pinmap_*": SheetConfig(
        field_map=PIN_FIELD_MAP,
        domain_cls=PinEntry,
    ),
}


# ── 5. ExportHandler: how to write domain rows back on export.
EXPORT_HANDLERS = [
    ExportHandler(
        pattern="Pinmap_*",
        field_map=PIN_FIELD_MAP,
        domain_cls=PinEntry,
    ),
]


# ── 6. Thin wrappers around excel_toolkit's import/export so callers
#       don't have to know about the configs.
def import_pinmap(session, path: str | Path):
    return import_xlsx(session, path, sheet_configs=SHEET_CONFIGS)


def export_pinmap(session, output_path: str | Path) -> Path:
    return export_domain_xlsx(session, output_path, EXPORT_HANDLERS)
