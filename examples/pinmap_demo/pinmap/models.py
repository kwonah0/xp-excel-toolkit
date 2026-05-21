"""Domain ORM model for pinmap.

The PinEntry table is declared against ``excel_toolkit.Base`` — that's
what guarantees ``init_db()`` sees infra tables AND pin_entry in the
same MetaData.create_all() pass. The host re-exports its own ``Base``
alias from :mod:`pinmap.api` so downstream code doesn't need to know
the dependency comes from excel_toolkit.
"""

from __future__ import annotations

from sqlalchemy import ForeignKey, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from excel_toolkit import Base, ExcelSheet, register_audit_target


class PinEntry(Base):
    __tablename__ = "pin_entry"

    id: Mapped[int] = mapped_column(primary_key=True)
    pin_no:    Mapped[str | None] = mapped_column(Text)
    name:      Mapped[str | None] = mapped_column(Text)
    direction: Mapped[str | None] = mapped_column(Text)

    sheet_id:  Mapped[int | None] = mapped_column(ForeignKey("excel_sheet.id"))
    excel_row: Mapped[int | None]
    sheet: Mapped[ExcelSheet | None] = relationship()


#: Excel header → domain field
PIN_FIELD_MAP = {
    "Pin": "pin_no",
    "Name": "name",
    "Dir": "direction",
}


# Register the audit target at import time. As long as :mod:`pinmap.api`
# (which transitively imports this module) is imported before init_db()
# runs, the SQLite UPDATE/DELETE triggers for pin_entry will be installed.
register_audit_target("pin_entry", list(PIN_FIELD_MAP.values()))
