"""Domain models for register map and memory map specifications."""

from __future__ import annotations

from sqlalchemy import ForeignKey, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from excel_toolkit.models import Base, ExcelSheet


# -- Register (level2_* sheets) --------------------------------------------

class Register(Base):
    __tablename__ = "register"

    id: Mapped[int] = mapped_column(primary_key=True)
    type: Mapped[str | None] = mapped_column(Text)      # RW2, RW1, RO, WO
    indx: Mapped[str | None] = mapped_column(Text)      # hex index e.g. "57"
    page: Mapped[str | None] = mapped_column(Text)      # page number as string
    para: Mapped[str | None] = mapped_column(Text)      # parameter index "0","1",...
    name: Mapped[str | None] = mapped_column(Text)      # register name
    d7: Mapped[str | None] = mapped_column(Text)
    d6: Mapped[str | None] = mapped_column(Text)
    d5: Mapped[str | None] = mapped_column(Text)
    d4: Mapped[str | None] = mapped_column(Text)
    d3: Mapped[str | None] = mapped_column(Text)
    d2: Mapped[str | None] = mapped_column(Text)
    d1: Mapped[str | None] = mapped_column(Text)
    d0: Mapped[str | None] = mapped_column(Text)
    init: Mapped[str | None] = mapped_column(Text)      # initial value e.g. "0x00"

    # Excel origin tracking
    sheet_id: Mapped[int | None] = mapped_column(ForeignKey("excel_sheet.id"))
    excel_row: Mapped[int | None]

    sheet: Mapped[ExcelSheet | None] = relationship()


REGMAP_FIELD_MAP: dict[str, str] = {
    "TYPE": "type",
    "INDX": "indx",
    "PAGE": "page",
    "PARA": "para",
    "NAME": "name",
    "D7": "d7",
    "D6": "d6",
    "D5": "d5",
    "D4": "d4",
    "D3": "d3",
    "D2": "d2",
    "D1": "d1",
    "D0": "d0",
    "INIT": "init",
}


# -- MemoryMapEntry (memorymap sheet) --------------------------------------

class MemoryMapEntry(Base):
    __tablename__ = "memorymap_entry"

    id: Mapped[int] = mapped_column(primary_key=True)
    baseaddr: Mapped[str | None] = mapped_column(Text)     # e.g. "0xB0"
    group: Mapped[str | None] = mapped_column(Text)         # same as NAME in level2
    midgroup: Mapped[str | None] = mapped_column(Text)
    comment: Mapped[str | None] = mapped_column(Text)
    special: Mapped[str | None] = mapped_column(Text)

    # Excel origin tracking
    sheet_id: Mapped[int | None] = mapped_column(ForeignKey("excel_sheet.id"))
    excel_row: Mapped[int | None]

    sheet: Mapped[ExcelSheet | None] = relationship()


MEMMAP_FIELD_MAP: dict[str, str] = {
    "BASEADDR": "baseaddr",
    "Group": "group",
    "midgroup": "midgroup",
    "Comment": "comment",
    "special": "special",
}
