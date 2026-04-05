"""Domain models for register map and memory map specifications."""

from __future__ import annotations

from sqlalchemy import ForeignKey, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from dsm.models import Base, ExcelSheet


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


# -- Domain class registry (tablename → class) --------------------------------

DOMAIN_REGISTRY: dict[str, type] = {
    "register": Register,
    "memorymap_entry": MemoryMapEntry,
}

FIELD_MAP_REGISTRY: dict[str, dict[str, str]] = {
    "register": REGMAP_FIELD_MAP,
    "memorymap_entry": MEMMAP_FIELD_MAP,
}


# -- Default sheet configs (pattern → domain mapping) -------------------------
# import_xlsx() uses this when sheet_configs is not provided.

# Default patterns seeded into the DB on first import.
DEFAULT_SHEET_CONFIGS = [
    {"pattern": "level2_*", "domain_type": "register",
     "field_map": REGMAP_FIELD_MAP, "header_row": None},
    {"pattern": "memorymap", "domain_type": "memorymap_entry",
     "field_map": MEMMAP_FIELD_MAP, "header_row": None},
]


def seed_default_configs(session) -> None:
    """Insert default SheetConfigEntry rows if table is empty."""
    import json
    from dsm.models import SheetConfigEntry

    if session.query(SheetConfigEntry).count() > 0:
        return

    for cfg in DEFAULT_SHEET_CONFIGS:
        session.add(SheetConfigEntry(
            pattern=cfg["pattern"],
            domain_type=cfg["domain_type"],
            field_map_json=json.dumps(cfg["field_map"]),
            header_row=cfg["header_row"],
        ))
    session.flush()


def _default_sheet_configs():
    from dsm.xlsx_parser import SheetConfig

    return {
        "level2_*": SheetConfig(field_map=REGMAP_FIELD_MAP, domain_cls=Register),
        "memorymap": SheetConfig(field_map=MEMMAP_FIELD_MAP, domain_cls=MemoryMapEntry),
    }
