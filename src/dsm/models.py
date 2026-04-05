"""SQLAlchemy ORM models for Excel metadata and domain data."""

from __future__ import annotations

from sqlalchemy import ForeignKey, Index, Text, create_engine
from sqlalchemy.orm import (
    DeclarativeBase, Mapped, mapped_column, relationship, sessionmaker,
)
from sqlalchemy.types import JSON


class Base(DeclarativeBase):
    pass


# ── Excel metadata models ──────────────────────────────────────────

class ExcelWorkbook(Base):
    __tablename__ = "excel_workbook"

    id: Mapped[int] = mapped_column(primary_key=True)
    filename: Mapped[str]
    blob: Mapped[bytes | None]  # 원본 바이너리 (round-trip용)

    sheets: Mapped[list[ExcelSheet]] = relationship(
        back_populates="workbook", cascade="all, delete-orphan"
    )


class ExcelSheet(Base):
    __tablename__ = "excel_sheet"

    id: Mapped[int] = mapped_column(primary_key=True)
    workbook_id: Mapped[int] = mapped_column(ForeignKey("excel_workbook.id"))
    name: Mapped[str]
    header_row: Mapped[int | None]

    workbook: Mapped[ExcelWorkbook] = relationship(back_populates="sheets")
    cells: Mapped[list[ExcelCell]] = relationship(
        back_populates="sheet", cascade="all, delete-orphan"
    )
    merges: Mapped[list[ExcelMerge]] = relationship(
        back_populates="sheet", cascade="all, delete-orphan"
    )


class ExcelMerge(Base):
    __tablename__ = "excel_merge"

    id: Mapped[int] = mapped_column(primary_key=True)
    sheet_id: Mapped[int] = mapped_column(ForeignKey("excel_sheet.id"))
    min_row: Mapped[int]
    min_col: Mapped[int]
    max_row: Mapped[int]
    max_col: Mapped[int]

    sheet: Mapped[ExcelSheet] = relationship(back_populates="merges")


class ExcelCell(Base):
    __tablename__ = "excel_cell"

    id: Mapped[int] = mapped_column(primary_key=True)
    sheet_id: Mapped[int] = mapped_column(ForeignKey("excel_sheet.id"))
    row: Mapped[int]
    col: Mapped[int]
    raw_value: Mapped[str | None] = mapped_column(Text)
    style: Mapped[dict | None] = mapped_column(JSON)  # {bg_color, font_bold, number_format, ...}
    comment: Mapped[str | None] = mapped_column(Text)  # Excel cell note/comment text
    merge_id: Mapped[int | None] = mapped_column(ForeignKey("excel_merge.id"))
    is_merge_origin: Mapped[bool] = mapped_column(default=False)

    sheet: Mapped[ExcelSheet] = relationship(back_populates="cells")
    merge: Mapped[ExcelMerge | None] = relationship()

    __table_args__ = (
        Index("ix_cell_sheet_row_col", "sheet_id", "row", "col", unique=True),
    )


# ── DB setup helper ─────────────────────────────────────────────────

# ── Sheet config (stored in DB) ─────────────────────────────────────

class SheetConfigEntry(Base):
    """Per-sheet import configuration stored in the database."""
    __tablename__ = "sheet_config"

    id: Mapped[int] = mapped_column(primary_key=True)
    pattern: Mapped[str] = mapped_column(Text)              # fnmatch pattern, e.g. "level2_*"
    domain_type: Mapped[str | None] = mapped_column(Text)   # "register" or "memorymap_entry"
    field_map_json: Mapped[str | None] = mapped_column(Text)  # JSON: {"TYPE": "type", ...}
    header_row: Mapped[int | None] = mapped_column(default=None)


# ── DB setup helper ─────────────────────────────────────────────────

def init_db(db_url: str = "sqlite:///excel_data.db"):
    engine = create_engine(db_url, echo=False)
    Base.metadata.create_all(engine)
    return sessionmaker(bind=engine)
