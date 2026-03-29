"""SQLAlchemy ORM models for Excel metadata and domain data."""

from sqlalchemy import (
    Boolean, Column, Float, ForeignKey, Index, Integer, LargeBinary,
    String, Text, create_engine,
)
from sqlalchemy.orm import DeclarativeBase, Session, relationship, sessionmaker
from sqlalchemy.types import JSON


class Base(DeclarativeBase):
    pass


# ── Excel metadata models ──────────────────────────────────────────

class ExcelWorkbook(Base):
    __tablename__ = "excel_workbook"

    id = Column(Integer, primary_key=True)
    filename = Column(String, nullable=False)
    blob = Column(LargeBinary, nullable=True)  # 원본 바이너리 (round-trip용)

    sheets = relationship(
        "ExcelSheet", back_populates="workbook", cascade="all, delete-orphan"
    )


class ExcelSheet(Base):
    __tablename__ = "excel_sheet"

    id = Column(Integer, primary_key=True)
    workbook_id = Column(Integer, ForeignKey("excel_workbook.id"), nullable=False)
    name = Column(String, nullable=False)
    header_row = Column(Integer, nullable=True)

    workbook = relationship("ExcelWorkbook", back_populates="sheets")
    cells = relationship(
        "ExcelCell", back_populates="sheet", cascade="all, delete-orphan"
    )
    merges = relationship(
        "ExcelMerge", back_populates="sheet", cascade="all, delete-orphan"
    )


class ExcelMerge(Base):
    __tablename__ = "excel_merge"

    id = Column(Integer, primary_key=True)
    sheet_id = Column(Integer, ForeignKey("excel_sheet.id"), nullable=False)
    min_row = Column(Integer, nullable=False)
    min_col = Column(Integer, nullable=False)
    max_row = Column(Integer, nullable=False)
    max_col = Column(Integer, nullable=False)

    sheet = relationship("ExcelSheet", back_populates="merges")


class ExcelCell(Base):
    __tablename__ = "excel_cell"

    id = Column(Integer, primary_key=True)
    sheet_id = Column(Integer, ForeignKey("excel_sheet.id"), nullable=False)
    row = Column(Integer, nullable=False)
    col = Column(Integer, nullable=False)
    raw_value = Column(Text, nullable=True)
    style = Column(JSON, nullable=True)  # {bg_color, font_bold, number_format, ...}
    merge_id = Column(Integer, ForeignKey("excel_merge.id"), nullable=True)
    is_merge_origin = Column(Boolean, default=False)

    sheet = relationship("ExcelSheet", back_populates="cells")
    merge = relationship("ExcelMerge")

    __table_args__ = (
        Index("ix_cell_sheet_row_col", "sheet_id", "row", "col", unique=True),
    )


# ── Domain model ────────────────────────────────────────────────────

class Employee(Base):
    __tablename__ = "employee"

    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=False)
    department = Column(String, nullable=True)
    position = Column(String, nullable=True)
    salary = Column(Float, nullable=True)

    # Excel origin tracking
    sheet_id = Column(Integer, ForeignKey("excel_sheet.id"), nullable=True)
    excel_row = Column(Integer, nullable=True)

    sheet = relationship("ExcelSheet")


# ── Column mapping (domain field → Excel column index) ──────────────

COLUMN_MAP: dict[str, int] = {}


def get_cell(session: Session, emp: Employee, field: str) -> ExcelCell | None:
    """Retrieve the ExcelCell linked to a specific Employee field."""
    col = COLUMN_MAP.get(field)
    if col is None:
        return None
    return (
        session.query(ExcelCell)
        .filter_by(sheet_id=emp.sheet_id, row=emp.excel_row, col=col)
        .first()
    )


# ── DB setup helper ─────────────────────────────────────────────────

def init_db(db_url: str = "sqlite:///excel_data.db"):
    engine = create_engine(db_url, echo=False)
    Base.metadata.create_all(engine)
    return sessionmaker(bind=engine)
