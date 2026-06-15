"""SQLAlchemy ORM models for Excel metadata + generic mechanisms.

This module owns the schema for Excel structure (workbook/sheet/cell/merge),
sheet-config registry, and an audit-trigger mechanism. Domain packages
attach their own models to the shared :class:`Base` and register audit
targets via :func:`register_audit_target`.
"""

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
    cached_value: Mapped[str | None] = mapped_column(Text)  # formula result (data_only)
    style: Mapped[dict | None] = mapped_column(JSON)  # {bg_color, font_bold, number_format, ...}
    comment: Mapped[str | None] = mapped_column(Text)  # Excel cell note/comment text
    formula_type: Mapped[str | None] = mapped_column(Text)  # "array" or "dataTable"
    formula_ref: Mapped[str | None] = mapped_column(Text)   # e.g. "B1:B10"
    merge_id: Mapped[int | None] = mapped_column(ForeignKey("excel_merge.id"))
    is_merge_origin: Mapped[bool] = mapped_column(default=False)

    sheet: Mapped[ExcelSheet] = relationship(back_populates="cells")
    merge: Mapped[ExcelMerge | None] = relationship()

    __table_args__ = (
        Index("ix_cell_sheet_row_col", "sheet_id", "row", "col", unique=True),
    )


# ── Sheet config (stored in DB) ─────────────────────────────────────

class SheetConfigEntry(Base):
    """Per-sheet import configuration stored in the database."""
    __tablename__ = "sheet_config"

    id: Mapped[int] = mapped_column(primary_key=True)
    pattern: Mapped[str] = mapped_column(Text)              # fnmatch pattern, e.g. "level2_*"
    domain_type: Mapped[str | None] = mapped_column(Text)   # "register" or "memorymap_entry"
    field_map_json: Mapped[str | None] = mapped_column(Text)  # JSON: {"TYPE": "type", ...}
    header_row: Mapped[int | None] = mapped_column(default=None)
    parser_func_ref: Mapped[str | None] = mapped_column(Text)  # "module:func" for custom parsers


# ── Audit log (change tracking) ────────────────────────────────────

class ChangeLog(Base):
    """Audit log for domain model changes (UPDATE / DELETE)."""
    __tablename__ = "change_log"

    id: Mapped[int] = mapped_column(primary_key=True)
    timestamp: Mapped[str] = mapped_column(Text)       # ISO 8601 (UTC)
    table_name: Mapped[str] = mapped_column(Text)
    row_id: Mapped[int]
    column_name: Mapped[str] = mapped_column(Text)
    old_value: Mapped[str | None] = mapped_column(Text)
    new_value: Mapped[str | None] = mapped_column(Text)
    operation: Mapped[str] = mapped_column(Text)        # "UPDATE" / "DELETE"


# Registry: table_name -> list of column names to audit
AUDIT_TARGETS: dict[str, list[str]] = {}


def register_audit_target(table_name: str, columns: list[str]) -> None:
    """Register a table and its columns for audit trigger creation."""
    AUDIT_TARGETS[table_name] = columns


def create_audit_triggers(engine) -> None:
    """Create SQLite triggers for all registered audit targets.

    Public so callers that build their own engine (instead of going through
    ``init_db``) can install the change-log triggers themselves.
    """
    from sqlalchemy import text

    with engine.connect() as conn:
        for table_name, columns in AUDIT_TARGETS.items():
            # UPDATE triggers: one per column (quote identifiers for SQL reserved words)
            for col in columns:
                trigger_name = f"trg_{table_name}_update_{col}"
                conn.execute(text(f"""
                    CREATE TRIGGER IF NOT EXISTS [{trigger_name}]
                    AFTER UPDATE OF [{col}] ON [{table_name}]
                    WHEN OLD.[{col}] IS NOT NEW.[{col}]
                    BEGIN
                        INSERT INTO change_log
                            (timestamp, table_name, row_id, column_name,
                             old_value, new_value, operation)
                        VALUES
                            (datetime('now'), '{table_name}', NEW.id, '{col}',
                             OLD.[{col}], NEW.[{col}], 'UPDATE');
                    END;
                """))

            # DELETE trigger: one per table, logs all columns
            trigger_name = f"trg_{table_name}_delete"
            inserts = "\n".join(
                f"""INSERT INTO change_log
                    (timestamp, table_name, row_id, column_name,
                     old_value, new_value, operation)
                VALUES
                    (datetime('now'), '{table_name}', OLD.id, '{col}',
                     OLD.[{col}], NULL, 'DELETE');"""
                for col in columns
            )
            conn.execute(text(f"""
                CREATE TRIGGER IF NOT EXISTS [{trigger_name}]
                AFTER DELETE ON [{table_name}]
                BEGIN
                    {inserts}
                END;
            """))

        conn.commit()


# ── DB setup helper ─────────────────────────────────────────────────

def init_db(db_url: str = "sqlite:///excel_data.db"):
    engine = create_engine(db_url, echo=False)
    Base.metadata.create_all(engine)
    create_audit_triggers(engine)
    return sessionmaker(bind=engine)
