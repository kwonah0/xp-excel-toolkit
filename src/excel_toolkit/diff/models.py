"""Diff ORM models and result container."""

from __future__ import annotations

from sqlalchemy import Text, create_engine
from sqlalchemy.orm import (
    DeclarativeBase, Mapped, mapped_column, sessionmaker,
)


# Fields to compare (imported from domain models — kept as aliases for convenience)
from excel_toolkit.domain_models import REGMAP_FIELDS as _REG_FIELDS  # noqa: E402
from excel_toolkit.domain_models import MEMMAP_FIELDS as _MEMMAP_FIELDS  # noqa: E402


# ── Diff DB models ────────────────────────────────────────────────

class DiffBase(DeclarativeBase):
    pass


class DiffMeta(DiffBase):
    __tablename__ = "diff_meta"

    id: Mapped[int] = mapped_column(primary_key=True)
    old_path: Mapped[str] = mapped_column(Text)
    new_path: Mapped[str] = mapped_column(Text)
    created_at: Mapped[str] = mapped_column(Text)
    added_regs: Mapped[int] = mapped_column(default=0)
    removed_regs: Mapped[int] = mapped_column(default=0)
    changed_regs: Mapped[int] = mapped_column(default=0)
    added_memmap: Mapped[int] = mapped_column(default=0)
    removed_memmap: Mapped[int] = mapped_column(default=0)
    changed_memmap: Mapped[int] = mapped_column(default=0)


class DiffRegister(DiffBase):
    """One row per register: old/new values side-by-side."""
    __tablename__ = "diff_register"

    id: Mapped[int] = mapped_column(primary_key=True)
    status: Mapped[str] = mapped_column(Text)  # added / removed / changed
    sheet: Mapped[str | None] = mapped_column(Text)
    # -- old values --
    old_type: Mapped[str | None] = mapped_column(Text)
    old_indx: Mapped[str | None] = mapped_column(Text)
    old_page: Mapped[str | None] = mapped_column(Text)
    old_para: Mapped[str | None] = mapped_column(Text)
    old_name: Mapped[str | None] = mapped_column(Text)
    old_d7: Mapped[str | None] = mapped_column(Text)
    old_d6: Mapped[str | None] = mapped_column(Text)
    old_d5: Mapped[str | None] = mapped_column(Text)
    old_d4: Mapped[str | None] = mapped_column(Text)
    old_d3: Mapped[str | None] = mapped_column(Text)
    old_d2: Mapped[str | None] = mapped_column(Text)
    old_d1: Mapped[str | None] = mapped_column(Text)
    old_d0: Mapped[str | None] = mapped_column(Text)
    old_init: Mapped[str | None] = mapped_column(Text)
    # -- new values --
    new_type: Mapped[str | None] = mapped_column(Text)
    new_indx: Mapped[str | None] = mapped_column(Text)
    new_page: Mapped[str | None] = mapped_column(Text)
    new_para: Mapped[str | None] = mapped_column(Text)
    new_name: Mapped[str | None] = mapped_column(Text)
    new_d7: Mapped[str | None] = mapped_column(Text)
    new_d6: Mapped[str | None] = mapped_column(Text)
    new_d5: Mapped[str | None] = mapped_column(Text)
    new_d4: Mapped[str | None] = mapped_column(Text)
    new_d3: Mapped[str | None] = mapped_column(Text)
    new_d2: Mapped[str | None] = mapped_column(Text)
    new_d1: Mapped[str | None] = mapped_column(Text)
    new_d0: Mapped[str | None] = mapped_column(Text)
    new_init: Mapped[str | None] = mapped_column(Text)


class DiffMemmap(DiffBase):
    __tablename__ = "diff_memmap"

    id: Mapped[int] = mapped_column(primary_key=True)
    status: Mapped[str] = mapped_column(Text)  # added / removed / changed
    old_baseaddr: Mapped[str | None] = mapped_column(Text)
    old_group: Mapped[str | None] = mapped_column(Text)
    old_midgroup: Mapped[str | None] = mapped_column(Text)
    old_comment: Mapped[str | None] = mapped_column(Text)
    old_special: Mapped[str | None] = mapped_column(Text)
    new_baseaddr: Mapped[str | None] = mapped_column(Text)
    new_group: Mapped[str | None] = mapped_column(Text)
    new_midgroup: Mapped[str | None] = mapped_column(Text)
    new_comment: Mapped[str | None] = mapped_column(Text)
    new_special: Mapped[str | None] = mapped_column(Text)


class DiffCell(DiffBase):
    """One row per changed/added/removed cell."""
    __tablename__ = "diff_cell"

    id: Mapped[int] = mapped_column(primary_key=True)
    status: Mapped[str] = mapped_column(Text)  # added / removed / changed
    sheet: Mapped[str | None] = mapped_column(Text)
    row: Mapped[int] = mapped_column()
    col: Mapped[int] = mapped_column()
    # For smart diff: track original row numbers from both sides
    old_row: Mapped[int | None] = mapped_column(default=None)
    new_row: Mapped[int | None] = mapped_column(default=None)
    old_value: Mapped[str | None] = mapped_column(Text)
    new_value: Mapped[str | None] = mapped_column(Text)
    old_comment: Mapped[str | None] = mapped_column(Text)
    new_comment: Mapped[str | None] = mapped_column(Text)
    # Style diff (JSON string, populated when compare_style=True)
    old_style: Mapped[str | None] = mapped_column(Text)
    new_style: Mapped[str | None] = mapped_column(Text)
    # Merge range diff (e.g. "R1C1:R3C5", populated when compare_merge=True)
    old_merge_range: Mapped[str | None] = mapped_column(Text)
    new_merge_range: Mapped[str | None] = mapped_column(Text)
    # Formula strings (populated when source DB has cached_value, i.e. --with-formulas import)
    old_formula: Mapped[str | None] = mapped_column(Text)
    new_formula: Mapped[str | None] = mapped_column(Text)


def init_diff_db(db_url: str) -> sessionmaker:
    engine = create_engine(db_url, echo=False)
    DiffBase.metadata.create_all(engine)
    return sessionmaker(bind=engine)


# ── DiffResult (holds ORM objects directly) ────────────────────────

class DiffResult:
    """Full diff between two databases — holds ORM objects directly."""

    def __init__(self) -> None:
        self.registers: list[DiffRegister] = []
        self.memmap: list[DiffMemmap] = []
        self.cells: list[DiffCell] = []

    def _filter_regs(self, status: str) -> list[DiffRegister]:
        return [r for r in self.registers if r.status == status]

    def _filter_mm(self, status: str) -> list[DiffMemmap]:
        return [m for m in self.memmap if m.status == status]
