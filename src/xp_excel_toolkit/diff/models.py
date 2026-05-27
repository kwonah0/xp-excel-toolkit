"""Diff ORM models — generic cell-level diff schema.

Domain packages add their own diff tables on the same DiffBase so that a
single init_diff_db() call creates every registered table.
"""

from __future__ import annotations

from sqlalchemy import Text, create_engine
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, sessionmaker


class DiffBase(DeclarativeBase):
    pass


class DiffCell(DiffBase):
    """One row per changed/added/removed/moved cell."""
    __tablename__ = "diff_cell"

    id: Mapped[int] = mapped_column(primary_key=True)
    status: Mapped[str] = mapped_column(Text)  # added / removed / changed / moved
    sheet: Mapped[str | None] = mapped_column(Text)
    row: Mapped[int] = mapped_column()
    col: Mapped[int] = mapped_column()
    # Smart diff tracks original row numbers from both sides
    old_row: Mapped[int | None] = mapped_column(default=None)
    new_row: Mapped[int | None] = mapped_column(default=None)
    old_value: Mapped[str | None] = mapped_column(Text)
    new_value: Mapped[str | None] = mapped_column(Text)
    old_comment: Mapped[str | None] = mapped_column(Text)
    new_comment: Mapped[str | None] = mapped_column(Text)
    # JSON-encoded; populated when compare_style=True
    old_style: Mapped[str | None] = mapped_column(Text)
    new_style: Mapped[str | None] = mapped_column(Text)
    # e.g. "R1C1:R3C5"; populated when compare_merge=True
    old_merge_range: Mapped[str | None] = mapped_column(Text)
    new_merge_range: Mapped[str | None] = mapped_column(Text)
    # Formula strings; populated when source DB has cached_value
    old_formula: Mapped[str | None] = mapped_column(Text)
    new_formula: Mapped[str | None] = mapped_column(Text)


def init_diff_db(db_url: str) -> sessionmaker:
    """Create every table currently registered on DiffBase and return a sessionmaker.

    Import order matters: classes that subclass DiffBase must be imported
    before this call so SQLAlchemy has registered them on DiffBase.metadata.
    """
    engine = create_engine(db_url, echo=False)
    DiffBase.metadata.create_all(engine)
    return sessionmaker(bind=engine)
