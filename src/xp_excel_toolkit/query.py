"""Read-only DB-side helpers.

These read from the ExcelCell / ExcelSheet tables and return plain
Python structures — they never insert into domain tables.
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from typing import Any, Literal

from sqlalchemy import select
from sqlalchemy.orm import Session

from xp_excel_toolkit.models import ExcelCell

HeaderMatch = Literal["any", "all"]


def find_header_row_db(
    session: Session,
    sheet_id: int,
    expected_headers: Iterable[str],
    *,
    match: HeaderMatch = "any",
    max_scan: int = 30,
) -> int | None:
    """Auto-detect the header row from ExcelCell rows in the DB.

    Args:
        expected_headers: Header strings the row is expected to contain.
            Any iterable of strings (list, tuple, set, ``dict.keys()``).
            At least one non-empty string is required.
        match: ``"any"`` (default) returns the first row containing at
            least one of ``expected_headers``; ``"all"`` requires every
            expected header to be present.
        max_scan: Stop scanning after this row number.

    Returns:
        1-based row index, or ``None`` if no match is found within
        ``max_scan``.

    Raises:
        ValueError: if ``expected_headers`` is empty or contains only
            blank strings.
    """
    expected = {s.strip() for s in expected_headers if s and s.strip()}
    if not expected:
        raise ValueError(
            "expected_headers must contain at least one non-empty string"
        )

    rows = session.execute(
        select(ExcelCell.row, ExcelCell.raw_value)
        .where(ExcelCell.sheet_id == sheet_id, ExcelCell.row <= max_scan)
        .order_by(ExcelCell.row, ExcelCell.col)
    ).all()

    by_row: dict[int, set[str]] = {}
    for r, v in rows:
        if v and v.strip():
            by_row.setdefault(r, set()).add(v.strip())

    for r in sorted(by_row):
        values = by_row[r]
        if match == "all":
            if expected.issubset(values):
                return r
        else:  # "any"
            if values & expected:
                return r
    return None


def iter_rows_by_header(
    session: Session,
    sheet_id: int,
    header_row: int,
    *,
    headers: Iterable[str] | None = None,
) -> Iterator[tuple[int, dict[str, Any]]]:
    """Yield ``(row_idx, {header_str: value})`` for every non-empty data row.

    xp_excel_toolkit reads ``header_row`` internally to build a ``col_idx → header_str``
    mapping, then projects each subsequent row through it. The caller never
    needs to know column indices — only header strings.

    Layout assumption: single-row header, flat columns. Multi-row /
    compound / pivot layouts must use raw ``ExcelCell`` SELECTs.

    Args:
        session: SQLAlchemy session.
        sheet_id: ``ExcelSheet`` id.
        header_row: 1-based index of the header row.
        headers: Optional whitelist of header strings. If given, only
            these columns appear in the output dicts (and the SQL fetch
            is narrowed accordingly). If ``None``, every column whose
            header cell holds a non-empty string is included.

    Yields:
        ``(row_idx, cells)`` per data row. ``cells`` is a dict keyed by
        the header string. Rows where every mapped value is ``None`` are
        skipped.
    """
    header_cells = session.execute(
        select(ExcelCell.col, ExcelCell.raw_value)
        .where(ExcelCell.sheet_id == sheet_id, ExcelCell.row == header_row)
    ).all()

    col_to_header: dict[int, str] = {}
    for col, v in header_cells:
        if v and v.strip():
            col_to_header[col] = v.strip()

    if headers is not None:
        wanted = {s.strip() for s in headers if s and s.strip()}
        col_to_header = {c: h for c, h in col_to_header.items() if h in wanted}

    if not col_to_header:
        return

    data_cells = session.execute(
        select(ExcelCell.row, ExcelCell.col, ExcelCell.raw_value)
        .where(
            ExcelCell.sheet_id == sheet_id,
            ExcelCell.row > header_row,
            ExcelCell.col.in_(col_to_header.keys()),
        )
        .order_by(ExcelCell.row, ExcelCell.col)
    ).all()

    rows: dict[int, dict[str, Any]] = {}
    for r, c, v in data_cells:
        rows.setdefault(r, {})[col_to_header[c]] = v

    for row_idx in sorted(rows):
        cells = rows[row_idx]
        if any(v is not None for v in cells.values()):
            yield row_idx, cells
