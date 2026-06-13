"""Merge cell resolver — works from openpyxl worksheet, DB records, or raw bounds."""

from __future__ import annotations

from collections.abc import Iterable
from typing import TYPE_CHECKING, NamedTuple, Protocol

from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

from xp_excel_toolkit.models import ExcelCell, ExcelMerge

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet
    from sqlalchemy.orm import Session


class RangeLike(Protocol):
    min_row: int
    min_col: int
    max_row: int
    max_col: int


class MergeBounds(NamedTuple):
    min_row: int
    min_col: int
    max_row: int
    max_col: int


def _build_merge_map(ranges: Iterable[RangeLike]) -> dict[tuple[int, int], tuple[int, int]]:
    merge_map: dict[tuple[int, int], tuple[int, int]] = {}
    for mr in ranges:
        origin = (mr.min_row, mr.min_col)
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                merge_map[(r, c)] = origin
    return merge_map


class MergeResolver:
    """Resolve merged cell values for a sheet.

    Construction paths:
      - MergeResolver.from_worksheet(ws)            — openpyxl worksheet (import time)
      - MergeResolver.from_db(session, sheet_id)    — ExcelMerge + ExcelCell records
      - MergeResolver.from_bounds(bounds, values)   — raw 1-based inclusive bounds
    """

    def __init__(
        self,
        ranges: list[RangeLike],
        merge_map: dict[tuple[int, int], tuple[int, int]],
        origin_values: dict[tuple[int, int], str | None],
    ) -> None:
        self.ranges = ranges
        self._merge_map = merge_map
        self._origin_values = origin_values

    @classmethod
    def from_worksheet(cls, ws: Worksheet) -> MergeResolver:
        """Build from an openpyxl worksheet (import time)."""
        ranges = list(ws.merged_cells.ranges)
        merge_map = _build_merge_map(ranges)

        origin_values: dict[tuple[int, int], str | None] = {}
        for mr in ranges:
            origin = (mr.min_row, mr.min_col)
            val = ws.cell(row=mr.min_row, column=mr.min_col).value
            if val is None:
                origin_values[origin] = None
            elif isinstance(val, ArrayFormula):
                origin_values[origin] = val.text if val.text else str(val.ref)
            elif isinstance(val, DataTableFormula):
                origin_values[origin] = None
            else:
                origin_values[origin] = str(val)

        return cls(ranges, merge_map, origin_values)

    @classmethod
    def from_db(cls, session: Session, sheet_id: int) -> MergeResolver:
        """Build from DB records (no xlsx file needed)."""
        ranges = (
            session.query(ExcelMerge)
            .filter(ExcelMerge.sheet_id == sheet_id)
            .all()
        )
        merge_map = _build_merge_map(ranges)

        origin_values: dict[tuple[int, int], str | None] = {}
        if ranges:
            origin_cells = (
                session.query(ExcelCell.row, ExcelCell.col, ExcelCell.raw_value)
                .filter(
                    ExcelCell.sheet_id == sheet_id,
                    ExcelCell.is_merge_origin.is_(True),
                )
                .all()
            )
            for row, col, raw_value in origin_cells:
                if (row, col) in merge_map:
                    origin_values[(row, col)] = raw_value

        return cls(ranges, merge_map, origin_values)

    @classmethod
    def from_bounds(
        cls,
        bounds: Iterable[tuple[int, int, int, int]],
        origin_values: dict[tuple[int, int], str | None],
    ) -> MergeResolver:
        """Build from raw 1-based inclusive (min_row, min_col, max_row, max_col) bounds."""
        ranges = [MergeBounds(*b) for b in bounds]
        return cls(ranges, _build_merge_map(ranges), dict(origin_values))

    def is_merged(self, row: int, col: int) -> bool:
        return (row, col) in self._merge_map

    def is_origin(self, row: int, col: int) -> bool:
        origin = self._merge_map.get((row, col))
        return origin == (row, col) if origin else False

    def get_origin(self, row: int, col: int) -> tuple[int, int] | None:
        return self._merge_map.get((row, col))

    def get_value(self, row: int, col: int) -> str | None:
        """Return the origin value for a merged cell."""
        origin = self._merge_map.get((row, col))
        return self._origin_values.get(origin) if origin else None
