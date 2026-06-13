"""xlrd-based parser for .xls files. Cells/merges only.

For a one-shot ``Cell → domain row`` workflow, run a builder (see
:mod:`xp_excel_toolkit.query` and the domain package) against the imported
ExcelCell rows after this returns.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import xlrd
from sqlalchemy import insert
from sqlalchemy.orm import Session

from xp_excel_toolkit.ingest.xlsx import BULK_CHUNK
from xp_excel_toolkit.merge import MergeResolver
from xp_excel_toolkit.models import ExcelCell, ExcelMerge, ExcelSheet, ExcelWorkbook

_NO_FILL_COLOUR_INDEX = 64
_AUTOMATIC_FONT_COLOUR_INDEX = 32767
_BORDER_NAMES = {1: "thin", 2: "medium", 5: "thick", 6: "double"}


def _extract_style_xls(book: xlrd.Book, xf_index: int) -> dict[str, Any] | None:
    """Extract style info from an xlrd XF record."""
    xf_list = getattr(book, "xf_list", [])
    if not 0 <= xf_index < len(xf_list):
        return None
    xf = xf_list[xf_index]
    colour_map = getattr(book, "colour_map", {})
    style: dict[str, Any] = {}

    bg_idx = getattr(getattr(xf, "background", None), "pattern_colour_index", None)
    if bg_idx and bg_idx != _NO_FILL_COLOUR_INDEX:
        colour = colour_map.get(bg_idx)
        if colour:
            style["bg_color"] = "#{:02X}{:02X}{:02X}".format(*colour)

    font_list = getattr(book, "font_list", [])
    font_index = getattr(xf, "font_index", None)
    if font_index is not None and 0 <= font_index < len(font_list):
        font = font_list[font_index]
        if getattr(font, "bold", False):
            style["font_bold"] = True
        if getattr(font, "struck_out", False):
            style["font_strike"] = True
        colour_idx = getattr(font, "colour_index", None)
        if colour_idx and colour_idx != _AUTOMATIC_FONT_COLOUR_INDEX:
            colour = colour_map.get(colour_idx)
            if colour:
                style["font_color"] = "#{:02X}{:02X}{:02X}".format(*colour)

    fmt = getattr(book, "format_map", {}).get(getattr(xf, "format_key", None))
    fmt_str = getattr(fmt, "format_str", None)
    if fmt_str and fmt_str != "General":
        style["number_format"] = fmt_str

    border = getattr(xf, "border", None)
    for attr, side in (
        ("left_line_style", "left"),
        ("right_line_style", "right"),
        ("top_line_style", "top"),
        ("bottom_line_style", "bottom"),
    ):
        name = _BORDER_NAMES.get(getattr(border, attr, 0))
        if name:
            style[f"border_{side}"] = name

    return style if style else None


def import_xls(
    session: Session,
    path: str | Path,
    sheet_index: int = 0,
    header_row: int | None = None,
) -> ExcelSheet:
    """Import a .xls file into the DB (cells/merges only).

    ``header_row`` is only stored if the caller passes it. xp_excel_toolkit itself
    does not try to guess; the domain builder detects the header via
    :func:`xp_excel_toolkit.query.find_header_row_db` with its own header keyword
    list and updates ``ExcelSheet.header_row``.

    xlrd uses 0-based row/col indices. We store them as 1-based in the
    DB for consistency with openpyxl.
    """
    path = Path(path)
    blob = path.read_bytes()

    book = xlrd.open_workbook(str(path), formatting_info=True)
    sheet = book.sheet_by_index(sheet_index)

    wb_obj = ExcelWorkbook(filename=path.name, blob=blob)
    session.add(wb_obj)
    session.flush()

    sheet_obj = ExcelSheet(workbook_id=wb_obj.id, name=sheet.name)
    session.add(sheet_obj)
    session.flush()

    sid = sheet_obj.id

    # -- Merge ranges (convert 0-based exclusive -> 1-based inclusive) --
    bounds: list[tuple[int, int, int, int]] = []
    origin_values: dict[tuple[int, int], str | None] = {}
    for rlo, rhi, clo, chi in sheet.merged_cells:
        bounds.append((rlo + 1, clo + 1, rhi, chi))
        val = sheet.cell_value(rlo, clo)
        origin_values[(rlo + 1, clo + 1)] = str(val) if val != "" else None

    merger = MergeResolver.from_bounds(bounds, origin_values)

    merge_dicts = [
        {
            "sheet_id": sid,
            "min_row": mr.min_row,
            "min_col": mr.min_col,
            "max_row": mr.max_row,
            "max_col": mr.max_col,
        }
        for mr in merger.ranges
    ]

    merge_id_map: dict[str, int] = {}  # "min_row:min_col" -> merge.id
    if merge_dicts:
        for i in range(0, len(merge_dicts), BULK_CHUNK):
            session.execute(insert(ExcelMerge), merge_dicts[i:i + BULK_CHUNK])
        session.flush()

        for mid, mrow, mcol in (
            session.query(ExcelMerge.id, ExcelMerge.min_row, ExcelMerge.min_col)
            .filter(ExcelMerge.sheet_id == sid)
            .all()
        ):
            merge_id_map[f"{mrow}:{mcol}"] = mid

    # -- Cells (Core bulk insert) --
    cell_dicts: list[dict[str, Any]] = []
    for r0 in range(sheet.nrows):
        for c0 in range(sheet.ncols):
            cell = sheet.cell(r0, c0)
            r, c = r0 + 1, c0 + 1

            if merger.is_merged(r, c):
                raw_val = merger.get_value(r, c)
            else:
                raw_val = str(cell.value) if cell.value != "" else None

            merge_key = None
            origin = merger.get_origin(r, c)
            if origin:
                merge_key = merge_id_map.get(f"{origin[0]}:{origin[1]}")

            style = (
                _extract_style_xls(book, cell.xf_index)
                if cell.xf_index is not None else None
            )

            cell_dicts.append({
                "sheet_id": sid,
                "row": r,
                "col": c,
                "raw_value": raw_val,
                "style": style,
                "merge_id": merge_key,
                "is_merge_origin": merger.is_origin(r, c),
            })

    if cell_dicts:
        for i in range(0, len(cell_dicts), BULK_CHUNK):
            session.execute(insert(ExcelCell), cell_dicts[i:i + BULK_CHUNK])
        session.flush()

    # Header row is left to the domain builder. Caller may pass an
    # explicit 0-based ``header_row``; otherwise it stays NULL until
    # a builder sets it.
    if header_row is not None:
        sheet_obj.header_row = header_row + 1  # store as 1-based

    session.flush()
    return sheet_obj
