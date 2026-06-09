"""xlrd-based parser for .xls files. Cells/merges only.

For a one-shot ``Cell → domain row`` workflow, run a builder (see
:mod:`xp_excel_toolkit.helpers` and the domain package) against the imported
ExcelCell rows after this returns.
"""

from __future__ import annotations

from pathlib import Path

import xlrd
from sqlalchemy.orm import Session

from xp_excel_toolkit.models import ExcelCell, ExcelMerge, ExcelSheet, ExcelWorkbook


def _extract_style_xls(book: xlrd.Book, xf_index: int) -> dict | None:
    """Extract style info from an xlrd XF record."""
    style: dict = {}

    try:
        xf = book.xf_list[xf_index]
    except (IndexError, AttributeError):
        return None

    try:
        bg_idx = xf.background.pattern_colour_index
        if bg_idx and bg_idx != 64:  # 64 = default/no fill
            colour = book.colour_map.get(bg_idx)
            if colour:
                style["bg_color"] = "#{:02X}{:02X}{:02X}".format(*colour)
    except AttributeError:
        pass

    try:
        font = book.font_list[xf.font_index]
        if font.bold:
            style["font_bold"] = True
        if font.struck_out:
            style["font_strike"] = True
        colour_idx = font.colour_index
        if colour_idx and colour_idx != 32767:  # 32767 = automatic
            colour = book.colour_map.get(colour_idx)
            if colour:
                style["font_color"] = "#{:02X}{:02X}{:02X}".format(*colour)
    except (AttributeError, IndexError):
        pass

    try:
        fmt_key = xf.format_key
        fmt_map = book.format_map
        if fmt_key in fmt_map:
            fmt_str = fmt_map[fmt_key].format_str
            if fmt_str and fmt_str != "General":
                style["number_format"] = fmt_str
    except AttributeError:
        pass

    border_names = {0: None, 1: "thin", 2: "medium", 5: "thick", 6: "double"}
    try:
        border = xf.border
        for attr, side in [
            ("left_line_style", "left"),
            ("right_line_style", "right"),
            ("top_line_style", "top"),
            ("bottom_line_style", "bottom"),
        ]:
            val = getattr(border, attr, 0)
            if val and val in border_names and border_names[val]:
                style[f"border_{side}"] = border_names[val]
    except AttributeError:
        pass

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
    :func:`xp_excel_toolkit.helpers.find_header_row_db` with its own header keyword
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

    # -- Merge ranges (convert 0-based exclusive -> 1-based inclusive) --
    merge_map: dict[tuple[int, int], tuple[int, int]] = {}
    merge_db: dict[str, ExcelMerge] = {}

    for rlo, rhi, clo, chi in sheet.merged_cells:
        origin = (rlo, clo)
        for r in range(rlo, rhi):
            for c in range(clo, chi):
                merge_map[(r, c)] = origin

        m = ExcelMerge(
            sheet_id=sheet_obj.id,
            min_row=rlo + 1,
            min_col=clo + 1,
            max_row=rhi,  # exclusive -> inclusive
            max_col=chi,
        )
        session.add(m)
        session.flush()
        merge_db[f"{rlo}:{clo}"] = m

    # -- Origin cell values used to fill merged regions --
    origin_values: dict[tuple[int, int], str | None] = {}
    for origin in set(merge_map.values()):
        or_, oc = origin
        val = sheet.cell_value(or_, oc)
        origin_values[origin] = str(val) if val != "" else None

    # -- Cells --
    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            cell = sheet.cell(r, c)

            if (r, c) in merge_map:
                origin = merge_map[(r, c)]
                raw_val = origin_values.get(origin)
                merge_key = merge_db[f"{origin[0]}:{origin[1]}"].id
                is_origin = (r, c) == origin
            else:
                raw_val = str(cell.value) if cell.value != "" else None
                merge_key = None
                is_origin = False

            style = (
                _extract_style_xls(book, cell.xf_index)
                if hasattr(cell, "xf_index") else None
            )

            session.add(ExcelCell(
                sheet_id=sheet_obj.id,
                row=r + 1,
                col=c + 1,
                raw_value=raw_val,
                style=style,
                merge_id=merge_key,
                is_merge_origin=is_origin,
            ))

    session.flush()

    # Header row is left to the domain builder. Caller may pass an
    # explicit 0-based ``header_row``; otherwise it stays NULL until
    # a builder sets it.
    if header_row is not None:
        sheet_obj.header_row = header_row + 1  # store as 1-based

    session.flush()
    return sheet_obj
