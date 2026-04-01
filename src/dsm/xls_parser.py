"""xlrd-based parser for .xls files."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import xlrd
from sqlalchemy.orm import Session

from dsm.models import ExcelCell, ExcelMerge, ExcelSheet, ExcelWorkbook


def _extract_style_xls(book: xlrd.Book, xf_index: int) -> dict | None:
    """Extract style info from xlrd XF record."""
    style: dict = {}

    try:
        xf = book.xf_list[xf_index]
    except (IndexError, AttributeError):
        return None

    # Background color
    try:
        bg_idx = xf.background.pattern_colour_index
        if bg_idx and bg_idx != 64:  # 64 = default/no fill
            colour = book.colour_map.get(bg_idx)
            if colour:
                style["bg_color"] = "#{:02X}{:02X}{:02X}".format(*colour)
    except AttributeError:
        pass

    # Font
    try:
        font = book.font_list[xf.font_index]
        if font.bold:
            style["font_bold"] = True
        colour_idx = font.colour_index
        if colour_idx and colour_idx != 32767:  # 32767 = automatic
            colour = book.colour_map.get(colour_idx)
            if colour:
                style["font_color"] = "#{:02X}{:02X}{:02X}".format(*colour)
    except (AttributeError, IndexError):
        pass

    # Number format
    try:
        fmt_key = xf.format_key
        fmt_map = book.format_map
        if fmt_key in fmt_map:
            fmt_str = fmt_map[fmt_key].format_str
            if fmt_str and fmt_str != "General":
                style["number_format"] = fmt_str
    except AttributeError:
        pass

    # Border
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


def _build_merge_map_xls(sheet: xlrd.sheet.Sheet) -> dict[tuple[int, int], tuple[int, int]]:
    """Build mapping: (row, col) -> (origin_row, origin_col) for merged cells."""
    merge_map: dict[tuple[int, int], tuple[int, int]] = {}
    for rlo, rhi, clo, chi in sheet.merged_cells:
        origin = (rlo, clo)
        for r in range(rlo, rhi):
            for c in range(clo, chi):
                merge_map[(r, c)] = origin
    return merge_map


def find_header_row_xls(sheet: xlrd.sheet.Sheet, min_cols: int = 3) -> int | None:
    """Auto-detect header row (0-based) in xls sheet."""
    for r in range(min(30, sheet.nrows)):
        str_count = 0
        for c in range(sheet.ncols):
            val = sheet.cell_value(r, c)
            if isinstance(val, str) and val.strip():
                str_count += 1
        if str_count >= min_cols:
            return r
    return None


def import_xls(
    session: Session,
    path: str | Path,
    sheet_index: int = 0,
    header_row: int | None = None,
    field_map: dict[str, str] | None = None,
    domain_cls: type | None = None,
    column_map: dict[str, int] | None = None,
) -> ExcelSheet:
    """
    Import a .xls file into the DB.

    Args:
        session: SQLAlchemy session
        path: Path to .xls file
        sheet_index: Sheet index (default: 0)
        header_row: 0-based header row index (auto-detected if None)
        field_map: Mapping of Excel column header -> domain field name
        domain_cls: ORM class to instantiate for each data row
        column_map: Optional dict to populate with {field_name: col_index}

    Note: xlrd uses 0-based row/col indices. We store them as 1-based
          in the DB for consistency with openpyxl.
    """
    path = Path(path)
    blob = path.read_bytes()

    book = xlrd.open_workbook(str(path), formatting_info=True)
    sheet = book.sheet_by_index(sheet_index)

    # Create workbook record
    wb_obj = ExcelWorkbook(filename=path.name, blob=blob)
    session.add(wb_obj)
    session.flush()

    # Create sheet record
    sheet_obj = ExcelSheet(workbook_id=wb_obj.id, name=sheet.name)
    session.add(sheet_obj)
    session.flush()

    # -- Merge ranges (convert 0-based exclusive -> 1-based inclusive) --
    merge_map = _build_merge_map_xls(sheet)
    merge_db: dict[str, ExcelMerge] = {}

    for rlo, rhi, clo, chi in sheet.merged_cells:
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

    # -- Cells --
    origin_values: dict[tuple[int, int], str | None] = {}
    for (r, c), (or_, oc) in merge_map.items():
        if (or_, oc) not in origin_values:
            val = sheet.cell_value(or_, oc)
            origin_values[(or_, oc)] = str(val) if val != "" else None

    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            cell = sheet.cell(r, c)

            if (r, c) in merge_map:
                origin = merge_map[(r, c)]
                raw_val = origin_values.get(origin)
            else:
                raw_val = str(cell.value) if cell.value != "" else None

            merge_key = None
            is_origin = False
            if (r, c) in merge_map:
                or_, oc = merge_map[(r, c)]
                merge_key = merge_db[f"{or_}:{oc}"].id
                is_origin = (r == or_ and c == oc)

            style = _extract_style_xls(book, cell.xf_index) if hasattr(cell, 'xf_index') else None

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

    # -- Header detection + domain objects --
    if header_row is None:
        header_row_0 = find_header_row_xls(sheet)
    else:
        header_row_0 = header_row

    if header_row_0 is not None:
        sheet_obj.header_row = header_row_0 + 1  # store as 1-based

    if header_row_0 is not None and field_map and domain_cls:
        headers: dict[int, str] = {}
        for c in range(sheet.ncols):
            val = sheet.cell_value(header_row_0, c)
            if isinstance(val, str) and val.strip():
                headers[c] = val.strip()

        col_to_field: dict[int, str] = {}
        for col_0, header_text in headers.items():
            if header_text in field_map:
                domain_field = field_map[header_text]
                col_to_field[col_0] = domain_field
                if column_map is not None:
                    column_map[domain_field] = col_0 + 1  # 1-based

        for r in range(header_row_0 + 1, sheet.nrows):
            row_data: dict[str, Any] = {}
            has_data = False
            for col_0, field_name in col_to_field.items():
                val = sheet.cell_value(r, col_0)
                if (val == "" or val is None) and (r, col_0) in merge_map:
                    or_, oc = merge_map[(r, col_0)]
                    val = sheet.cell_value(or_, oc)
                if val != "" and val is not None:
                    has_data = True
                    row_data[field_name] = str(val)
                else:
                    row_data[field_name] = None

            if has_data:
                try:
                    with session.begin_nested():
                        obj = domain_cls(
                            sheet_id=sheet_obj.id,
                            excel_row=r + 1,
                            **row_data,
                        )
                        session.add(obj)
                        session.flush()
                except Exception:
                    pass

    session.flush()
    return sheet_obj
