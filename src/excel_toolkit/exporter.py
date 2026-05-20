"""Export Excel data from the ORM model back to .xlsx files."""

from __future__ import annotations

import fnmatch
import io
from collections.abc import Callable
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.comments import Comment
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
from openpyxl.worksheet.worksheet import Worksheet
from sqlalchemy.orm import Session

from excel_toolkit.models import ExcelCell, ExcelMerge, ExcelSheet, ExcelWorkbook
from excel_toolkit.domain_models import (
    REGMAP_FIELD_MAP, MEMMAP_FIELD_MAP,
    Register, MemoryMapEntry, OverviewEntry,
)


# ── Cell writing helpers ─────────────────────────────────────────────

def _write_cell(ws: Worksheet, row: int, col: int, val) -> None:
    """Write a value to a cell, handling merged cells."""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for mr in ws.merged_cells.ranges:
            if (mr.min_row <= row <= mr.max_row
                    and mr.min_col <= col <= mr.max_col):
                ws.cell(row=mr.min_row, column=mr.min_col).value = val
                break
    else:
        cell.value = val


def _build_column_map(
    ws: Worksheet,
    header_row: int,
    field_map: dict[str, str],
) -> dict[str, int]:
    """Build field_name → column_index mapping from header row."""
    col_map: dict[str, int] = {}
    for cell in ws[header_row]:
        if cell.value and isinstance(cell.value, str):
            header = cell.value.strip()
            if header in field_map:
                col_map[field_map[header]] = cell.column
    return col_map


# ── Export: domain models → xlsx ─────────────────────────────────────

def _export_flat_table(
    session: Session,
    ws: Worksheet,
    sheet_obj: ExcelSheet,
    domain_cls: type,
    field_map: dict[str, str],
) -> int:
    """Export Register or MemoryMapEntry rows back to worksheet cells.

    Returns number of rows written.
    """
    if not sheet_obj.header_row:
        return 0

    col_map = _build_column_map(ws, sheet_obj.header_row, field_map)
    if not col_map:
        return 0

    rows = (
        session.query(domain_cls)
        .filter_by(sheet_id=sheet_obj.id)
        .all()
    )

    for row_obj in rows:
        for field_name, col_idx in col_map.items():
            val = getattr(row_obj, field_name, None)
            _write_cell(ws, row_obj.excel_row, col_idx, val)

    return len(rows)


def _export_overview(
    session: Session,
    ws: Worksheet,
    sheet_obj: ExcelSheet,
) -> int:
    """Export OverviewEntry rows back to worksheet cells.

    Returns number of rows written.
    """
    entries = (
        session.query(OverviewEntry)
        .filter_by(sheet_id=sheet_obj.id)
        .order_by(OverviewEntry.excel_row)
        .all()
    )

    for entry in entries:
        row = entry.excel_row
        if entry.is_category:
            _write_cell(ws, row, 1, f"#{entry.category}")
        elif entry.is_commented:
            _write_cell(ws, row, 1, f"#{entry.key}")
        else:
            _write_cell(ws, row, 1, entry.key)

        _write_cell(ws, row, 2, entry.value)
        _write_cell(ws, row, 3, entry.comment)

    return len(entries)


def export_xlsx(
    session: Session,
    output_path: str | Path,
    on_progress: Callable[[str], None] | None = None,
) -> Path:
    """Export DB domain models back to xlsx, preserving original formatting.

    Loads the original BLOB and overwrites only domain model cells.
    All formatting, merges, styles, and non-domain sheets are preserved.
    """
    output_path = Path(output_path)

    wb_obj = session.query(ExcelWorkbook).first()
    if not wb_obj or not wb_obj.blob:
        raise ValueError("No workbook found in DB or BLOB is missing")

    wb = openpyxl.load_workbook(io.BytesIO(wb_obj.blob))

    # Sheet pattern → (domain_cls, field_map | None)
    _SHEET_HANDLERS = [
        ("level2_*", Register, REGMAP_FIELD_MAP),
        ("memorymap", MemoryMapEntry, MEMMAP_FIELD_MAP),
        ("Overview", OverviewEntry, None),
    ]

    total = 0
    for sheet_obj in wb_obj.sheets:
        if sheet_obj.name not in wb.sheetnames:
            continue
        ws = wb[sheet_obj.name]

        for pattern, domain_cls, field_map in _SHEET_HANDLERS:
            if not fnmatch.fnmatch(sheet_obj.name.lower(), pattern.lower()):
                continue

            if domain_cls is OverviewEntry:
                count = _export_overview(session, ws, sheet_obj)
            else:
                count = _export_flat_table(
                    session, ws, sheet_obj, domain_cls, field_map,
                )

            if count and on_progress:
                on_progress(f"  {sheet_obj.name}: {count} rows")
            total += count
            break  # matched, no need to try other patterns

    wb.save(output_path)
    wb.close()

    if on_progress:
        on_progress(f"Exported {total} domain rows to {output_path}")

    return output_path


# ── Legacy / utility exports ─────────────────────────────────────────

# Keep export_regmap_xlsx as alias for backward compatibility
def export_regmap_xlsx(
    session: Session,
    workbook_id: int,
    output_path: str | Path,
    column_map: dict[str, int] | None = None,
) -> Path:
    """Write modified Register data back to .xlsx (legacy API)."""
    return export_xlsx(session, output_path)


def _apply_style(cell: Cell, style: dict | None) -> None:
    """Apply style dict from ExcelCell to an openpyxl cell."""
    if not style:
        return

    bg = style.get("bg_color")
    if bg:
        color = bg.lstrip("#")
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

    font_kwargs = {}
    if style.get("font_bold"):
        font_kwargs["bold"] = True
    fc = style.get("font_color")
    if fc:
        font_kwargs["color"] = fc.lstrip("#")
    if font_kwargs:
        cell.font = Font(**font_kwargs)

    nf = style.get("number_format")
    if nf:
        cell.number_format = nf

    border_sides = {}
    for side_name in ("left", "right", "top", "bottom"):
        bs = style.get(f"border_{side_name}")
        if bs:
            border_sides[side_name] = Side(style=bs)
    if border_sides:
        cell.border = Border(**border_sides)


def export_from_cells(
    session: Session,
    sheet_id: int,
    output_path: str | Path,
) -> Path:
    """Build an .xlsx from ExcelCell and ExcelMerge records (no BLOB needed).

    This creates a fresh xlsx entirely from the DB cell data, suitable for
    partial exports like split-by-IP.
    """
    output_path = Path(output_path)

    sheet_obj = session.get(ExcelSheet, sheet_id)
    if not sheet_obj:
        raise ValueError(f"Sheet {sheet_id} not found")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_obj.name

    cells = (
        session.query(ExcelCell)
        .filter(ExcelCell.sheet_id == sheet_id)
        .all()
    )
    for cell_rec in cells:
        if cell_rec.formula_type == "array" and cell_rec.formula_ref:
            value = ArrayFormula(ref=cell_rec.formula_ref, text=cell_rec.raw_value)
        elif cell_rec.formula_type == "dataTable" and cell_rec.formula_ref:
            value = DataTableFormula(ref=cell_rec.formula_ref)
        else:
            value = cell_rec.raw_value
        c = ws.cell(row=cell_rec.row, column=cell_rec.col, value=value)
        _apply_style(c, cell_rec.style)
        if cell_rec.comment:
            c.comment = Comment(cell_rec.comment, "")

    merges = (
        session.query(ExcelMerge)
        .filter(ExcelMerge.sheet_id == sheet_id)
        .all()
    )
    for m in merges:
        ws.merge_cells(
            start_row=m.min_row, start_column=m.min_col,
            end_row=m.max_row, end_column=m.max_col,
        )

    if sheet_obj.header_row:
        ws.freeze_panes = f"A{sheet_obj.header_row + 1}"

    wb.save(output_path)
    wb.close()
    return output_path
