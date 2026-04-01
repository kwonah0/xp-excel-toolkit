"""Split a register map by sheet, with each IP as a sub-sheet."""

from __future__ import annotations

from pathlib import Path

import openpyxl
from sqlalchemy import distinct, insert
from sqlalchemy.orm import Session

from dsm.merge import MergeResolver
from dsm.models import ExcelCell, ExcelMerge, ExcelSheet, ExcelWorkbook
from dsm.domain_models import REGMAP_FIELD_MAP, Register
from dsm.xlsx_parser import _import_ws, _BULK_CHUNK


def _build_ip_sheet(
    session: Session,
    src_sheet: ExcelSheet,
    merger: MergeResolver,
    ip_name: str,
    registers: list[Register],
    dst_wb_id: int,
) -> ExcelSheet:
    """Create a new ExcelSheet in DB for one IP, with remapped rows."""
    src_rows = sorted({reg.excel_row for reg in registers})
    row_map = {src_r: dst_i for dst_i, src_r in enumerate(src_rows, start=2)}

    ip_sheet = ExcelSheet(
        workbook_id=dst_wb_id,
        name=ip_name,
        header_row=1,
    )
    session.add(ip_sheet)
    session.flush()

    sid = ip_sheet.id

    # Copy header row cells (bulk insert)
    header_cells = (
        session.query(ExcelCell)
        .filter(ExcelCell.sheet_id == src_sheet.id, ExcelCell.row == src_sheet.header_row)
        .all()
    )
    if header_cells:
        hdr_dicts = [
            {
                "sheet_id": sid, "row": 1, "col": cell.col,
                "raw_value": cell.raw_value, "style": cell.style,
                "comment": cell.comment,
                "merge_id": None, "is_merge_origin": False,
            }
            for cell in header_cells
        ]
        session.execute(insert(ExcelCell), hdr_dicts)

    # Copy merges with row remapping (bulk insert + query back)
    src_merges = (
        session.query(ExcelMerge)
        .filter(ExcelMerge.sheet_id == src_sheet.id)
        .all()
    )

    # Build merge dicts and track which src merge.id they came from
    merge_dicts: list[dict] = []
    src_merge_ids: list[int] = []  # parallel list: src merge.id per dict
    for merge in src_merges:
        covered = [r for r in src_rows if merge.min_row <= r <= merge.max_row]

        if merge.min_row == merge.max_row and merge.min_row in row_map:
            merge_dicts.append({
                "sheet_id": sid,
                "min_row": row_map[merge.min_row],
                "min_col": merge.min_col,
                "max_row": row_map[merge.min_row],
                "max_col": merge.max_col,
            })
            src_merge_ids.append(merge.id)
        elif len(covered) > 1:
            merge_dicts.append({
                "sheet_id": sid,
                "min_row": row_map[covered[0]],
                "min_col": merge.min_col,
                "max_row": row_map[covered[-1]],
                "max_col": merge.max_col,
            })
            src_merge_ids.append(merge.id)

    merge_id_map: dict[int, int] = {}  # src_merge.id -> new_merge.id
    if merge_dicts:
        for i in range(0, len(merge_dicts), _BULK_CHUNK):
            session.execute(insert(ExcelMerge), merge_dicts[i:i + _BULK_CHUNK])
        session.flush()

        # Query back new merge IDs, match by (min_row, min_col)
        new_merges = (
            session.query(ExcelMerge.id, ExcelMerge.min_row, ExcelMerge.min_col)
            .filter(ExcelMerge.sheet_id == sid)
            .all()
        )
        # Build reverse lookup: (min_row, min_col) -> new_id
        new_merge_lookup = {(mrow, mcol): mid for mid, mrow, mcol in new_merges}
        for md, src_id in zip(merge_dicts, src_merge_ids):
            new_id = new_merge_lookup.get((md["min_row"], md["min_col"]))
            if new_id:
                merge_id_map[src_id] = new_id

    # Copy data cells (bulk insert)
    data_cells = (
        session.query(ExcelCell)
        .filter(
            ExcelCell.sheet_id == src_sheet.id,
            ExcelCell.row.in_(src_rows),
        )
        .all()
    )
    cell_dicts: list[dict] = []
    for cell in data_cells:
        raw_value = cell.raw_value
        if raw_value is None and merger.is_merged(cell.row, cell.col):
            raw_value = merger.get_value(cell.row, cell.col)

        new_merge_id = merge_id_map.get(cell.merge_id) if cell.merge_id else None
        is_origin = merger.is_origin(cell.row, cell.col) if cell.merge_id else False

        cell_dicts.append({
            "sheet_id": sid,
            "row": row_map[cell.row],
            "col": cell.col,
            "raw_value": raw_value,
            "style": cell.style,
            "comment": cell.comment,
            "merge_id": new_merge_id,
            "is_merge_origin": is_origin,
        })

    if cell_dicts:
        for i in range(0, len(cell_dicts), _BULK_CHUNK):
            session.execute(insert(ExcelCell), cell_dicts[i:i + _BULK_CHUNK])
        session.flush()

    return ip_sheet


def split_regmap(
    session: Session,
    path: str | Path,
    output_dir: str | Path,
) -> dict[str, Path]:
    """Split a register map xlsx: one output file per source sheet,
    with each IP as a separate sheet inside.

    e.g. regmap_sample.xlsx (sheets: level2_common, level2_buscon)
         → level2_common.xlsx  (sheets: SENSOR_A, AMPLIFIER, ...)
         → level2_buscon.xlsx  (sheets: GPIO_PORT, TIMER_A, ...)

    Args:
        session: SQLAlchemy session.
        path: Path to the source regmap xlsx.
        output_dir: Directory to write output files.

    Returns:
        Dict mapping source sheet name to output file path.
    """
    path = Path(path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Load workbook once for all sheets
    blob = path.read_bytes()
    wb_xl = openpyxl.load_workbook(path, data_only=False)

    wb_obj = ExcelWorkbook(filename=path.name, blob=blob)
    session.add(wb_obj)
    session.flush()

    results: dict[str, Path] = {}

    # Only process level2_* sheets (register bit-field specs)
    level2_sheets = [sn for sn in wb_xl.sheetnames if sn.startswith("level2_")]

    for sn in level2_sheets:
        ws = wb_xl[sn]
        src_sheet = _import_ws(
            session, ws, wb_obj.id,
            field_map=REGMAP_FIELD_MAP,
            domain_cls=Register,
        )
        merger = MergeResolver.from_db(session, src_sheet.id)

        ip_names = (
            session.query(distinct(Register.name))
            .filter(Register.sheet_id == src_sheet.id, Register.name.isnot(None))
            .all()
        )
        if not ip_names:
            continue

        out_wb = ExcelWorkbook(filename=f"{sn}.xlsx")
        session.add(out_wb)
        session.flush()

        ip_sheets: list[int] = []

        for (ip_name,) in ip_names:
            ip_name = ip_name.strip()
            if not ip_name:
                continue

            registers = (
                session.query(Register)
                .filter(Register.sheet_id == src_sheet.id, Register.name == ip_name)
                .order_by(Register.excel_row)
                .all()
            )
            if not registers:
                continue

            ip_sheet = _build_ip_sheet(
                session, src_sheet, merger, ip_name, registers, out_wb.id,
            )
            ip_sheets.append(ip_sheet.id)

        out_path = output_dir / f"{sn}.xlsx"
        _export_multi_sheet(session, ip_sheets, out_path)
        results[sn] = out_path

    wb_xl.close()
    return results


def _export_multi_sheet(
    session: Session,
    sheet_ids: list[int],
    output_path: Path,
) -> None:
    """Export multiple ExcelSheet records into one xlsx file, each as a sheet."""
    from openpyxl.comments import Comment
    from dsm.exporter import _apply_style

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    for sheet_id in sheet_ids:
        sheet_obj = session.get(ExcelSheet, sheet_id)
        ws = wb.create_sheet(title=sheet_obj.name)

        # Write cells
        cells = (
            session.query(ExcelCell)
            .filter(ExcelCell.sheet_id == sheet_id)
            .all()
        )
        for cell_rec in cells:
            c = ws.cell(row=cell_rec.row, column=cell_rec.col, value=cell_rec.raw_value)
            _apply_style(c, cell_rec.style)
            if cell_rec.comment:
                c.comment = Comment(cell_rec.comment, "")

        # Apply merges
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


def split_regmap_from_db(
    session: Session,
    workbook_filename: str,
    output_dir: str | Path,
) -> dict[str, Path]:
    """Split regmap using already-imported DB data (no xlsx reload needed).

    Args:
        session: SQLAlchemy session with previously imported data.
        workbook_filename: Filename to match in ExcelWorkbook table.
        output_dir: Directory to write output files.

    Returns:
        Dict mapping source sheet name to output file path.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    wb_obj = (
        session.query(ExcelWorkbook)
        .filter_by(filename=workbook_filename)
        .first()
    )
    if not wb_obj:
        raise ValueError(
            f"Workbook '{workbook_filename}' not found in DB. Run 'dsm import' first."
        )

    level2_sheets = (
        session.query(ExcelSheet)
        .filter(ExcelSheet.workbook_id == wb_obj.id, ExcelSheet.name.like("level2_%"))
        .all()
    )

    results: dict[str, Path] = {}

    for src_sheet in level2_sheets:
        merger = MergeResolver.from_db(session, src_sheet.id)

        ip_names = (
            session.query(distinct(Register.name))
            .filter(Register.sheet_id == src_sheet.id, Register.name.isnot(None))
            .all()
        )
        if not ip_names:
            continue

        out_wb = ExcelWorkbook(filename=f"{src_sheet.name}.xlsx")
        session.add(out_wb)
        session.flush()

        ip_sheets: list[int] = []

        for (ip_name,) in ip_names:
            ip_name = ip_name.strip()
            if not ip_name:
                continue

            registers = (
                session.query(Register)
                .filter(Register.sheet_id == src_sheet.id, Register.name == ip_name)
                .order_by(Register.excel_row)
                .all()
            )
            if not registers:
                continue

            ip_sheet = _build_ip_sheet(
                session, src_sheet, merger, ip_name, registers, out_wb.id,
            )
            ip_sheets.append(ip_sheet.id)

        out_path = output_dir / f"{src_sheet.name}.xlsx"
        _export_multi_sheet(session, ip_sheets, out_path)
        results[src_sheet.name] = out_path

    return results
