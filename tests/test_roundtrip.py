"""End-to-end test: create regmap sample -> import -> modify -> export -> verify."""

import os
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
os.chdir(PROJECT_ROOT)

from excel_toolkit import (
    Base, ExcelCell, ExcelMerge, ExcelSheet, ExcelWorkbook,
    REGMAP_FIELD_MAP, Register,
    import_sheet, export_regmap_xlsx, init_db,
)

sys.path.insert(0, str(PROJECT_ROOT / "scripts"))
from create_regmap_sample import create_regmap_xlsx


def banner(text: str):
    print(f"\n{'='*60}")
    print(f"  {text}")
    print(f"{'='*60}")


def classify_merges(merges):
    h, v = [], []
    for m in merges:
        if m.max_col > m.min_col:
            h.append(m)
        else:
            v.append(m)
    return h, v


def test_regmap():
    banner("TEST: Register Map Roundtrip (.xlsx)")

    # -- 1. Create sample -----------------------------------------------
    xlsx_path = create_regmap_xlsx()

    # -- 2. Import ------------------------------------------------------
    db_path = "sqlite:///test_regmap.db"
    Path("test_regmap.db").unlink(missing_ok=True)

    SessionMaker = init_db(db_path)
    session = SessionMaker()

    column_map: dict[str, int] = {}
    sheet = import_sheet(
        session, xlsx_path,
        field_map=REGMAP_FIELD_MAP,
        domain_cls=Register,
        column_map=column_map,
    )
    session.commit()

    print(f"\n[Import] Sheet: {sheet.name}, header_row: {sheet.header_row}")
    print(f"[Import] Column map: {column_map}")

    # -- 3. Verify registers --------------------------------------------
    regs = session.query(Register).order_by(Register.excel_row).all()
    print(f"\n[Import] {len(regs)} registers loaded:")
    for r in regs:
        print(f"  row={r.excel_row}: {r.type:>3} idx={r.indx} p={r.page} "
              f"para={r.para} {r.name:>8}  "
              f"D7={r.d7}  D6={r.d6}  D5={r.d5}  D4={r.d4}  "
              f"D3={r.d3}  D2={r.d2}  D1={r.d1}  D0={r.d0}  init={r.init}")

    assert len(regs) == 17, f"Expected 17 registers, got {len(regs)}"
    print(f"[Import] {len(regs)} registers loaded OK")

    # All values should be strings
    for r in regs:
        for field in ["type", "indx", "page", "para", "name", "init"]:
            val = getattr(r, field)
            assert val is None or isinstance(val, str), f"Field {field} is not str: {val!r}"
    print("[Import] All values are strings OK")

    # Check vertical merge fill (INDX, PAGE) — SENSOR_A spans rows 2-5, all share indx=57
    reg_sensor = session.query(Register).filter_by(name="SENSOR_A", para="1").first()
    assert reg_sensor.indx == "57", f"Expected INDX=57, got {reg_sensor.indx}"
    assert reg_sensor.page == "0", f"Expected PAGE=0, got {reg_sensor.page}"
    print("[Import] Vertical merge fill (INDX/PAGE) OK")

    # Check horizontal merge fill (bit fields) — SENSOR_A row 2 has MODE[1:0] in D6-D5
    reg_ctrl = session.query(Register).filter_by(name="SENSOR_A", para="0").first()
    assert reg_ctrl.d6 == "MODE[1:0]", f"Expected MODE[1:0] in D6, got {reg_ctrl.d6}"
    assert reg_ctrl.d5 == "MODE[1:0]", f"Expected MODE[1:0] in D5, got {reg_ctrl.d5}"
    print("[Import] Horizontal merge fill (bit fields) OK")

    # -- 4. Verify merges -----------------------------------------------
    merges = session.query(ExcelMerge).filter_by(sheet_id=sheet.id).all()
    h_merges, v_merges = classify_merges(merges)
    print(f"\n[Merge] Total: {len(merges)} — Horizontal: {len(h_merges)}, Vertical: {len(v_merges)}")

    for m in merges:
        kind = "H" if m.max_col > m.min_col else "V"
        print(f"  r{m.min_row}c{m.min_col}:r{m.max_row}c{m.max_col} ({kind})")

    assert len(h_merges) == 26, f"Expected 26 horizontal merges, got {len(h_merges)}"
    assert len(v_merges) == 16, f"Expected 16 vertical merges, got {len(v_merges)}"
    print("[Merge] Merge counts OK")

    # -- 5. Verify styles (colors) --------------------------------------
    # Green cell (1-bit, e.g. EN at row 2, D7 = col 6)
    en_cell = session.query(ExcelCell).filter_by(
        sheet_id=sheet.id, row=2, col=6
    ).first()
    assert en_cell and en_cell.style, "EN cell should have style"
    print(f"[Style] EN (1-bit green): {en_cell.style}")

    # Yellow cell (multi-bit, e.g. MODE[1:0] origin at row 2, D6 = col 7)
    mode_cell = session.query(ExcelCell).filter_by(
        sheet_id=sheet.id, row=2, col=7
    ).first()
    assert mode_cell and mode_cell.style, "MODE cell should have style"
    print(f"[Style] MODE[1:0] (merged yellow): {mode_cell.style}")

    # Gray cell (RSVD at row 2, D3 = col 10)
    rsvd_cell = session.query(ExcelCell).filter_by(
        sheet_id=sheet.id, row=2, col=10
    ).first()
    assert rsvd_cell and rsvd_cell.style, "RSVD cell should have style"
    print(f"[Style] RSVD (gray): {rsvd_cell.style}")

    # -- 6. Modify + Export ---------------------------------------------
    print("\n[Modify] Changing SENSOR_A[para=0].INIT from 0x00 to 0x80...")
    reg_ctrl.init = "0x80"
    session.commit()

    wb = session.query(ExcelWorkbook).first()
    output = Path("samples/regmap_output.xlsx")
    export_regmap_xlsx(session, wb.id, output, column_map=column_map)
    print(f"[Export] Written to: {output}")

    # -- 7. Verify export -----------------------------------------------
    import openpyxl as oxl
    wb_check = oxl.load_workbook(output)
    ws_check = wb_check.active

    # INIT value updated
    init_cell = ws_check.cell(row=2, column=14)
    assert init_cell.value == "0x80", f"Expected 0x80, got {init_cell.value}"
    print("[Verify] INIT value updated OK")

    # Merges preserved
    merge_count = len(ws_check.merged_cells.ranges)
    print(f"[Verify] Merged ranges preserved: {merge_count}")
    assert merge_count == 42, f"Expected 42 merges, got {merge_count}"

    # Colors preserved (check EN cell green)
    en_check = ws_check.cell(row=2, column=6)
    if en_check.fill and en_check.fill.fgColor:
        print(f"[Verify] EN cell color preserved: {en_check.fill.fgColor.rgb}")

    # MODE[1:0] yellow preserved
    mode_check = ws_check.cell(row=2, column=7)
    if mode_check.fill and mode_check.fill.fgColor:
        print(f"[Verify] MODE cell color preserved: {mode_check.fill.fgColor.rgb}")

    wb_check.close()
    session.close()
    print("\n  Register map roundtrip PASSED")


if __name__ == "__main__":
    test_regmap()
