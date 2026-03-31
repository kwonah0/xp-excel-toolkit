"""Case 0: Import regmap_sample.xlsx → DB, inspect registers and cells.

사용법:
    uv run python tests/manual/case_0_import.py

이 스크립트는:
  1. regmap_sample.xlsx의 level2_common 시트를 파싱하여 SQLite DB에 저장
  2. Register 도메인 객체 조회 (모든 값은 string)
  3. 세로 merge(INDX, PAGE) 값이 올바르게 전파되었는지 확인
  4. 가로 merge(bit field) 값이 올바르게 전파되었는지 확인
  5. ExcelCell의 스타일(색상) 정보 확인
  6. ExcelMerge 레코드 조회
"""

import os
import sys
from pathlib import Path

# 프로젝트 루트로 이동
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
os.chdir(PROJECT_ROOT)

from excel_toolkit import (
    ExcelCell, ExcelMerge, ExcelSheet,
    Register, REGMAP_FIELD_MAP,
    import_xlsx, init_db,
)

SAMPLE = Path("samples/regmap_sample.xlsx")
DB_PATH = Path("samples/case_0.db")


def main():
    # ── 1. DB 초기화 + Import ──────────────────────────────────
    DB_PATH.unlink(missing_ok=True)
    Session = init_db(f"sqlite:///{DB_PATH}")

    with Session() as session:
        # import_xlsx에 domain_cls와 field_map을 넘기면
        # 헤더를 자동 감지하고, 각 데이터 행을 Register 객체로 생성합니다.
        sheet = import_xlsx(
            session,
            SAMPLE,
            sheet_name="level2_common",
            field_map=REGMAP_FIELD_MAP,
            domain_cls=Register,
        )
        session.commit()

        print(f"Sheet: {sheet.name}, header_row: {sheet.header_row}")

        # ── 2. Register 목록 조회 ──────────────────────────────
        regs = (
            session.query(Register)
            .filter_by(sheet_id=sheet.id)
            .order_by(Register.excel_row)
            .all()
        )
        print(f"\n{'='*80}")
        print(f"  Registers ({len(regs)}개)")
        print(f"{'='*80}")
        print(f"{'row':>4} {'TYPE':>4} {'INDX':>5} {'PAGE':>5} {'PARA':>5} "
              f"{'NAME':>10}  {'D7':>12} {'D6':>12} {'D5':>12} {'D4':>12} "
              f"{'D3':>12} {'D2':>12} {'D1':>12} {'D0':>12}  {'INIT':>6}")
        print("-" * 160)
        for r in regs:
            print(f"{r.excel_row:>4} {r.type or '-':>4} {r.indx or '-':>5} "
                  f"{r.page or '-':>5} {r.para or '-':>5} {r.name or '-':>10}  "
                  f"{r.d7 or '-':>12} {r.d6 or '-':>12} {r.d5 or '-':>12} "
                  f"{r.d4 or '-':>12} {r.d3 or '-':>12} {r.d2 or '-':>12} "
                  f"{r.d1 or '-':>12} {r.d0 or '-':>12}  {r.init or '-':>6}")

        # ── 3. 세로 merge 확인 (INDX, PAGE 전파) ──────────────
        print(f"\n{'='*80}")
        print(f"  Vertical merge 확인")
        print(f"{'='*80}")
        # SENSOR_A 그룹: INDX=57이 4행 연속
        sensor_regs = [r for r in regs if r.name == "SENSOR_A"]
        for r in sensor_regs:
            print(f"  row={r.excel_row}: INDX={r.indx}, PAGE={r.page}")
            assert r.indx == "57", f"Expected INDX=57, got {r.indx}"
            assert r.page == "0", f"Expected PAGE=0, got {r.page}"
        print("  → SENSOR_A: INDX/PAGE 전파 OK")

        # ── 4. 가로 merge 확인 (bit field 전파) ────────────────
        print(f"\n{'='*80}")
        print(f"  Horizontal merge 확인 (bit fields)")
        print(f"{'='*80}")
        ctrl = sensor_regs[0]  # PARA=0, CTRL register
        # MODE[1:0]이 D6, D5에 모두 채워져야 함
        print(f"  SENSOR_A PARA=0: D6={ctrl.d6}, D5={ctrl.d5}")
        assert ctrl.d6 == "MODE[1:0]" and ctrl.d5 == "MODE[1:0]"
        print("  → MODE[1:0] merge 전파 OK")

        # RSVD가 D3~D0에 모두 채워져야 함
        print(f"  SENSOR_A PARA=0: D3={ctrl.d3}, D2={ctrl.d2}, D1={ctrl.d1}, D0={ctrl.d0}")
        assert all(getattr(ctrl, f"d{i}") == "RSVD" for i in range(4))
        print("  → RSVD merge 전파 OK")

        # ── 5. 스타일(색상) 확인 ───────────────────────────────
        print(f"\n{'='*80}")
        print(f"  Cell 스타일 확인")
        print(f"{'='*80}")

        # EN 셀 (1-bit → 초록)
        en_cell = session.query(ExcelCell).filter_by(
            sheet_id=sheet.id, row=ctrl.excel_row, col=6  # D7 = col 6
        ).first()
        print(f"  EN cell (row={en_cell.row}, col={en_cell.col}):")
        print(f"    value = {en_cell.raw_value}")
        print(f"    style = {en_cell.style}")
        if en_cell.style:
            print(f"    bg_color = {en_cell.style.get('bg_color')}")

        # MODE[1:0] 셀 (multi-bit → 노란)
        mode_cell = session.query(ExcelCell).filter_by(
            sheet_id=sheet.id, row=ctrl.excel_row, col=7  # D6 = col 7
        ).first()
        print(f"  MODE[1:0] cell (row={mode_cell.row}, col={mode_cell.col}):")
        print(f"    value = {mode_cell.raw_value}")
        print(f"    style = {mode_cell.style}")

        # RSVD 셀 (reserved → 회색)
        rsvd_cell = session.query(ExcelCell).filter_by(
            sheet_id=sheet.id, row=ctrl.excel_row, col=10  # D3 = col 10
        ).first()
        print(f"  RSVD cell (row={rsvd_cell.row}, col={rsvd_cell.col}):")
        print(f"    value = {rsvd_cell.raw_value}")
        print(f"    style = {rsvd_cell.style}")

        # ── 6. Merge 레코드 조회 ───────────────────────────────
        print(f"\n{'='*80}")
        print(f"  ExcelMerge 레코드")
        print(f"{'='*80}")
        merges = (
            session.query(ExcelMerge)
            .filter_by(sheet_id=sheet.id)
            .order_by(ExcelMerge.min_row, ExcelMerge.min_col)
            .all()
        )
        h_count = v_count = 0
        for m in merges:
            kind = "V" if m.min_col == m.max_col else "H"
            if kind == "H":
                h_count += 1
            else:
                v_count += 1
            print(f"  r{m.min_row}c{m.min_col}:r{m.max_row}c{m.max_col} ({kind})")
        print(f"\n  총 {len(merges)}개 — Horizontal: {h_count}, Vertical: {v_count}")

    print(f"\nDB 저장 위치: {DB_PATH.resolve()}")
    print("Done!")


if __name__ == "__main__":
    main()
