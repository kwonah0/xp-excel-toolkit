"""Case 1: Parse level2_* 시트와 memorymap 시트를 각각 파싱.

사용법:
    uv run python tests/manual/case_1_parse_sheets.py

이 스크립트는:
  1. parse_level2()로 level2_common, level2_buscon 시트를 각각 파싱
  2. parse_memorymap()으로 memorymap 시트를 파싱
  3. 각 시트별 Register / MemoryMapEntry 개수 확인
  4. IP별 레지스터 그룹핑
  5. memorymap에서 group과 baseaddr 매핑 확인
"""

import os
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
os.chdir(PROJECT_ROOT)

from sqlalchemy import distinct

from excel_toolkit import (
    Register, MemoryMapEntry,
    parse_level2, parse_memorymap,
    init_db,
)

SAMPLE = Path("samples/regmap_sample.xlsx")
DB_PATH = Path("samples/case_1.db")


def main():
    DB_PATH.unlink(missing_ok=True)
    Session = init_db(f"sqlite:///{DB_PATH}")

    with Session() as session:
        # ── 1. level2_common 파싱 ──────────────────────────────
        print("=" * 60)
        print("  Parsing level2_common")
        print("=" * 60)

        sheet_common = parse_level2(session, SAMPLE, "level2_common")
        session.commit()

        regs_common = (
            session.query(Register)
            .filter_by(sheet_id=sheet_common.id)
            .order_by(Register.excel_row)
            .all()
        )
        print(f"  Sheet: {sheet_common.name}")
        print(f"  Registers: {len(regs_common)}개")

        # IP별 그룹핑
        ip_groups: dict[str, list[Register]] = {}
        for r in regs_common:
            ip_groups.setdefault(r.name, []).append(r)

        print(f"\n  IP별 레지스터 수:")
        for ip_name, ip_regs in ip_groups.items():
            indx_set = {r.indx for r in ip_regs}
            page_set = {r.page for r in ip_regs}
            print(f"    {ip_name:>12}: {len(ip_regs)}개  "
                  f"INDX={indx_set}  PAGE={page_set}")

        # ── 2. level2_buscon 파싱 ──────────────────────────────
        print(f"\n{'=' * 60}")
        print("  Parsing level2_buscon")
        print("=" * 60)

        sheet_buscon = parse_level2(session, SAMPLE, "level2_buscon")
        session.commit()

        regs_buscon = (
            session.query(Register)
            .filter_by(sheet_id=sheet_buscon.id)
            .order_by(Register.excel_row)
            .all()
        )
        print(f"  Sheet: {sheet_buscon.name}")
        print(f"  Registers: {len(regs_buscon)}개")

        ip_groups_b: dict[str, list[Register]] = {}
        for r in regs_buscon:
            ip_groups_b.setdefault(r.name, []).append(r)

        print(f"\n  IP별 레지스터 수:")
        for ip_name, ip_regs in ip_groups_b.items():
            indx_set = {r.indx for r in ip_regs}
            page_set = {r.page for r in ip_regs}
            print(f"    {ip_name:>12}: {len(ip_regs)}개  "
                  f"INDX={indx_set}  PAGE={page_set}")

        # ── 3. memorymap 파싱 ──────────────────────────────────
        print(f"\n{'=' * 60}")
        print("  Parsing memorymap")
        print("=" * 60)

        sheet_mm = parse_memorymap(session, SAMPLE)
        session.commit()

        entries = (
            session.query(MemoryMapEntry)
            .filter_by(sheet_id=sheet_mm.id)
            .order_by(MemoryMapEntry.excel_row)
            .all()
        )
        print(f"  Sheet: {sheet_mm.name}")
        print(f"  Entries: {len(entries)}개\n")

        print(f"  {'BASEADDR':>10} {'Group':>12} {'midgroup':>10} "
              f"{'Comment':<30} {'special':<10}")
        print("  " + "-" * 80)
        for e in entries:
            print(f"  {e.baseaddr or '-':>10} {e.group or '-':>12} "
                  f"{e.midgroup or '-':>10} {e.comment or '-':<30} "
                  f"{e.special or '-':<10}")

        # ── 4. Cross-reference: memorymap ↔ level2 ────────────
        print(f"\n{'=' * 60}")
        print("  Cross-reference: memorymap ↔ level2 registers")
        print("=" * 60)

        all_ip_names = set(ip_groups.keys()) | set(ip_groups_b.keys())
        mm_groups = {e.group for e in entries if e.group}

        print(f"\n  memorymap에 있는 IP: {sorted(mm_groups)}")
        print(f"  level2에 있는 IP:    {sorted(all_ip_names)}")

        matched = mm_groups & all_ip_names
        mm_only = mm_groups - all_ip_names
        l2_only = all_ip_names - mm_groups

        print(f"\n  매칭:          {sorted(matched)}")
        if mm_only:
            print(f"  memorymap only: {sorted(mm_only)}")
        if l2_only:
            print(f"  level2 only:    {sorted(l2_only)}")

    print(f"\nDB 저장 위치: {DB_PATH.resolve()}")
    print("Done!")


if __name__ == "__main__":
    main()
