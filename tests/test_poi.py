"""POI(.xls) fidelity edit-in-place — write_xls_poi 가 원본+diff 로 셀만 고치고 수식/병합/
타시트를 보존하는지. JVM(jpype)+POI jar(env XP_POI_JARS) 없으면 skip."""
from __future__ import annotations

import os
from pathlib import Path

import pytest

_HAVE_JPYPE = True
try:
    import jpype  # noqa: F401
except Exception:
    _HAVE_JPYPE = False

_JARS = os.environ.get("XP_POI_JARS")
pytestmark = pytest.mark.skipif(
    not (_HAVE_JPYPE and _JARS and list(Path(_JARS).glob("*.jar"))),
    reason="jpype + POI jars(XP_POI_JARS) 필요")


def _make_fixture(path: Path) -> None:
    """POI 로 .xls fixture 생성 — data(수식·병합) + other(보존 확인)."""
    from xp_excel_toolkit.ingest.poi import ensure_jvm
    ensure_jvm(heap="2g")
    from org.apache.poi.hssf.usermodel import HSSFWorkbook
    from org.apache.poi.ss.util import CellRangeAddress
    from java.io import FileOutputStream
    wb = HSSFWorkbook()
    d = wb.createSheet("data")
    d.createRow(0).createCell(0).setCellValue("x")
    d.getRow(0).createCell(1).setCellValue(10.0)        # B1
    d.createRow(1).createCell(1).setCellValue(20.0)     # B2
    d.createRow(2).createCell(1).setCellFormula("B1+B2")  # B3 수식
    d.createRow(3).createCell(0).setCellValue("merged")
    d.addMergedRegion(CellRangeAddress(3, 3, 0, 1))     # A4:B4 병합
    wb.createSheet("other").createRow(0).createCell(0).setCellValue("keep")
    fos = FileOutputStream(str(path))
    try:
        wb.write(fos)
    finally:
        fos.close()
    wb.close()


def test_write_xls_poi_preserves_fidelity(tmp_path):
    from xp_excel_toolkit.ingest.poi import write_xls_poi, ensure_jvm
    src = tmp_path / "fixture.xls"
    out = tmp_path / "out.xls"
    _make_fixture(src)

    # B1(0,1) = 100 으로 편집 (원본+diff)
    write_xls_poi(src, {"data": [{"row": 0, "col": 1, "value": 100}]}, out)

    # 재오픈 검증
    ensure_jvm()
    from org.apache.poi.hssf.usermodel import HSSFWorkbook
    from java.io import FileInputStream
    fis = FileInputStream(str(out))
    wb = HSSFWorkbook(fis); fis.close()
    try:
        d = wb.getSheet("data")
        assert d.getRow(0).getCell(1).getNumericCellValue() == 100.0     # 편집 반영
        assert str(d.getRow(2).getCell(1).getCellFormula()) == "B1+B2"   # 수식 보존
        assert d.getRow(0).getCell(0).getStringCellValue() == "x"        # 미편집 셀 보존
        assert [str(r.formatAsString()) for r in d.getMergedRegions()] == ["A4:B4"]  # 병합 보존
        assert wb.getSheet("other").getRow(0).getCell(0).getStringCellValue() == "keep"  # 타시트 보존
    finally:
        wb.close()
