"""Apache POI (HSSF) 백엔드 — 레거시 .xls 를 **fidelity 보존하며 편집-in-place**.

LibreOffice 전체변환(느림·NFS 5분)의 대안. 원본 .xls 를 POI 로 열어 셀 edit(diff)만 얹어
저장하므로 **수식·셀스타일·병합·차트레코드·미편집 시트가 보존**된다. JVM(jpype)은 프로세스당
1회 기동(warmup)해 공유한다. 메모리는 전체 워크북 로드라 셀 수에 비례(GB) — 읽기는 저메모리
xlrd(`import_xls`) 권장이고, 이 모듈은 **쓰기 경로** 전용.

JVM/POI 타입은 이 모듈 안에만 갇힌다(밖은 순수 파이썬 경로/값). POI jar 경로는 인자 `jars`
또는 env `XP_POI_JARS`(디렉터리·`*.jar`). JRE 없으면 호출자가 LibreOffice 로 fallback.
"""
from __future__ import annotations

import os
from pathlib import Path


def _poi_jars(jars: list[str] | str | None) -> list[str]:
    if jars is None:
        jars = os.environ.get("XP_POI_JARS")
    if not jars:
        raise RuntimeError(
            "POI jars not configured — pass jars= or set XP_POI_JARS to a dir of *.jar")
    if isinstance(jars, str):
        d = Path(jars)
        if d.is_dir():
            return [str(p) for p in sorted(d.glob("*.jar"))]
        return [jars]
    return list(jars)


def ensure_jvm(jars: list[str] | str | None = None, heap: str = "4g",
               jvmpath: str | None = None) -> None:
    """jpype JVM lazy singleton — 프로세스당 1회 기동. 이미 떠 있으면 no-op(재기동 불가).
    대용량 BIFF 위해 IOUtils 안전한계 상향. heap 은 첫 호출에만 적용.

    jvmpath = 번들 JRE 의 `libjvm.so` 경로(인자 또는 env `XP_POI_JVM`) — 시스템 java 불요
    (벤더링 배포). 미지정이면 jpype 기본(시스템 JVM) 사용."""
    import jpype
    import jpype.imports  # noqa: F401  (java import 활성화)
    if jpype.isJVMStarted():
        return
    jvmpath = jvmpath or os.environ.get("XP_POI_JVM") or jpype.getDefaultJVMPath()
    jpype.startJVM(jvmpath, f"-Xmx{heap}", classpath=_poi_jars(jars))
    from org.apache.poi.util import IOUtils
    IOUtils.setByteArrayMaxOverride(2_000_000_000)


def _set_cell(cell, value) -> None:
    """spec 셀은 문자열/JSON 이 대부분 — 문자열로. 숫자/불리언은 그 타입으로(원 kind 보존)."""
    if isinstance(value, bool):
        cell.setCellValue(value)
    elif isinstance(value, (int, float)):
        cell.setCellValue(float(value))
    else:
        cell.setCellValue(str(value))


def write_xls_poi(src_xls: str | Path, edits_by_sheet: dict[str, list[dict]],
                  out_xls: str | Path, *, jars=None, heap: str = "4g",
                  jvmpath: str | None = None) -> Path:
    """원본 `src_xls` 를 열어 `edits_by_sheet` 만 적용 후 `out_xls` 로 저장(나머지 전부 보존).

    edits_by_sheet: ``{sheet_name: [{"row": r, "col": c, "value": v}, ...]}`` — 0-based
    row/col. 미존재 row/cell 은 생성(append). 수식은 `setForceFormulaRecalculation` 으로
    Excel 이 열 때 재계산. 반환 = out_xls(Path).
    """
    ensure_jvm(jars, heap, jvmpath)
    from org.apache.poi.hssf.usermodel import HSSFWorkbook
    from java.io import FileInputStream, FileOutputStream

    fis = FileInputStream(str(src_xls))
    try:
        wb = HSSFWorkbook(fis)
    finally:
        fis.close()
    try:
        for sheet_name, edits in edits_by_sheet.items():
            sh = wb.getSheet(sheet_name)
            if sh is None:
                raise KeyError(f"sheet not found in {src_xls}: {sheet_name!r}")
            for e in edits:
                row = sh.getRow(e["row"]) or sh.createRow(e["row"])
                cell = row.getCell(e["col"]) or row.createCell(e["col"])
                _set_cell(cell, e["value"])
        if any(edits_by_sheet.values()):
            wb.setForceFormulaRecalculation(True)
        fos = FileOutputStream(str(out_xls))
        try:
            wb.write(fos)
        finally:
            fos.close()
    finally:
        wb.close()
    return Path(out_xls)
