# xp-excel-toolkit cookbook

Excel 파일을 SQLite 안에 cell 단위로 보존하고, 그 위에서 도메인 모델 ·
diff · export 를 자유롭게 쌓는 toolkit. 도메인 무관 (regmap 도 그 위에
서 만든 한 사례).

이 문서는 **어떻게 쓰는지** 를 task 기반으로 모은다. 각 레시피는
"원하는 것 → 코드 → 결과 → 주의" 형태.

---

## 핵심 개념

**4-table cell schema** — Excel 의 워크북 한 권을 4 개 테이블로 분해.

| 테이블 | 내용 |
|---|---|
| `excel_workbook` | 파일 메타데이터 + 원본 BLOB (round-trip 용) |
| `excel_sheet` | 시트 이름, 헤더 행 위치 |
| `excel_cell` | 각 셀의 `raw_value`/`cached_value` (formula 결과)/`style` (JSON)/`comment`/`formula_type`/`merge_id` |
| `excel_merge` | merge range (R, C 범위) |

`excel_cell.style` 은 dict — `bg_color`, `font_bold`, `font_color`,
`number_format`, `border_left/right/top/bottom` 까지 보존. round-trip
export 시 그대로 재현.

**도메인 분리** — toolkit 은 "cell" 까지만 안다. Register / Bom / Netlist
같은 도메인 모델은 사용자가 자기 패키지에 정의. toolkit 은 그걸 만드
는 데 필요한 helper (header lookup, row 순회, merge 충돌 처리, cell
diff) 만 제공.

```
사용자 패키지
└── domain models (Register, MemoryMapEntry) + builders
        ↑ uses
xp_excel_toolkit
├── ExcelCell/Sheet/Merge/Workbook  (parse 결과)
├── import_xlsx / import_xls          (xlsx/xls → DB)
├── helpers (find_header_row_db, iter_rows_by_header)
├── exporter (DB → xlsx)
├── merge helpers (detect/resolve conflicts)
└── diff (DiffBase, DiffCell, diff_cells)
```

---

## 설치

```bash
# devpi 에서 (조직 내부 index)
uv add xp-excel-toolkit

# 또는 PyPI / 로컬 path
uv add --editable ../excel-toolkit/main
```

요구사항 — Python 3.12+. 의존성 — `openpyxl`, `sqlalchemy>=2.0`,
`xlrd`/`xlwt` (xls 지원).

---

## Quick start

```python
from xp_excel_toolkit import init_db, import_xlsx, ExcelSheet, ExcelCell

Session = init_db("sqlite:///work.db")
with Session() as s:
    sheets = import_xlsx(s, "register_map.xlsx")
    s.commit()

# 셀 조회
with Session() as s:
    sheet = s.query(ExcelSheet).filter_by(name="level2_common").one()
    a1 = s.query(ExcelCell).filter_by(sheet_id=sheet.id, row=1, col=1).one()
    print(a1.raw_value, a1.style)
```

---

## Recipe 1 — xlsx 전체를 DB 로 import

**원함** — 워크북 한 권의 모든 시트 · cell · merge · style 을 SQLite 에
저장.

```python
from xp_excel_toolkit import init_db, import_xlsx

Session = init_db("sqlite:///work.db")
with Session() as s:
    sheets = import_xlsx(s, "design.xlsx", on_progress=print)
    # sheets: list[ExcelSheet]
    s.commit()
    print(f"{len(sheets)} sheets, {sum(len(sh.cells) for sh in sheets)} cells")
```

**결과** — 원본 BLOB 도 `excel_workbook.blob` 에 함께 저장돼서, 나중에
export 시 원본 위에 변경만 덮어쓰는 정확한 round-trip 가능.

**주의** — `with_formulas=True` 안 주면 formula 는 cached value 만 들
어감 (formula 문자열 X). formula 자체 보존하려면 Recipe 11 참고.

---

## Recipe 2 — xls (구형 포맷) 자동 변환 후 import

**원함** — `.xls` 파일을 받았지만 `.xlsx` 로 동일하게 다루고 싶다.

```python
from xp_excel_toolkit import resolve_db, init_db, import_xlsx

# resolve_db: 입력이 xlsx 면 그대로, xls 면 LibreOffice 로 .xlsx 변환
# 결과를 DB 화. cache 적중 시 skip.
db_path, was_cached = resolve_db("legacy.xls", on_progress=print)
print("db:", db_path, "from-cache:", was_cached)

# 또는 변환만:
from xp_excel_toolkit import convert_xls_to_xlsx
xlsx_path = convert_xls_to_xlsx("legacy.xls")
```

**결과** — cache 디렉토리 (`<cwd>/.xltk_cache/` 기본,
`xp_excel_toolkit.config.CACHE_DIR = "<path>"` 로 변경) 안에
`<name>_<hash>.xlsx` + `.db` 가 저장. 같은 파일 재호출 시 hash 일치하면
캐시 hit.

**주의** — LibreOffice 가 설치돼 있어야 함 (`/usr/bin/libreoffice`,
`/usr/bin/soffice` 자동 탐색). docker · CI 환경에서는 명시 설치 필요.

---

## Recipe 3 — 한 sheet 만 import

**원함** — 워크북에서 한 sheet 만 필요. 빠르게.

```python
from xp_excel_toolkit import init_db, import_sheet

Session = init_db("sqlite:///work.db")
with Session() as s:
    sheet = import_sheet(s, "design.xlsx", sheet_name="level2_common")
    s.commit()
    print(f"sheet_id={sheet.id} cells={len(sheet.cells)}")
```

`sheet_name=None` 이면 active sheet (보통 첫 sheet).

---

## Recipe 4 — Header 기반 row 순회 (도메인 builder 의 핵심 패턴)

**원함** — `excel_cell` 에 row/col 으로 저장된 데이터를, 헤더 이름으로
접근하면서 한 행씩 처리.

```python
from xp_excel_toolkit import find_header_row_db, iter_rows_by_header

with Session() as s:
    sheet_id = sheet.id  # Recipe 3 의 결과

    # 헤더 행 자동 탐지 — 기대하는 헤더 이름이 한 row 에 모이는 곳
    header_row = find_header_row_db(
        s, sheet_id,
        expected_headers=["NAME", "INDX", "PAGE"],
        match="any",   # "all" 이면 모두 일치하는 첫 row
    )

    # 그 row 의 col → header 매핑을 자동 빌드, 데이터 row 를 dict 로
    for row_idx, row in iter_rows_by_header(
        s, sheet_id, header_row,
        headers=["NAME", "INDX", "D7", "D6", "D5", "D4"],  # 화이트리스트
    ):
        print(row_idx, row)
        # 출력 예: (5, {"NAME": "FOO", "INDX": "0x00", "D7": "0", ...})
```

**결과** — 사용자 코드가 column index 를 절대 만지지 않음. 헤더 텍스트
가 바뀌면 (예: "NAME" → "Name") 한 군데만 갱신.

**주의** — multi-row header / pivot / 비-flat 레이아웃 (예: overview
형식의 key-value) 은 `ExcelCell` 을 직접 SELECT 해서 처리.

---

## Recipe 5 — 도메인 모델 정의 + 자체 builder 작성

**원함** — Register 도메인 모델 (excel-toolkit 모름) 을 만들고, 위 헬퍼
로 cell 에서 채워 넣기.

```python
# my_pkg/models.py
from sqlalchemy import ForeignKey, Text
from sqlalchemy.orm import Mapped, mapped_column

from xp_excel_toolkit import Base, ExcelSheet  # Base 공유


class Register(Base):
    __tablename__ = "register"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str | None] = mapped_column(Text)
    indx: Mapped[str | None] = mapped_column(Text)
    init: Mapped[str | None] = mapped_column(Text)

    # 원본 cell 추적 — 어느 sheet 의 어느 row 에서 왔는지
    sheet_id: Mapped[int | None] = mapped_column(ForeignKey("excel_sheet.id"))
    excel_row: Mapped[int | None]


# my_pkg/builders.py
from sqlalchemy.orm import Session
from xp_excel_toolkit import (
    ExcelSheet, find_header_row_db, iter_rows_by_header,
)
from my_pkg.models import Register

REGISTER_HEADERS = ["NAME", "INDX", "INIT"]
FIELD_MAP = {"NAME": "name", "INDX": "indx", "INIT": "init"}


def build_registers(session: Session) -> int:
    """level2_* sheet 마다 데이터 row 를 Register 로 변환."""
    count = 0
    for sheet in session.query(ExcelSheet).filter(
        ExcelSheet.name.like("level2_%")
    ).all():
        hdr = find_header_row_db(session, sheet.id, REGISTER_HEADERS, match="any")
        if not hdr:
            continue
        for row_idx, row in iter_rows_by_header(
            session, sheet.id, hdr, headers=REGISTER_HEADERS
        ):
            r = Register(
                sheet_id=sheet.id,
                excel_row=row_idx,
                **{FIELD_MAP[h]: row.get(h) for h in REGISTER_HEADERS},
            )
            session.add(r)
            count += 1
    session.flush()
    return count
```

**사용**:

```python
from xp_excel_toolkit import init_db, import_xlsx
from my_pkg.builders import build_registers
from my_pkg.models import Register   # ← import 만 해도 Base.metadata 에 등록

Session = init_db("sqlite:///work.db")
with Session() as s:
    import_xlsx(s, "design.xlsx")     # phase 1 — cell
    build_registers(s)                # phase 2 — domain
    s.commit()

    for r in s.query(Register).limit(5):
        print(r.name, r.indx, "from sheet_id", r.sheet_id, "row", r.excel_row)
```

**결과** — toolkit 의 cell 테이블과 사용자 도메인 테이블이 같은 `Base`
metadata 공유. `init_db()` 한 번이 양쪽 모두 `create_all`.

**주의** — `init_db()` 호출 전에 도메인 모델 모듈을 import 해 둬야 등록
완료. 일반적으로 `my_pkg/__init__.py` 에서 `from my_pkg.models import ...`
하면 자동 chain.

---

## Recipe 6 — DB → xlsx 로 export (셀 + 스타일 + merge)

**원함** — DB 의 한 sheet 를 그대로 xlsx 한 파일로 출력.

```python
from xp_excel_toolkit import export_from_cells

with Session() as s:
    out = export_from_cells(s, sheet_id=42, output_path="out/sheet42.xlsx")
    print("wrote", out)
```

`export_from_cells` 는 **fresh wb 생성** — `excel_cell` 의 모든 row
+ style + comment + formula + merge 를 새 xlsx 에 쓴다. partial export
(예: split-by-IP) 용도.

**전체 wb 를 원본 BLOB 위에서 round-trip** 하고 싶으면 (모든 sheet + 도
메인 변경 일부만 덮어쓰기), `excel_workbook.blob` 을 BytesIO 로 열어
openpyxl 로 수정하는 패턴을 사용자 코드에서 작성.

---

## Recipe 7 — 두 워크북의 cell-level diff

**원함** — 어제 본 `design.xlsx` 와 오늘 받은 `design.xlsx` 를 셀 단위
로 비교, 변경 / 추가 / 삭제 / 이동 분류.

```python
from xp_excel_toolkit import init_db, import_xlsx
from xp_excel_toolkit.diff import (
    diff_cells, load_cells_by_sheet, load_merge_ranges,
)

# 두 DB 로 import
SessionOld = init_db("sqlite:///old.db")
SessionNew = init_db("sqlite:///new.db")

with SessionOld() as old, SessionNew() as new:
    import_xlsx(old, "old.xlsx"); old.commit()
    import_xlsx(new, "new.xlsx"); new.commit()

    diffs = diff_cells(
        load_cells_by_sheet(old),
        load_cells_by_sheet(new),
        compare_style=True,           # 셀 스타일도 비교
        compare_merge=True,
        merges_a=load_merge_ranges(old),
        merges_b=load_merge_ranges(new),
        on_progress=print,
    )

for d in diffs[:5]:
    print(d.status, d.sheet, f"r{d.row}c{d.col}", d.old_value, "→", d.new_value)
```

**결과** — `diffs` 는 `DiffCell` 객체 리스트. 각 항목의 `status` ∈
`{added, removed, changed, moved}`. **smart diff** (difflib 의
SequenceMatcher) 기반이라 행 1 개가 삽입되어도 그 아래가 통째 changed
로 cascade 되지 않음. 같은 시그니처가 다른 위치로 옮겨가면 `moved` 로
재분류.

**주의** — `compare_style=True` 면 비교 대상이 커서 큰 wb (1만 행+) 에
선 느려질 수 있음. 기본은 값만 비교.

---

## Recipe 8 — 도메인 diff 모델을 같은 DiffBase 위에 정의

**원함** — cell diff 외에 Register 단위 diff (added/removed/changed
register) 도 같은 SQLite 안에 한 번에 저장.

```python
# my_pkg/diff_models.py
from sqlalchemy import Text
from sqlalchemy.orm import Mapped, mapped_column

from xp_excel_toolkit.diff import DiffBase   # ← 공유 Base


class DiffRegister(DiffBase):
    __tablename__ = "diff_register"
    id: Mapped[int] = mapped_column(primary_key=True)
    status: Mapped[str] = mapped_column(Text)   # added/removed/changed
    sheet: Mapped[str | None] = mapped_column(Text)
    old_name: Mapped[str | None] = mapped_column(Text)
    new_name: Mapped[str | None] = mapped_column(Text)
    # ... 필드 늘리기
```

```python
from xp_excel_toolkit.diff import init_diff_db
from my_pkg import diff_models   # ← import 만 해도 DiffRegister 등록

DiffSession = init_diff_db("sqlite:///diff.db")
# 같은 호출 1번에 diff_cell + diff_register 둘 다 create_all
```

**주의** — 등록 순서가 invisible dependency. 사용자에게 `from my_pkg
import diff` 같은 entry 만 노출하고, 그 모듈 안에서 모델들을 import 해
두면 chain 으로 정리.

---

## Recipe 9 — Merge cell 에 쓸 때 충돌 감지 + 정책 기반 해결

**원함** — 도메인 patcher 가 cell 몇 개를 덮어쓸 건데, 그 중 일부가
merge 범위 안에 있어 정합성이 깨질 수 있다. 무엇이 깨지는지 미리 보고,
정책으로 해결.

```python
from xp_excel_toolkit.helpers import (
    detect_merge_conflicts, resolve_conflicts, MERGE_POLICIES,
)

writes: dict[tuple[int, int], object] = {
    (5, 3): "new value",       # row=5, col=3
    (5, 4): "another",
    (10, 2): "x",
}

with Session() as s:
    conflicts = detect_merge_conflicts(s, sheet_id=42, writes=writes)
    for c in conflicts:
        print(c)    # MergeWriteConflict: 어느 range, 어느 좌표가 disagree

    # 정책: error / propagate / split / unmerge / origin_wins
    # (MERGE_POLICIES 튜플로 노출)
    final_writes = resolve_conflicts(
        policy="propagate",      # merge 안의 모든 셀에 같은 값
        conflicts=conflicts,
        writes=writes,
        ws=ws,                   # openpyxl Worksheet (export 단계에서 활용)
        session=s,
    )
```

| 정책 | 의미 |
|---|---|
| `error` | 충돌 즉시 raise |
| `propagate` | merge 범위의 모든 cell 에 동일 값 |
| `split` | merge 해제 후 각 셀에 개별 값 |
| `unmerge` | merge 해제 후 origin 위치에만 값 |
| `origin_wins` | merge origin 좌표의 write 만 채택 |

**주의** — `ws` 인자는 정책에 따라 merge 해제 / 셀 분리 필요한 경우에
만 활용. 단순 detect 만 할 거면 ws 없이도 가능 (resolve 만 ws 필요).

---

## Recipe 10 — Change log (변경 audit)

**원함** — 도메인 모델 (Register 등) 의 UPDATE / DELETE 가 자동으로
`change_log` 테이블에 기록.

```python
from xp_excel_toolkit import register_audit_target, init_db

# 도메인 모델 정의 후, init_db 전에 등록
register_audit_target("register", ["name", "indx", "init"])
register_audit_target("memorymap_entry", ["baseaddr", "group"])

Session = init_db("sqlite:///work.db")
# → register / memorymap_entry 의 지정 컬럼에 UPDATE/DELETE trigger 자동
#   생성. change_log 테이블에 (timestamp, table, row_id, column,
#   old_value, new_value, operation) 한 row 씩 기록.
```

**조회**:

```python
from xp_excel_toolkit import ChangeLog

with Session() as s:
    for log in s.query(ChangeLog).order_by(ChangeLog.timestamp.desc()).limit(20):
        print(log.timestamp, log.table_name, log.row_id, log.column_name,
              log.old_value, "→", log.new_value, f"({log.operation})")
```

**주의** — SQLite trigger 기반이라 bulk insert 자체 (INSERT) 는 audit
되지 않음 — UPDATE / DELETE 만. 또 ORM 의 `session.bulk_*` 가 아닌 직
접 SQL UPDATE 도 잡힘.

---

## Recipe 11 — Formula 보존 import + cached value

**원함** — 셀의 formula 식과 그 계산 결과를 둘 다 보존.

```python
from xp_excel_toolkit import init_db, import_xlsx, ExcelCell

Session = init_db("sqlite:///work.db")
with Session() as s:
    import_xlsx(s, "calc.xlsx", with_formulas=True)
    s.commit()

with Session() as s:
    cells_with_formula = (
        s.query(ExcelCell)
        .filter(ExcelCell.raw_value.like("=%"))
        .limit(5)
    )
    for c in cells_with_formula:
        print(f"r{c.row}c{c.col} formula={c.raw_value!r} cached={c.cached_value!r}")
```

**결과** — `raw_value` = `"=A1+B1"` 같은 formula 문자열, `cached_value`
= openpyxl 이 `data_only=True` 로 다시 읽은 마지막 계산 결과 (예:
`"42"`).

**주의** — `with_formulas=True` 는 openpyxl 이 워크북을 두 번 (formula
모드 + value 모드) 로딩하므로 import 시간 약 2배. 평상시엔 기본 (False)
충분.

---

## API reference 요약

### 모델 (`xp_excel_toolkit.models`)

| 심볼 | 역할 |
|---|---|
| `Base` | SQLAlchemy `DeclarativeBase` (cell + 도메인 공유) |
| `ExcelWorkbook` | 파일 메타 + 원본 BLOB |
| `ExcelSheet` | 시트 단위, `header_row` 보존 |
| `ExcelCell` | 셀: `raw_value`, `cached_value`, `style` (dict), `comment`, `formula_type/ref`, `merge_id`, `is_merge_origin` |
| `ExcelMerge` | merge range (min/max row/col) |
| `SheetConfigEntry` | (옵션) sheet 별 import 설정 저장용 row |
| `ChangeLog` | audit trail row |
| `init_db(db_url)` | engine + sessionmaker, `Base.metadata.create_all`, trigger 자동 생성 |
| `register_audit_target(table, columns)` | UPDATE/DELETE trigger 등록 |

### Import (`xp_excel_toolkit.xlsx_parser` / `xls_parser`)

| 함수 | 시그니처 |
|---|---|
| `import_xlsx(session, path, *, on_progress=None, with_formulas=False)` | wb 한 권 → DB. `list[ExcelSheet]` 반환 |
| `import_sheet(session, path, sheet_name=None, header_row=None)` | 한 sheet 만 |
| `import_xls(session, path, sheet_name=None)` | `.xls` 직접 import (xlrd) |
| `extract_cell_value(value)` | openpyxl raw value → `(raw, formula_type, formula_ref)` |
| `extract_style(cell)` | openpyxl Cell → style dict |
| `find_header_row(ws, expected_headers, *, match="any", max_scan=30)` | openpyxl Worksheet 안에서 header row 탐색 |

### 변환 / 캐시 (`xp_excel_toolkit.convert`)

| 함수 | 역할 |
|---|---|
| `convert_xls_to_xlsx(xls_path, output_dir=None, timeout=600)` | LibreOffice 로 `.xls` → `.xlsx` |
| `validate_xlsx_format(path)` | xlsx 무결성 (zip + xl/ 디렉토리) 검증 |
| `cache_key(path)` | `(filename, hash)` 튜플 — 캐시 식별자 |
| `ensure_xlsx_cached(path)` | xlsx 면 그대로 반환, xls 면 캐시 디렉토리에 변환 |
| `resolve_db(path, *, import_fn=None, with_formulas=False)` | xlsx/xls → 캐시된 `.db` path 반환 (없으면 import). `import_fn` 으로 import 함수 교체 가능 |

### Helpers (`xp_excel_toolkit.helpers`)

| 함수 | 역할 |
|---|---|
| `find_header_row_db(session, sheet_id, expected_headers, *, match="any")` | DB 의 `ExcelCell` 만 보고 header row 찾기 |
| `iter_rows_by_header(session, sheet_id, header_row, *, headers=None)` | `(row_idx, {header: value})` iterator |
| `detect_merge_conflicts(session, sheet_id, writes)` | 의도된 write 가 merge 범위 안에서 충돌하는지 검사 |
| `resolve_conflicts(policy, conflicts, writes, ws, session)` | 정책 (`MERGE_POLICIES`) 으로 충돌 해결 |
| `MergeWriteConflict` | dataclass — 어느 range, 어떤 셀들이 disagree |

### Export (`xp_excel_toolkit.exporter`)

| 함수 | 역할 |
|---|---|
| `apply_style(cell, style)` | style dict 를 openpyxl Cell 에 적용 |
| `export_from_cells(session, sheet_id, output_path)` | DB 의 한 sheet → 새 xlsx |

### Diff (`xp_excel_toolkit.diff`)

| 심볼 | 역할 |
|---|---|
| `DiffBase` | diff 테이블들의 공유 declarative base |
| `DiffCell` | 셀 diff row (status, sheet, row, col, old/new value/comment/style/formula/merge_range) |
| `init_diff_db(db_url)` | DiffBase 등록 모두 create_all |
| `load_cells(session)` | `(sheet, row, col) → ExcelCell` |
| `load_cells_by_sheet(session)` | `{sheet: [(row_num, {col: cell}), ...]}` |
| `load_merge_ranges(session)` | `{merge_id: "R..C..:R..C.."}` |
| `cell_display_value(cell)` | cached_value 우선, 없으면 raw_value |
| `cell_formula(cell)` | formula 문자열 또는 None |
| `row_signature(cols)` | SequenceMatcher 용 해시 가능한 tuple |
| `diff_cells(sheet_rows_a, sheet_rows_b, *, compare_comment=False, compare_style=False, compare_merge=False, merges_a=None, merges_b=None)` | smart cell diff → `list[DiffCell]` |

---

## 함께 보면 좋은 자료

- merge 충돌 정책 상세: `xp_excel_toolkit.helpers` 의 docstring
