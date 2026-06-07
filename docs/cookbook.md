# excel_toolkit Cookbook

`.xlsx` / `.xls`를 SQLite로 reify하고, 도메인 모델을 얹은 뒤, **원본 서식을 그대로 유지하며 다시 xlsx로 round-trip** 하는 라이브러리.

핵심 컨셉 4개:

1. **세션 1개 = 워크북 + 시트 + 셀 + (선택) 도메인 모델** 이 모두 들어있는 SQLite DB.
2. 원본 binary가 `excel_workbook.blob`에 박혀서 export 시 서식 보존.
3. 시트 → 도메인 매핑은 **`SheetConfig`** (flat-table은 `field_map + domain_cls`, 비정형은 `parser_func`).
4. 도메인 round-trip은 **`ExportHandler`** + `export_domain_xlsx()`.

---

## 0. 설치

```bash
uv add "excel-toolkit @ git+https://github.com/kwonah0/excel-toolkit.git@main"
```

Python 3.12+. 의존성: `openpyxl`, `sqlalchemy>=2.0`, `xlrd`, `xlwt`. `.xls` 자동 변환을 쓰려면 LibreOffice가 PATH에 있어야 한다.

---

## 1. 데이터 모델 한눈에

| 테이블 | 역할 |
|---|---|
| `excel_workbook` | 원본 파일 binary(`blob`) + filename |
| `excel_sheet` | 시트 이름 + 헤더 행 |
| `excel_merge` | 머지 영역 (min/max row/col) |
| `excel_cell` | 셀 단위 (raw_value, cached_value, style JSON, comment, formula info, merge_id) |
| `sheet_config` | 패턴 → 도메인 매핑 (DB에 영속) |
| `change_log` | UPDATE/DELETE 감사 로그 (트리거로 자동 기록) |

도메인 테이블은 **너가 정의한다**. `excel_toolkit.Base`를 상속하면 같은 `metadata.create_all()`에 묶여서 `init_db()`가 자동으로 만들어 준다.

---

## 2. Recipe: 빈 DB 세션

```python
from excel_toolkit import init_db

Session = init_db("sqlite:///myapp.db")   # 또는 "sqlite:///:memory:"

with Session() as session:
    ...
    session.commit()
```

`init_db()`가 하는 일: (1) 엔진 생성, (2) `Base.metadata.create_all`, (3) `AUDIT_TARGETS`에 등록된 테이블에 SQLite 트리거 생성.

---

## 3. Recipe: 가장 단순한 import — cells만

도메인 모델 없이 그냥 raw cell + merge + style + comment 만 DB에 박고 싶을 때.

```python
from excel_toolkit import init_db, import_xlsx

Session = init_db("sqlite:///myapp.db")
with Session() as session:
    sheets = import_xlsx(session, "input.xlsx", sheet_configs={})
    session.commit()

    for sh in sheets:
        print(sh.name, "header_row=", sh.header_row)
```

> ⚠️ **`sheet_configs={}` 를 꼭 명시.** 생략하면 `SheetConfigEntry` 테이블을 조회해서 등록된 도메인 매핑을 적용한다 — 비어 있으면 무해하지만, 의도치 않은 default가 끼는 걸 막으려면 명시.

---

## 4. Recipe: 자체 도메인 모델 정의

> **왜 `excel_toolkit.Base` 를 상속해야 하는가**: `excel_workbook` / `excel_sheet` / `excel_cell` / `excel_merge` 인프라 테이블과 너의 도메인 테이블이 **같은 `Base.metadata`** 에 들어가야 `init_db()` 가 한 번의 `create_all()` 로 둘 다 만들고, 도메인 테이블이 `excel_sheet.id` 를 FK로 잡을 수 있고, audit 트리거가 같은 엔진 위에 깔린다. 호스트가 별도 `Base` 를 만들면 metadata가 분리되어 이 모든 게 깨진다. **그냥 `from excel_toolkit import Base` 한 줄**.

```python
# myapp/models.py
from sqlalchemy import ForeignKey, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from excel_toolkit import Base, ExcelSheet, register_audit_target


class PinEntry(Base):
    __tablename__ = "pin_entry"

    id: Mapped[int] = mapped_column(primary_key=True)
    pin_no:    Mapped[str | None] = mapped_column(Text)
    name:      Mapped[str | None] = mapped_column(Text)
    direction: Mapped[str | None] = mapped_column(Text)

    sheet_id:  Mapped[int | None] = mapped_column(ForeignKey("excel_sheet.id"))
    excel_row: Mapped[int | None]
    sheet: Mapped[ExcelSheet | None] = relationship()


PIN_FIELD_MAP = {
    "Pin": "pin_no",
    "Name": "name",
    "Dir": "direction",
}

# 감사 로그 트리거 자동 생성 대상으로 등록 (init_db 이전에 실행돼야 함)
register_audit_target("pin_entry", list(PIN_FIELD_MAP.values()))
```

규칙:

- `field_map`은 **Excel 헤더 → 도메인 필드명** 매핑. 헤더 행은 자동 탐지 (field_map 키 중 하나라도 매칭되는 첫 행).
- `sheet_id` + `excel_row`를 반드시 두어야 round-trip export가 원래 행으로 되돌아간다.
- `register_audit_target`은 `init_db()` **이전에** 실행돼야 트리거가 잡힌다 — 보통 `import myapp.models`가 먼저면 자동으로 충족.

---

## 5. Recipe: SheetConfig로 패턴 → 도메인 매핑

```python
from excel_toolkit import init_db, import_xlsx, SheetConfig
from myapp.models import PinEntry, PIN_FIELD_MAP

sheet_configs = {
    # fnmatch 패턴 (대소문자 무시). 정확 매칭 → 패턴 순으로 검색
    "Pinmap_*": SheetConfig(
        field_map=PIN_FIELD_MAP,
        domain_cls=PinEntry,
        # header_row=None  → 자동 탐지
    ),
}

Session = init_db("sqlite:///pinmap.db")
with Session() as session:
    import_xlsx(
        session,
        "chip_spec.xlsx",
        sheet_configs=sheet_configs,
        on_progress=print,     # 진행 로그 (선택)
    )
    session.commit()

    rows = session.query(PinEntry).filter_by(direction="OUT").all()
```

매칭되지 않는 시트는 **infra 테이블(cells/merges)만 채워지고 도메인 행은 안 만들어진다** — 인덱스 시트나 다이어그램 시트가 섞여 있어도 안전.

---

## 6. Recipe: 비-flat 시트는 `parser_func`

key-value, 가변 header, section 구분이 있는 시트는 flat-map으로 표현 불가. parser를 직접 줘서 도메인 행을 채워 넣게 한다.

시그니처:

```python
from openpyxl.worksheet.worksheet import Worksheet
from sqlalchemy.orm import Session

def parse_my_section(session: Session, ws: Worksheet, sheet_id: int) -> None:
    """sheet_id에 묶이는 도메인 행을 직접 insert."""
    ...
```

등록:

```python
sheet_configs = {
    "VoltageDomains": SheetConfig(
        domain_cls=VoltageDomainEntry,
        parser_func=parse_my_section,
    ),
}
```

규칙: parser_func가 있으면 flat-table 로직은 **skip**. 너가 알아서 `excel_row`, `sheet_id`를 채워 insert 해야 한다. bulk insert가 빠르고, `sqlalchemy.insert` + 500개 청크 패턴이 무난.

---

## 7. Recipe: `.xls` 자동 변환 + 캐시

```python
from pathlib import Path
from excel_toolkit import ensure_xlsx_cached, init_db, import_xlsx

# .xls → __xltk__/<stem>_<hash>.xlsx 변환 (이미 있으면 재사용)
xlsx = ensure_xlsx_cached(Path("legacy.xls"))

Session = init_db("sqlite:///myapp.db")
with Session() as session:
    import_xlsx(session, xlsx, sheet_configs=...)
    session.commit()
```

OR 한 줄로 (`.xlsx`/`.xls`/`.db` 모두 처리):

```python
from excel_toolkit import resolve_db

db_path, is_cached = resolve_db(Path("legacy.xls"), on_progress=print)
# db_path → SQLite 경로 (캐시되어 있으면 재사용)
```

`.xls`를 `.xlsx`로 확장자만 바꿔두면 `validate_xlsx_format()`이 OLE2 매직바이트 감지해서 명확한 에러를 던진다.

LibreOffice 경로 자동 탐지 실패 시:

```python
import excel_toolkit.convert
excel_toolkit.convert.LIBREOFFICE_PATH = "/path/to/soffice"
```

---

## 8. Recipe: round-trip export — 원본 서식 보존

```python
from excel_toolkit import ExportHandler, export_domain_xlsx
from myapp.models import PinEntry, PIN_FIELD_MAP

with Session() as session:
    # 도메인 모델 수정
    pin = session.query(PinEntry).filter_by(pin_no="A1").one()
    pin.direction = "OUT"
    session.commit()

    # 원본 BLOB 위에 변경된 셀만 덮어쓰기
    export_domain_xlsx(
        session,
        "modified.xlsx",
        handlers=[
            ExportHandler(
                pattern="Pinmap_*",
                field_map=PIN_FIELD_MAP,
                domain_cls=PinEntry,
            ),
        ],
    )
```

여러 시트 종류를 한 번에 export:

```python
handlers = [
    ExportHandler("Pinmap_*",      field_map=PIN_FIELD_MAP,    domain_cls=PinEntry),
    ExportHandler("Memmap",        field_map=MEMMAP_FIELD_MAP, domain_cls=MemmapEntry),
    ExportHandler("VoltageDomains", exporter_func=write_voltage_domains),
]
export_domain_xlsx(session, "out.xlsx", handlers)
```

- **fnmatch 패턴이 첫 매칭** 되는 핸들러가 적용된다. 정확 매칭이 필요하면 패턴을 정확히.
- `exporter_func` 시그니처: `func(session, ws, sheet_obj) -> int` (쓴 행 수). 비정형 export용.
- non-domain 시트(Notes 등)는 원본 그대로 통과.

원본 BLOB 없이 cells만으로 새 xlsx 생성하려면:

```python
from excel_toolkit import export_from_cells
export_from_cells(session, sheet_id=42, output_path="extracted.xlsx")
```

---

## 9. Recipe: cell 단위 / merge / style / comment

```python
from excel_toolkit import ExcelCell, MergeResolver

# 머지 해상도 (DB에서 복원 — xlsx 파일 없이도 가능)
merger = MergeResolver.from_db(session, sheet_id=42)
if merger.is_merged(row=5, col=2):
    print("origin =", merger.get_origin(5, 2))
    print("value  =", merger.get_value(5, 2))   # origin의 raw_value

# 스타일 JSON 구조 (extract_style이 만든 것)
cell = session.query(ExcelCell).filter_by(sheet_id=42, row=3, col=1).one()
cell.style
# → {"bg_color": "#FFFF00", "font_bold": True, "number_format": "0.00",
#     "border_top": "thin", ...}

# 코멘트
cell.comment   # str | None

# 수식 (with_formulas=True로 import 했을 때만)
cell.raw_value      # "=SUM(B2:B10)"
cell.cached_value   # "42"
cell.formula_type   # None / "array" / "dataTable"
cell.formula_ref    # "B1:B10"
```

`import_xlsx(..., with_formulas=True)` 로 호출하면 워크북을 두 번 로드해서 formula string + cached result를 모두 채운다. 성능 비용 있음.

---

## 10. Recipe: 도메인 row에서 다른 셀 reference

도메인 모델은 자기 출처(`sheet_id` + `excel_row`)를 안다. 거기서 sibling cell이나 임의 셀 address로 reference 잡는 helper는 라이브러리에 내장돼 있지 않고 — 호스트가 짧은 method를 모델/모듈에 붙여 쓰는 패턴이 의도된 사용법.

### 10.1 같은 row의 다른 column

```python
from sqlalchemy.orm import object_session
from excel_toolkit import ExcelCell

class PinEntry(Base):
    ...
    def cell(self, col: int) -> ExcelCell | None:
        session = object_session(self)
        return (
            session.query(ExcelCell)
            .filter_by(sheet_id=self.sheet_id, row=self.excel_row, col=col)
            .one_or_none()
        )

pin = session.query(PinEntry).filter_by(pin_no="A1").one()
note = pin.cell(col=5)
print(note.raw_value, note.comment, note.style)
```

### 10.2 Excel address ("D5") 로 가리키기

```python
from openpyxl.utils import coordinate_from_string, column_index_from_string

def cell_at(entry, address: str) -> ExcelCell | None:
    col_letter, row = coordinate_from_string(address)
    col = column_index_from_string(col_letter)
    session = object_session(entry)
    return (
        session.query(ExcelCell)
        .filter_by(sheet_id=entry.sheet_id, row=row, col=col)
        .one_or_none()
    )
```

시트 간 reference (`Sheet2!D5`):

```python
from excel_toolkit import ExcelSheet

def cell_at_address(session, sheet_id: int, address: str) -> ExcelCell | None:
    if "!" in address:
        sheet_name, coord = address.split("!", 1)
        sheet_id = session.query(ExcelSheet).filter_by(name=sheet_name).one().id
    else:
        coord = address
    col_letter, row = coordinate_from_string(coord)
    col = column_index_from_string(col_letter)
    return (
        session.query(ExcelCell)
        .filter_by(sheet_id=sheet_id, row=row, col=col)
        .one_or_none()
    )
```

### 10.3 머지 origin까지 자동 추적

```python
from excel_toolkit import MergeResolver

def cell_value_resolved(entry, col: int) -> str | None:
    session = object_session(entry)
    merger = MergeResolver.from_db(session, entry.sheet_id)
    if merger.is_merged(entry.excel_row, col):
        return merger.get_value(entry.excel_row, col)
    cell = (
        session.query(ExcelCell)
        .filter_by(sheet_id=entry.sheet_id, row=entry.excel_row, col=col)
        .one_or_none()
    )
    return cell.raw_value if cell else None
```

> `MergeResolver` 생성에 약간 비용이 있어서 시트당 한 번 만들어 재사용 권장.

### 10.4 SQLAlchemy relationship으로 명시적 sibling

자주 쓰는 sibling cell이면 query 시 join / eager-load 가능하게:

```python
from sqlalchemy.orm import relationship

class PinEntry(Base):
    __tablename__ = "pin_entry"
    ...
    sheet_id: Mapped[int | None] = mapped_column(ForeignKey("excel_sheet.id"))
    excel_row: Mapped[int | None]

    # column 5의 셀을 lazy relationship으로
    note_cell: Mapped[ExcelCell | None] = relationship(
        primaryjoin=(
            "and_(PinEntry.sheet_id == ExcelCell.sheet_id, "
            "     PinEntry.excel_row == ExcelCell.row, "
            "     ExcelCell.col == 5)"
        ),
        viewonly=True,
        uselist=False,
    )
```

장점: `selectinload(PinEntry.note_cell)` 같은 eager 최적화 가능. 단점: col index가 모델에 박힘 — 헤더 위치가 바뀌면 깨짐.

### 10.5 헤더 텍스트로 col 동적 찾기

```python
from excel_toolkit import ExcelSheet, ExcelCell

def column_index_for_header(session, sheet_id: int, header_text: str) -> int | None:
    sheet = session.get(ExcelSheet, sheet_id)
    if not sheet or not sheet.header_row:
        return None
    cell = (
        session.query(ExcelCell)
        .filter_by(sheet_id=sheet_id, row=sheet.header_row)
        .filter(ExcelCell.raw_value == header_text)
        .first()
    )
    return cell.col if cell else None
```

`build_column_map(ws, header_row, field_map)`는 worksheet 객체가 손에 있을 때 같은 일을 한 번에 dict로 만들어 준다.

### 10.6 cell address를 도메인 필드로 저장

entry가 다른 셀을 가리키는 게 1급 개념이면 address를 string으로 저장하고 dereference helper로:

```python
class Calculation(Base):
    __tablename__ = "calculation"
    id: Mapped[int] = mapped_column(primary_key=True)
    source_ref: Mapped[str | None] = mapped_column(Text)   # "D5" or "Sheet2!D5"
    sheet_id: Mapped[int | None] = mapped_column(ForeignKey("excel_sheet.id"))
    excel_row: Mapped[int | None]

    @property
    def source_value(self) -> str | None:
        session = object_session(self)
        target = cell_at_address(session, self.sheet_id, self.source_ref)
        return target.raw_value if target else None
```

### 10.7 한계

- Excel 수식 (`=B2+C2`) 은 자동 dereference 안 됨. 문자열 그대로 저장 — 호스트가 직접 parse하거나 `with_formulas=True` 의 `cached_value` 활용.
- R1C1 / structured reference (`Table1[Column]`) 변환 helper 없음.
- 시트 간 address 분해 (`Sheet2!D5` split) 는 호스트가 직접.

### 10.8 정리

| 의도 | 추천 |
|---|---|
| 같은 row 다른 col | 도메인 `cell(col)` 메서드 |
| Excel address ("D5") | `cell_at(addr)` + `openpyxl.utils` |
| sheet-prefixed ("Sheet2!D5") | `cell_at_address(session, sheet_id, addr)` |
| 머지 origin 자동 추적 | `MergeResolver` |
| 자주 쓰는 sibling (관계로) | `relationship(primaryjoin=...)` |
| 헤더로 col 찾기 | `column_index_for_header` / `build_column_map` |
| address 자체를 도메인 필드로 | `source_ref` 컬럼 + dereference 함수 |

---

## 11. Recipe: SQL로 임의 조회/수정 + change_log

`register_audit_target()`을 걸어두면 **UPDATE / DELETE는 모두 `change_log`로 자동 적재**(SQLite 트리거).

```python
from sqlalchemy import text
from excel_toolkit import ChangeLog

# 수정
session.execute(text("UPDATE pin_entry SET direction='RESET' WHERE pin_no='A1'"))
session.commit()

# 감사 로그
recent = session.query(ChangeLog).order_by(ChangeLog.id.desc()).limit(20).all()
for log in recent:
    print(log.timestamp, log.operation, log.table_name, log.row_id,
          log.column_name, log.old_value, "→", log.new_value)
```

undo 한 줄 (UPDATE만):

```python
log = session.get(ChangeLog, 42)
session.execute(
    text(f"UPDATE {log.table_name} SET {log.column_name} = :v WHERE id = :id"),
    {"v": log.old_value, "id": log.row_id},
)
```

---

## 12. Recipe: 시트 패턴을 DB에 영속화 (`SheetConfigEntry`)

런타임에 sheet_configs를 매번 지정하기 싫을 때, DB에 박아두면 `import_xlsx(session, path)` 만 호출해도 자동 적용. 단 **사용자 코드에서 도메인 클래스를 registry에 등록해 둬야** `domain_type` 문자열을 클래스로 풀 수 있다.

```python
import json
from excel_toolkit import SheetConfigEntry
from excel_toolkit.xlsx_parser import register_domain
from myapp.models import PinEntry, PIN_FIELD_MAP

# 패키지 import 시점에 한 번
register_domain("pin_entry", PinEntry, PIN_FIELD_MAP)

# 한 번만 DB에 패턴 영속화
session.add(SheetConfigEntry(
    pattern="Pinmap_*",
    domain_type="pin_entry",
    field_map_json=json.dumps(PIN_FIELD_MAP),
    header_row=None,
    parser_func_ref=None,   # 또는 "myapp.parsers:parse_pinmap"
))
session.commit()

# 이후부터는 sheet_configs 인자 없이도 OK
import_xlsx(session, "input.xlsx")
```

`parser_func_ref`는 `"module:func"` 형태. `importlib`로 동적 로드.

---

## 13. 알아둘 점 / 한계

- 셀 좌표는 모두 **1-based** (openpyxl 규칙). `.xls` 임포트도 내부 변환해서 1-based로 통일.
- `ExcelCell.raw_value`는 **항상 str**. 숫자/날짜도 str로 박힘 — 타입이 중요하면 도메인 필드에서 변환 책임.
- SQLite bulk insert는 `_BULK_CHUNK=500` 청크. 직접 parser_func 쓸 때도 따라 가는 게 안전.
- 캐시 디렉토리는 `cwd/__xltk__/`. 위치를 바꾸려면 `excel_toolkit.convert._CACHE_DIR_NAME`을 import 직후 덮어쓰기.
- `with_formulas=True` 는 워크북을 두 번 로드 — 큰 파일에선 비용 의식.
- `SheetConfig` 매칭은 **정확 매칭 → fnmatch 순**. 같은 패턴 두 번 쓰지 말 것.

---

## 14. API 한 페이지 요약

| 하고 싶은 것 | 부르는 함수 |
|---|---|
| 세션 만들기 (테이블 + 트리거 포함) | `init_db(url)` |
| .xlsx 전체 시트 import | `import_xlsx(session, path, sheet_configs=...)` |
| .xlsx 한 시트만 import | `import_sheet(session, path, sheet_name=..., field_map=..., domain_cls=...)` |
| .xls → .xlsx 캐시 | `ensure_xlsx_cached(path)` |
| 통합 (.db/.xlsx/.xls 알아서 처리) | `resolve_db(path)` |
| 머지 셀 조회 | `MergeResolver.from_db(session, sheet_id)` |
| 원본 BLOB 위에 round-trip export | `export_domain_xlsx(session, out, handlers)` |
| cells만으로 새 xlsx 생성 | `export_from_cells(session, sheet_id, out)` |
| 감사 로그 | `ChangeLog` + `register_audit_target(table, columns)` |
| 시트 매핑 영속화 | `SheetConfigEntry` + `register_domain(type, cls, field_map)` |
| 스타일/border 조작 helper | `apply_style`, `build_column_map`, `write_cell` |

---

## 15. Recipe: facade 패턴으로 dependency 숨기기

호스트 패키지 사용자가 `excel_toolkit`이라는 이름을 한 번도 보지 않게 만들고 싶다면, **명시적 facade 모듈** 하나에 re-export를 모아두자. `__init__.py`는 비워두고, 사용자는 `from pinmap.api import ...` 로 import.

```python
# pinmap/api.py — 호스트의 public 표면
from excel_toolkit import (
    Base,                  # 그대로 alias — 같은 identity, 같은 MetaData
    ChangeLog, ExcelCell, ExcelMerge, ExcelSheet, ExcelWorkbook,
    MergeResolver, init_db,
)

from pinmap.models   import PinEntry, PIN_FIELD_MAP
from pinmap.importer import SHEET_CONFIGS, import_pinmap
from pinmap.exporter import EXPORT_HANDLERS, export_pinmap

__all__ = [
    "Base", "init_db", "PinEntry",
    "import_pinmap", "export_pinmap",
    "ChangeLog", "ExcelWorkbook", "ExcelSheet", "ExcelCell",
    # ...
]
```

```python
# downstream — excel_toolkit이라는 단어를 안 본다
from pinmap.api import init_db, PinEntry, import_pinmap, export_pinmap
```

핵심:

- **`pinmap.api.Base` 는 `excel_toolkit.Base` 와 같은 객체** (alias). 같은 `MetaData` 를 공유하므로 `init_db()` 한 번에 인프라 + 도메인 테이블이 같이 생성된다. SQLAlchemy 2.0의 `DeclarativeBase` 는 un-mapped intermediate subclass 를 허용하지 않으니 alias 가 정답.
- `__init__.py` 를 비워두면 `import pinmap` 만으로 부수효과가 발생하지 않는다 (`register_audit_target` 호출은 `pinmap.api` import 시점까지 지연).
- excel_toolkit 의존을 끊거나 다른 백엔드로 갈아 끼울 때, **변하는 파일은 `pinmap.api` / `models.py` / `importer.py` / `exporter.py` 네 개뿐**.

---

## 16. 실행 가능한 예제

[`examples/pinmap_demo/`](../examples/pinmap_demo/) — 위 §15 의 facade 패턴을 그대로 구현한 작은 호스트 패키지:

```
examples/pinmap_demo/
├── make_sample.py      # 합성 xlsx 생성
├── main.py             # downstream-app 시뮬레이션 (pinmap.api 만 import)
└── pinmap/
    ├── __init__.py     # 비어 있음
    ├── api.py          # ← facade
    ├── models.py       # PinEntry + register_audit_target
    ├── importer.py     # import_pinmap (SheetConfig)
    └── exporter.py     # export_pinmap (ExportHandler)
```

실행:

```bash
uv run python examples/pinmap_demo/main.py
```

import → 도메인 row 조회 → 머지 fill 확인 → 두 컬럼 mutate → `change_log` 출력 → round-trip export → 결과 xlsx 재로드까지 한 번에 보여준다.
