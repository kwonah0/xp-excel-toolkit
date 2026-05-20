# excel-toolkit

## 개요

Excel(.xlsx/.xls) ↔ SQLite round-trip 인프라 라이브러리. 워크북/시트/셀(스타일·머지·코멘트 포함)을 SQLAlchemy 스키마로 영속화하고, 원본 binary를 `excel_workbook.blob`에 박아 두어 export 시 서식을 보존한다.

도메인 모델은 호스트 패키지(예: pinmap)가 정의. `excel_toolkit.Base`를 상속한 ORM 클래스 + `SheetConfig` 등록 + `ExportHandler` 로 round-trip 워크플로우를 완성한다.

- 사용법: [docs/cookbook.md](docs/cookbook.md)

## 구조

bare+worktree 레이아웃 (dw/workflower 패턴):

- `../.bare/` — bare repository
- `../.git` — `gitdir: ./.bare` 포인터
- `main/` — main worktree, **여기가 작업 스코프**

소스:

- `src/excel_toolkit/`
  - `models.py` — `Base`, `init_db`, `ExcelWorkbook/Sheet/Cell/Merge`, `ChangeLog`, `SheetConfigEntry`, audit trigger 생성
  - `xlsx_parser.py` — `import_xlsx`, `import_sheet`, `SheetConfig`, `find_header_row`, `extract_style`, `register_domain`
  - `xls_parser.py` — `import_xls` (xlrd 기반)
  - `convert.py` — `.xls → .xlsx` LibreOffice 변환, `__xltk__/` 캐시, `resolve_db`, `validate_xlsx_format`
  - `merge.py` — `MergeResolver` (worksheet 또는 DB에서 해석)
  - `exporter.py` — `export_domain_xlsx` (원본 BLOB 위에 도메인 round-trip), `export_from_cells` (cells에서 새 xlsx 재구성), `ExportHandler`

## 개발 명령어

```bash
# 의존성 동기화
uv sync

# 테스트
uv run pytest

# 패키지 import 확인
uv run python -c "import excel_toolkit; print(excel_toolkit.__all__)"
```
