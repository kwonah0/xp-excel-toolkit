---
description: "DSM (Design Specification Manager) — Register map Excel 관리 CLI. Import, Split, Diff, SQL 조회/수정, Merge/Export. 트리거: dsm, register map, regmap, excel import, 레지스터 맵"
---

# DSM — Design Specification Manager

Register map Excel 파일(.xlsx/.xls)을 SQLite DB로 관리하는 CLI 도구. Import, Split, Diff, SQL 조회/수정, Merge/Export 워크플로우를 제공한다.

## 프로젝트 경로

```
DSM_ROOT=~/1.projects/excel-toolkit/main
```

## CLI 명령어

```bash
# 활성화
cd $DSM_ROOT
source .venv/bin/activate

# Import: Excel → DB
dsm import <excel.xlsx>                    # → __dsm__/<stem>_<hash>.db (자동)
dsm import <excel.xlsx> --db <output.db>   # → 지정 경로에 DB 생성
dsm import <excel.xls>                     # .xls → auto-convert via LibreOffice

# Split: IP별 분리
dsm split <excel.xlsx>                     # → <stem>_split/ 디렉토리
dsm split --db <file.db>                   # DB만으로도 가능

# Diff: 두 파일 비교
dsm diff old.xlsx new.xlsx                 # → diff_*.csv 자동 생성 + 요약 stdout
dsm diff old.db new.db --format text       # 텍스트 전체 출력
dsm diff old.db new.db --format json       # JSON 출력
dsm diff old.db new.db --all               # cell + domain + comment + style + merge 전부

# SQL 조회/수정 (--db에 xlsx도 가능 → 자동 import)
dsm sql "SELECT * FROM register" --db <file.xlsx>
dsm sql "SELECT * FROM register" --db <file.db> --json
dsm sql "UPDATE register SET type='RW1' WHERE id=1" --db <file.db>

# Query 명령 (--db에 xlsx도 가능)
dsm query sheets --db <file.xlsx>
dsm query ips --db <file.db>
dsm query registers --db <file.xlsx>
dsm query registers --db <file.db> --ip SENSOR_A --json
dsm query memmap --db <file.db>

# Merge: 분리 파일 재결합
dsm merge --input-dir design_split/ --output merged.xlsx
dsm merge --input-dir design_split/ --base original.db --output patched.xlsx

# Change Log (감사 추적)
dsm log show --db <file.db>                   # 최근 20개 변경 이력
dsm log show --db <file.db> --table register  # 특정 테이블만
dsm log show --db <file.db> --last 50         # 최근 50개
dsm log undo 42 --db <file.db>                # 변경 되돌리기 (UPDATE만)
dsm log clear --db <file.db>                  # 전체 이력 삭제 (확인 필요)

# Config 관리
dsm config list --db <file.db>
dsm config add --db <file.db> --pattern "level2_*" --domain register
dsm config reset --db <file.db>
```

## DB Schema

현재 스키마는 [schema.md](schema.md) 참조.

스키마가 변경되면 codegen으로 재생성:
```bash
python scripts/codegen.py <excel.xlsx> --auto --apply --schema-doc skill/dsm-skill/schema.md
```

## 워크플로우

### 1. Import + 조회

```bash
# Excel → DB (xlsx를 직접 --db에 넘기면 자동 import)
dsm query sheets --db regmap.xlsx
dsm query ips --db regmap.xlsx
dsm query registers --db regmap.xlsx --ip SENSOR_A

# SQL 조회 (유연한 쿼리)
dsm sql "SELECT name, type, indx, init FROM register WHERE type='RW2'" --db regmap.xlsx
dsm sql "SELECT * FROM overview_entry WHERE category='General Option'" --db regmap.xlsx
dsm sql "SELECT baseaddr, group, comment FROM memorymap_entry" --db regmap.xlsx
```

### 2. Split → 수정 → Merge

```bash
# IP별 분리
dsm split regmap.xlsx

# 분리된 파일 수정 후 재결합 (patch merge — 원본 서식 보존)
dsm merge --input-dir regmap_split/ --base regmap.db --output regmap_patched.xlsx
```

### 3. Diff (파일 비교)

```bash
# 기본: CSV 파일 생성 + 요약
dsm diff old.xlsx new.xlsx

# 텍스트 상세 출력
dsm diff old.db new.db --format text

# 도메인 레벨 비교 포함
dsm diff old.db new.db --domain

# 전체 비교 (cell + domain + comment + style + merge)
dsm diff old.db new.db --all
```

### 4. SQL 수정 + 감사 추적

```bash
# 값 수정 (자동으로 change_log에 기록됨)
dsm sql "UPDATE register SET init='0xFF' WHERE name='SENSOR_A' AND para='0'" --db regmap.db

# 변경 이력 확인
dsm log show --db regmap.db

# 실수 되돌리기
dsm log undo 1 --db regmap.db

# Export (DB → xlsx, 원본 서식 유지)
dsm merge --input-dir regmap_split/ --base regmap.db --output modified.xlsx
```

### 5. Code Generation (스키마 변경 시)

Excel 컬럼 구조가 달라지면 codegen을 다시 실행:

```bash
# Auto 모드: level2_*와 memorymap 시트 자동 탐지 → domain_models.py 갱신 + schema.md 재생성
python scripts/codegen.py <new_excel.xlsx> --auto --apply --schema-doc skill/dsm-skill/schema.md

# 수동 모드: 특정 시트만
python scripts/codegen.py <excel.xlsx> --sheet level2_common --apply
```

## 시트 구조

| 패턴 | Domain Model | 설명 |
|------|-------------|------|
| `level2_*` | `Register` | 레지스터 맵 (TYPE, INDX, PAGE, PARA, NAME, D7..D0, INIT) |
| `memorymap` | `MemoryMapEntry` | 메모리 맵 (BASEADDR, Group, midgroup, Comment, special) |
| `overview` | `OverviewEntry` | Key-value 설정 (#카테고리 구분) |

## 주의사항

- DB는 SQLite 파일. `dsm sql`로 임의의 SQL 실행 가능.
- `--db`에 `.xlsx`/`.xls`를 직접 넘기면 `__dsm__/`에 자동 import 후 쿼리.
- `.xls` 파일은 LibreOffice를 통해 자동 `.xlsx` 변환 (`__dsm__/`에 캐시).
- `.xls`를 `.xlsx`로 확장자만 바꾸면 에러 발생 — 반드시 원래 확장자 유지.
- Diff 기본 출력은 CSV 파일 자동 생성 + stdout 요약.
- `--json` 플래그로 JSON 출력 가능 (파이프라인 연동 시 유용).
- Config 패턴은 fnmatch 문법 (`*`, `?`, `[seq]`).
- Patch merge는 원본 Excel BLOB 위에 변경분만 덮어쓰기 (서식/머지/스타일 보존).
- `dsm sql`로 UPDATE/DELETE 시 `change_log` 테이블에 자동 기록 (SQLite 트리거).
- Cache DB(`.xlsx` 기반 자동 생성)는 원본 변경 시 재생성되므로 audit도 초기화됨.
