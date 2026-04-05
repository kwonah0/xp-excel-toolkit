# DSM Manual Test Guide

## Setup

모든 명령은 프로젝트 루트(`excel-toolkit/main/`)에서 실행.

```bash
cd ~/1.projects/excel-toolkit/main
```

## 1. 샘플 생성

diff 테스트용 xlsx 두 개를 생성한다.

```bash
uv run python scripts/create_diff_samples.py
```

생성 파일:
- `samples/diff_old.xlsx` — 원본
- `samples/diff_new.xlsx` — 8가지 변경이 적용된 버전

### 알려진 차이점

| # | 시트 | 유형 | 설명 |
|---|------|------|------|
| 1 | level2_common | 값 변경 | SENSOR_A para=0 INIT: `0x00` → `0xAA` |
| 2 | level2_common | 행 삽입 | NEW_REG 추가 (SENSOR_A para=4) |
| 3 | level2_common | 행 삭제 | PLL_CFG para=2 (LOCKED/CAL_DONE) 제거 |
| 4 | level2_common | 코멘트 추가 | Cell (R2,C1) TYPE에 메모 추가 |
| 5 | level2_buscon | 값 변경 | TIMER_A PERIOD INIT: `0xFF` → `0x80` |
| 6 | memorymap | 값 변경 | PLL_CFG comment 텍스트 변경 |
| 7 | memorymap | 행 추가 | WATCHDOG 엔트리 추가 |
| 8 | memorymap | 행 삭제 | PWR_MGMT 엔트리 삭제 |

---

## 2. Cell-level diff (기본)

셀 값 비교. SequenceMatcher 기반 smart diff가 기본.

```bash
uv run dsm diff samples/diff_old.xlsx samples/diff_new.xlsx
```

기대 결과:
- `=== Cells (smart) ===` 섹션 출력
- Added: 14개 (NEW_REG의 14열)
- Removed: 14개 (PLL_CFG para=2의 14열)
- Changed: 값이 바뀐 셀들 (INIT 변경, memorymap 변경 등)
- `diff_diff_old_diff_new.db` 파일 생성

---

## 3. Domain-level diff (--domain)

레지스터와 메모리맵 도메인 모델 비교 추가.

```bash
uv run dsm diff samples/diff_old.xlsx samples/diff_new.xlsx --domain
```

기대 결과:
- `=== Registers ===` 섹션 출력: +1 added, -1 removed, ~2 changed
- `=== MemoryMap ===` 섹션 출력: +1 added, -1 removed, ~1 changed
- `=== Cells (smart) ===` 섹션도 출력

---

## 4. Positional diff (비교용)

행 위치 기반 비교. 행 삽입 시 cascade로 false positive 발생.

```bash
uv run dsm diff samples/diff_old.xlsx samples/diff_new.xlsx --positional
```

기대 결과:
- `=== Cells ===` (smart 표시 없음)
- Changed 수가 smart보다 훨씬 많음 (cascade 효과)

---

## 5. Comment 비교

```bash
uv run dsm diff samples/diff_old.xlsx samples/diff_new.xlsx --comment
```

기대 결과:
- Changed에 `comment:` 라인이 표시됨 (Diff #4)

---

## 6. 전체 비교 (--all)

셀 + 도메인 + 코멘트 + 스타일 + 머지 정보 모두 비교.

```bash
uv run dsm diff samples/diff_old.xlsx samples/diff_new.xlsx --all
```

기대 결과:
- Registers + MemoryMap + Cells 모두 출력
- Cell diff 수가 기본보다 많음 (스타일/머지 차이 포함)

---

## 7. Domain-only (--no-cells)

도메인 모델만 비교하고 셀 비교 제외.

```bash
uv run dsm diff samples/diff_old.xlsx samples/diff_new.xlsx --domain --no-cells
```

기대 결과:
- Registers + MemoryMap만 출력
- Cells 섹션 없음

---

## 8. JSON 출력

```bash
uv run dsm diff samples/diff_old.xlsx samples/diff_new.xlsx --json
```

기대 결과:
- JSON 포맷 출력 (`cell_diffs` 등)

---

## 9. DB 지정

diff 결과를 특정 DB에 저장.

```bash
uv run dsm diff samples/diff_old.xlsx samples/diff_new.xlsx --db /tmp/my_diff.db
```

저장된 DB 확인:
```bash
sqlite3 /tmp/my_diff.db "SELECT * FROM diff_meta;"
sqlite3 /tmp/my_diff.db "SELECT status, COUNT(*) FROM diff_cell GROUP BY status;"
```

---

## 10. DB-to-DB diff

xlsx 대신 미리 import된 DB끼리 비교 (더 빠름).

```bash
# 먼저 import
uv run dsm import samples/diff_old.xlsx --db /tmp/old.db
uv run dsm import samples/diff_new.xlsx --db /tmp/new.db

# DB끼리 diff
uv run dsm diff /tmp/old.db /tmp/new.db
```

---

## 11. 전체 워크플로우 (import → split → merge → diff)

```bash
# import
uv run dsm import samples/diff_old.xlsx --db /tmp/wf_old.db

# split
uv run dsm split --db /tmp/wf_old.db --output-dir /tmp/wf_split/

# (여기서 split된 xlsx를 수정한다고 가정)

# merge (patch)
uv run dsm merge --input-dir /tmp/wf_split/ --base /tmp/wf_old.db --output /tmp/wf_patched.xlsx

# diff
uv run dsm diff samples/diff_old.xlsx /tmp/wf_patched.xlsx
```

---

## Cleanup

테스트 후 생성된 파일 정리:

```bash
rm -f samples/diff_old.db samples/diff_new.db
rm -f diff_diff_old_diff_new.db
```
