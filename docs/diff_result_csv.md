# Diff Result CSV Format

`dsm diff` 명령의 `--format csv` 출력 포맷 명세.

## 생성 방법

```bash
# 직접 출력
dsm diff old.xlsx new.xlsx --format csv

# 파일로 저장
dsm diff old.xlsx new.xlsx --format csv -o diff_result.csv

# 확장자로 자동 감지
dsm diff old.xlsx new.xlsx -o diff_result.csv
```

## 컬럼 정의

| 컬럼 | 타입 | 설명 |
|------|------|------|
| `status` | string | 변경 유형: `added`, `removed`, `changed`, `moved` |
| `sheet` | string | 시트 이름 |
| `old_row` | int | 이전 파일의 행 번호 |
| `new_row` | int | 새 파일의 행 번호 |
| `col` | string | Excel 스타일 컬럼 문자 (A, B, ..., Z, AA, AB, ...) |
| `old_value` | string | 이전 셀 값 (결과값) |
| `new_value` | string | 새 셀 값 (결과값) |
| `old_formula` | string | 이전 수식 (`--with-formulas`로 import 시) |
| `new_formula` | string | 새 수식 (`--with-formulas`로 import 시) |
| `old_comment` | string | 이전 코멘트 (`--comment` 사용 시) |
| `new_comment` | string | 새 코멘트 (`--comment` 사용 시) |
| `old_style` | string | 이전 스타일 JSON (`--style` 사용 시) |
| `new_style` | string | 새 스타일 JSON (`--style` 사용 시) |

## Status 값

### `added`
새 파일에만 존재하는 셀. `old_value`는 비어 있음.

```csv
status,sheet,old_row,new_row,col,old_value,new_value,...
added,Sheet1,,5,A,,REG_X,...
added,Sheet1,,5,B,,0x10,...
```

### `removed`
이전 파일에만 존재하는 셀. `new_value`는 비어 있음.

```csv
status,sheet,old_row,new_row,col,old_value,new_value,...
removed,Sheet1,3,,A,REG_Y,,...
removed,Sheet1,3,,B,0x04,,...
```

### `changed`
양쪽 모두 존재하지만 값이 다른 셀.

```csv
status,sheet,old_row,new_row,col,old_value,new_value,...
changed,Sheet1,2,2,C,8,16,...
changed,Sheet1,7,7,A,REG_OLD,REG_NEW,...
```

### `moved`
값은 동일하지만 행 위치가 변경된 셀. `old_row`와 `new_row`가 다름.

```csv
status,sheet,old_row,new_row,col,old_value,new_value,...
moved,Sheet1,3,5,A,REG_B,REG_B,...
moved,Sheet1,3,5,B,0x04,0x04,...
```

## 도메인 레벨 diff

`--domain` 옵션 사용 시 Register/MemoryMap 변경도 포함됨. `old_row`, `new_row`, `col`은 비어 있고, 값 컬럼에 요약 정보가 들어감.

```csv
added,Sheet1,,,,,"REG:NEW_REG indx=0x10",...
removed,Sheet1,,,,"REG:OLD_REG indx=0x08",,...
changed,Sheet1,,,,"REG:MY_REG","REG:MY_REG [init:0x00->0xFF]",...
```

## 전체 예시

```csv
status,sheet,old_row,new_row,col,old_value,new_value,old_formula,new_formula,old_comment,new_comment,old_style,new_style
changed,level2_IP_A,10,10,D,RW,RO,,,,,,
changed,level2_IP_A,10,10,E,150,165,'=SUM(A1:A10),'=SUM(A1:A11),,,,
added,level2_IP_A,,12,A,,NEW_FIELD,,,,,,
removed,level2_IP_B,5,,B,OLD_VAL,,,,,,,
moved,level2_IP_B,8,3,A,REG_X,REG_X,,,,,,
moved,level2_IP_B,8,3,B,0x20,0x20,,,,,,
```

## 관련 옵션

| 옵션 | 설명 |
|------|------|
| `--with-formulas` | 수식 포함 import + diff 시 수식 표시 (old_formula/new_formula 채워짐) |
| `--comment` | 코멘트 비교 활성화 (old_comment/new_comment 채워짐) |
| `--style` | 스타일 비교 활성화 (old_style/new_style 채워짐) |
| `--merge` | 병합 셀 비교 활성화 |
| `--positional` | 위치 기반 비교 (smart diff 비활성화, moved 감지 안 됨) |
| `--save-db` | diff 결과를 SQLite DB로 저장 (기본: 저장 안 함) |
| `--limit N` | 출력 행 수 제한 (CSV에는 적용 안 됨) |
