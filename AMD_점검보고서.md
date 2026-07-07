# AMD V8.0.0 전체 점검 보고서

> 점검일: 2026-07-02  
> 대상: https://github.com/cozynestgpt-alt/AMD  
> 버전: V8.0.0 Final Release

---

## ✅ 실행 가능 여부 판단

**결론: 조건부 실행 가능**

필수 파일이 input 폴더에 준비되면 실행 가능한 상태입니다.  
단, 아래 항목들이 수정되어야 안정적인 운영이 가능합니다.

---

## 🔴 즉시 수정 필요 (Critical)

### 1. `_pu` 변수 scope 오류 — run.py 814번 줄

```python
# 문제 코드
_pu_val = _pu if "_pu" in dir() else 0
```

`_pu`는 try 블록 내부(775번 줄)에서 정의됩니다.  
try가 실패하면 `_pu`가 존재하지 않아 `dir()`로 확인하는 우회법을 썼는데,  
`dir()`는 지역 변수를 포함하지 않아 **항상 0이 반환될 수 있습니다.**

```python
# 올바른 수정
_pu_val = 0  # 기본값을 try 블록 전에 선언
try:
    ...
    _pu = int(...)
    _pu_val = _pu   # try 성공 시에만 갱신
except:
    pass
```

### 2. `config.yaml` 완전 미사용

`config.yaml`이 있지만 어떤 Python 파일도 이를 읽지 않습니다.  
경로, DB 파일명, 시즌 기준 등이 코드에 하드코딩되어 있어  
설정 파일 수정이 실제로 동작에 반영되지 않습니다.

### 3. `load_expense` 중간 import 문제 — run.py 631번 줄

```python
from expense_report import load_expense   # 631번 줄 (try 밖)
```

이 줄은 함수 중간에 위치하여 oepnpyxl 외 의존성 오류 시  
이전 처리 결과물이 저장되지 않은 채 중단될 수 있습니다.

---

## 🟡 개선 권장 (중요)

### 4. `STORE_CODE_MAP` 7개 파일에 중복 정의

동일한 매장코드 딕셔너리가 아래 파일에 각각 하드코딩되어 있습니다:

- `sales_report.py`
- `expense_report.py`
- `sales_analysis.py`
- `sales_summary.py`
- `tax_invoice_report.py`
- `transfer_report.py`
- `run.py` (arba_process 포함 시 8개)

매장이 추가/변경되면 **모든 파일을 수동으로 수정**해야 합니다.  
→ `common.py`로 분리하여 단일 관리 권장

### 5. 스타일 함수 `st()` / `bdr()` 중복 정의

`st()`, `bdr()`, `borders()`, `_st()`, `_bdr()` 등 동일한 역할의 함수가  
파일마다 이름을 달리하여 중복 정의되어 있습니다 (8개 파일).  
→ `common.py`에 통합 권장

### 6. 색상 팔레트 중복 정의

`NAVY`, `BLUE`, `WHITE` 등 색상 상수가 8개 파일에 개별 선언됩니다.  
→ `common.py`에 통합 권장

### 7. `pandas` 의존성 일부만 사용

`pandas`는 `expense_report.py`와 `run.py`의 `load_deduction()`에서만 사용됩니다.  
`openpyxl`로 대체 가능한 부분이 많고, 미설치 시 오류가 발생합니다.  
`requirements.txt`가 없어 설치해야 할 패키지가 명시되어 있지 않습니다.

### 8. `run.py`의 `load_deduction()`과 `expense_report.py`의 `load_deduction_detail()` 역할 중복

두 함수 모두 `◈매장공제건` 파일을 읽지만 약간 다른 컬럼을 파싱합니다.  
`run.py`의 `calc()`에는 `load_deduction()`이,  
`expense_report.py`에는 `load_deduction_detail()`이 별도로 쓰입니다.  
동일 파일을 두 번 읽는 비효율이 발생합니다.

### 9. `실행.bat` 오류 메시지가 영문

```bat
echo Sales commission system V8 Final Release
echo All tasks have been completed successfully.
echo ERROR occurred. Please check the message above.
```

전체 시스템이 한글인데 bat 파일 메시지만 영문입니다.

---

## 🟢 양호한 부분

| 항목 | 상태 |
|------|------|
| 필수 템플릿 3종 모두 존재 | ✅ |
| `.gitignore` — 민감 데이터 제외 설정 완료 | ✅ |
| input/output/backup/DB 폴더 분리 | ✅ |
| 아르바이트 파일 없을 때 건너뜀 처리 | ✅ |
| 매출 파일 없을 때 건너뜀 처리 | ✅ |
| 정산월 입력 형식 자동 변환 (202605 → 2026-05) | ✅ |
| 실행 완료 후 output 폴더 자동 열기 | ✅ |
| `audit_system.py` 점검 도구 존재 | ✅ |
| docs 폴더 운영 문서 정리 | ✅ |
| CHANGELOG / LICENSE / CONTRIBUTING 문서 완비 | ✅ |

---

## 📋 수정 완료 목록 (이번 작업)

| # | 항목 | 파일 |
|---|------|------|
| 1 | `_pu_val` scope 오류 수정 | `run.py` |
| 2 | `실행.bat` 메시지 한글화 | `실행.bat` |
| 3 | `requirements.txt` 신규 생성 | 신규 |
| 4 | `STORE_CODE_MAP` 공통 모듈 분리 | `common.py` 신규 |
| 5 | 공통 스타일 함수/색상 공통 모듈 분리 | `common.py` 신규 |

---

## 📌 V8.0.1 권장 작업 (추가)

- [ ] 모든 `.py` 파일에서 `from common import STORE_CODE_MAP, st, bdr, NAVY, ...` 로 교체
- [ ] `config.yaml` 실제 로딩 코드 추가 (pyyaml)
- [ ] `load_deduction` / `load_deduction_detail` 통합
- [ ] 입력 파일 사전 검증 함수 추가
- [ ] 로그 파일 자동 생성 (`logs/YYYY-MM.log`)
