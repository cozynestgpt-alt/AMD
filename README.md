# AMD
## Analysis & Management Dashboard

> 판매수수료 자동계산 및 백화점/유통 매출 분석 시스템

![Version](https://img.shields.io/badge/Version-V8.0.0-blue)
![Status](https://img.shields.io/badge/Status-Production-success)
![Python](https://img.shields.io/badge/Python-3.10+-yellow)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

# 프로젝트 소개

AMD(Analysis & Management Dashboard)는 더코지네스트컴퍼니의 판매수수료 계산 및 경영분석을 자동화하기 위해 개발된 업무 시스템입니다.

ERP에서 추출한 데이터를 이용하여 판매수수료를 계산하고, 월별 손익, 경영분석, 전년대비 분석 보고서를 자동 생성합니다.

---

# 주요 기능

- 판매수수료 자동 계산
- 판매수수료 작업시트 생성
- 월별 손익 DB 자동 생성
- 연도/분기/시즌별 손익분석
- 경영분석보고서 생성
- 전년대비보고서 생성
- 행사/신규/폐점/동일매장 분석
- Dashboard 자동 생성

---

# 프로젝트 구조

```
AMD
│
├── DB                     # 월별손익DB.xlsx (누적 손익 데이터베이스)
├── backup                 # DB 갱신 전 자동 백업본
├── docs                   # 운영 기준 문서 (저장소_구조.md 포함)
├── input                  # 매월 담당자가 채워 넣는 원본 데이터
├── output                 # 생성된 보고서 결과물
├── templates              # 보고서 서식 템플릿(xlsx)
│
├── run.py                 # 전체 파이프라인 진입점
├── common.py              # 공통 스타일/색상/매장코드 매핑
├── config.yaml            # 프로젝트 설정(KPI 계산식, 시즌/분기 정의 등)
├── audit_system.py        # 템플릿 참조·외부 링크 점검 스크립트
│
├── amd_report.py          # 판매수수료작업시트(AMD)
├── arba_process.py        # 아르바이트 급여 자동 취합
├── expense_report.py      # 경비지원 및 공제 집계
├── sales_report.py        # 매장별 매출집계
├── sales_analysis.py      # 매출집계(분석)
├── sales_summary.py       # 판매수수료집계
├── transfer_report.py     # 중간관리 판매수수료 이체내역
├── tax_invoice_report.py  # 세금계산서 출력
├── management_report.py   # 경영분석보고서
├── year_compare_report.py # 전년대비 판매수수료/손익 보고서
├── history_report.py      # 연도/분기/시즌별 손익분석
├── history_update.py      # 월별손익DB 갱신
│
├── 실행.bat
├── 보고서만생성.bat
├── 전년대비보고서만생성.bat
├── DB업데이트만.bat
│
├── README.md
├── CHANGELOG.md
├── CLAUDE.md
└── LICENSE.md
```

각 스크립트의 상세 입력/출력 파일 경로와 실행 흐름은
[`docs/저장소_구조.md`](docs/저장소_구조.md) 문서를 참고하세요.

---

# 프로그램 실행 순서

## 전체 실행

```
실행.bat
```

실행 내용

1. DB 업데이트
2. 판매수수료 계산
3. AMD 작업시트 생성
4. 월별 손익 DB 생성
5. 경영분석보고서 생성
6. 전년대비보고서 생성

---

## 보고서만 생성

```
보고서만생성.bat
```

생성되는 보고서

- 연도별·분기별·시즌별 손익분석
- 경영분석보고서
- 판매수수료 전년대비보고서

---

## DB만 업데이트

```
DB업데이트만.bat
```

---

# 입력 파일

input 폴더

예)

```
영*판매*.xlsx                        # 영업 판매 파일
사원마스터.xlsx / 사원현황_일용직_YYYY-MM.xlsx
급여상여명세서_매장직_YYYY-MM.xlsx / 급여상여명세서_일용직_YYYY-MM.xlsx
월별근태_일용직_YYYY-MM.xlsx
★*경비내역서*.xlsx / ◈*매장공제건*.xlsx / *재고실사*공제건*합계*.xlsx
*로젠*고객직배*.xlsx / 매장전화요금내역.xlsx
중간관리매장_매입세금계산서_기재사항정보.xlsx
매장구분_행사_신규_폐점_YYYYMM.xlsx
```

매월 행사/신규/폐점 파일만 교체하면 자동 반영됩니다. 전체 파일 목록과 어느
스크립트가 무엇을 읽는지는 `docs/저장소_구조.md`를 참고하세요.

---

# 출력 파일

output 폴더

생성 파일

```
output/YYYY-MM/아르바이트_정산서.xlsx
output/YYYY-MM/매출집계.xlsx
output/YYYY-MM/매출집계_분석.xlsx
output/YYYY-MM/개별통지문.xlsx

output/판매수수료작업시트_AMD.xlsx
output/경비지원및공제_집계.xlsx
output/판매수수료집계.xlsx
output/중간관리판매수수료이체내역.xlsx
output/세금계산서_YYYYMM.xlsx

output/경영분석보고서.xlsx
output/판매수수료_전년대비보고서.xlsx
output/연도별_분기별_시즌별_손익분석.xlsx
```

---

# 경영분석 보고서

생성되는 주요 분석

- 동일매장 분석
- 행사매장 분석
- 신규매장 분석
- 폐점매장 분석

---

# KPI 계산 기준

수금율

```
수금(V+) / 매출
```

원가배수

```
생산원가(V-) / (매출 / 1.1)
```

경비율

```
총경비(V-) / (매출 / 1.1)
```

영업이익률

```
영업이익(V-) / (매출 / 1.1)
```

공헌이익률

```
영업이익 / 전체 영업이익
```

---

# 개발 환경

Python

```
3.10 이상
```

주요 라이브러리

- openpyxl

운영환경

- Windows 10
- Windows 11

---

# GitHub

Repository

```
AMD
```

Release

```
V8.0.0 Final Release
```

---

# 버전 관리

| Version | 내용 |
|----------|------|
| V8.0.0 | 첫 운영 배포 |
| V8.0.1 | 운영 안정화(예정) |
| V9.0 | Enterprise MIS(예정) |

---

# 개발 원칙

- 프로그램보다 데이터를 우선한다.
- 입력 파일 형식을 최대한 유지한다.
- 보고서는 동일한 형식을 유지한다.
- 계산식 변경 시 CHANGELOG를 기록한다.
- Release 생성 후 운영에 반영한다.

---

# 향후 개발 계획

## V8.0.1 Stable

- README 개선
- CHANGELOG 작성
- 입력파일 검증
- 오류메시지 개선
- 로그 시스템
- config 파일 적용

---

## V9 Enterprise MIS

- Web Dashboard
- PDF 자동 생성
- AI 분석
- Scheduler
- Email 자동 발송
- Power BI 연동
- 사용자 권한 관리

---

# 라이선스

Copyright © The Cozynest Company

Internal Use Only

본 프로젝트는 더코지네스트컴퍼니 내부 업무용 프로젝트입니다.

회사 승인 없이 복사, 배포, 수정 및 재판매를 금합니다.

---

# 문의

프로젝트 : AMD (Analysis & Management Dashboard)

Version

```
V8.0.0 Final Release
```

Release Date

```
2026-07-01
```