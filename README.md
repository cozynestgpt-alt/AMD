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
├── DB
├── backup
├── docs
├── input
├── output
├── templates
│
├── amd_report.py
├── sales_report.py
├── sales_summary.py
├── history_report.py
├── management_report.py
├── year_compare_report.py
│
├── 실행.bat
├── 보고서만생성.bat
├── DB업데이트만.bat
│
├── README.md
├── CHANGELOG.md
└── LICENSE.md
```

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
영업판매.xlsx
사번마스터.xlsx
매장구분_행사_신규_폐점_YYYYMM.xlsx
```

매월 행사/신규/폐점 파일만 교체하면 자동 반영됩니다.

---

# 출력 파일

output 폴더

생성 파일

```
판매수수료작업시트_AMD.xlsx

경영분석보고서.xlsx

판매수수료_전년대비보고서.xlsx

연도별_분기별_시즌별_손익분석.xlsx
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