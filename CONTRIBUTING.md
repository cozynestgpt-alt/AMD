# AMD Development Guide

## 프로젝트 목적

AMD(Analysis & Management Dashboard)는
더코지네스트컴퍼니의 판매수수료 계산 및 경영분석 자동화를 위한 내부 프로젝트입니다.

본 문서는 AMD 프로젝트의 개발 원칙과 운영 절차를 정의합니다.

---

# 개발 원칙

## 1. 운영 안정성 우선

새로운 기능보다 기존 기능의 안정성을 우선합니다.

운영 중인 기능은 충분한 테스트 후 변경합니다.

---

## 2. 데이터 우선

입력 데이터 형식은 가능한 변경하지 않습니다.

기존 엑셀 양식과의 호환성을 유지합니다.

---

## 3. 보고서 형식 유지

기존 보고서의 구조는 유지합니다.

필요 시 새로운 시트를 추가하는 방식으로 개선합니다.

---

## 4. 버전 관리

모든 변경은 GitHub를 통해 관리합니다.

Commit 메시지는 변경 내용을 명확하게 작성합니다.

예시

V8.0.1-004 .gitignore 추가

V8.0.1-005 Architecture 문서 추가

---

## 5. 브랜치 전략

main

운영 버전

v8.x-stable

안정화 작업

feature/*

신규 기능 개발

hotfix/*

운영 긴급 수정

---

# Commit 규칙

좋은 예

V8.0.1-006 경영보고서 오류 수정

V8.0.1-007 README 개선

나쁜 예

update

fix

수정

---

# 코드 작성 원칙

Python

PEP8을 기본으로 합니다.

함수는 하나의 기능만 수행하도록 작성합니다.

중복 코드는 함수로 분리합니다.

Batch

UTF-8 사용

상대경로 사용

Python

하드코딩 최소화

주석 작성

예외 처리 구현

---

# 테스트

변경 후 반드시 다음 BAT를 테스트합니다.

실행.bat

보고서만생성.bat

DB업데이트만.bat

---

# Release 규칙

운영 버전은 GitHub Release를 생성합니다.

예)

V8.0.0

V8.0.1

V9.0.0

---

# 문의

Project

AMD

Repository

AMD

Company

The Cozynest Company