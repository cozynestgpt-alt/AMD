# 프로젝트 메모 (Claude Code가 세션마다 자동으로 읽습니다)

## 이 저장소
- 이름: cozynestgpt-alt/AMD
- 로컬 경로: C:\Users\azmang\Documents\결산\판매수수료\판매수수료_자동계산_시스템\AMD
- 용도: 판매수수료 계산 및 백화점 점포별 매출 분석

## Git 작업 규칙 (중요 — 반드시 지킬 것)
- **main이 유일한 작업 브랜치.** v8.0.1-stable, master, codex/v8.0.1-runtime-stability
  브랜치는 모두 정리되어 삭제됨 (2026-07-08). 더 이상 이 브랜치들을 만들거나 되살리지 말 것.
- **main 브랜치에 직접 push 불가능.** GitHub 브랜치 보호 규칙으로 막혀 있음 (2026-07-08 설정,
  이후 PR 필수로 강화):
  - PR을 통한 병합만 허용 (직접 push 차단, 리뷰 승인 0건이라도 병합 자체는 PR로만 가능)
  - force push 금지, 브랜치 삭제 금지
  - 관리자(본인 계정) 포함 전원에게 동일 적용 (enforce_admins)
- 모든 변경 작업은 아래 순서로 진행:
  1. main에서 새 브랜치 생성 (예: fix/버그이름, feature/기능이름)
  2. 브랜치에서 수정
  3. GitHub에 push
  4. PR(Pull Request) 생성 (`gh pr create`)
  5. 내용 확인 후 병합(Merge) (`gh pr merge`)
  6. 다 쓴 브랜치는 삭제 (원격은 보통 병합 시 자동 삭제됨, 로컬은 `git branch -d`)
- 강제 push(--force)는 절대 사용하지 말 것. 과거 어딘가(다른 로컬 클론/세션으로 추정)에서
  .git 히스토리 없이 재초기화된 뒤 main에 강제 push되어, main과 실제 작업 히스토리
  (v8.0.1-stable/master 계열)가 완전히 갈라지는 사고가 있었고 `--allow-unrelated-histories`
  병합으로 복구했음 (2026-07-08). 지금은 GitHub 쪽에서도 아예 막혀 있지만, 로컬 습관으로도
  사용하지 않는다.
- GitHub CLI(`gh`)가 설치·인증되어 있음. PR/브랜치 보호 규칙 확인·변경은 `gh api`로 가능.

## 백업/복구 참고
- `backup/main-orphan-20260706` 태그: 사고 이전 main의 원래 커밋(afe85da) 보존용
- `v8.0.1-stable-snapshot` 태그: v8.0.1-stable 브랜치를 태그로 전환하며 남긴 마지막 상태
- 문제가 생기면 이 태그들을 참고해서 복구 가능

## 알려진 과거 버그 (재발 방지용 기록)
- `management_report.py`: 제목(_title) 작성 직후 `ws.append([])`가 한 번 더 호출되어
  헤더 텍스트가 파란 헤더 배경(fillId=2, #BFE8F5)보다 한 행 아래로 밀리는 버그가 있었음.
  2026-07-08 수정 완료 (PR #5). 비슷한 패턴의 헤더 생성 코드를 새로 작성할 때 주의.
- 판매수수료 계산 로직: 매장코드 중복 정의, 미입점행사(ET_) 매장 분류 회귀,
  직배비공제 3가지 규칙, 롯데아울렛서울역점 누락 등 총 7건 수정 완료 (2026-07-08).

## 보고서 생성 스크립트 목록
- amd_report.py, arba_process.py, common.py, expense_report.py, management_report.py,
  run.py, sales_analysis.py, sales_report.py, sales_summary.py, tax_invoice_report.py,
  year_compare_report.py
- 출력 파일은 output/ 폴더에 생성됨 (.gitignore로 git 추적 제외)

## 작업 스타일
- 파일 수정 후에는 openpyxl 등으로 재계산/검증하고, 수식 오류 0건 확인 후 완료 보고할 것
- 커밋 메시지와 PR 설명은 한국어로 작성
