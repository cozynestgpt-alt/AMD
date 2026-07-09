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

## input 폴더 구조 (2026-07부터: input/YYYY-MM/ 정산월 폴더제)
- **2026-07부터 `input/` 은 정산월별 하위 폴더(`input/YYYY-MM/`, 예: `input/2026-06/`) 안에
  그 달 자료만 넣는 구조로 바뀌었다.** 그 전에는 `input/` 최상위에 이번 달 파일을 두고,
  지난 달 파일은 `input/이전/`에 몰아넣은 뒤 재귀 탐색(glob→rglob→loose 3단계)하는 방식이었는데,
  다른 달 파일과 헷갈릴 여지가 있어 PR #17(`feature/input-monthly-folder-structure`,
  2026-07-09)로 폐지했다. `input/이전/` 폴더는 더 이상 사용하지 않는다.
- `run.py`가 정산월(YYYY-MM)을 입력받으면 `input/YYYY-MM/` 폴더 **안에서만** 파일을 찾는다
  (다른 달 폴더는 절대 보지 않음). `find_input`/`find_input_one`은 `common.py`에 통합되어
  있고, 각 리포트 스크립트(amd_report.py, expense_report.py, sales_summary.py,
  transfer_report.py, tax_invoice_report.py, arba_process.py, management_report.py,
  year_compare_report.py)도 전부 이 구조를 전제로 동작한다.
- **사원마스터.xlsx, 매장전화요금내역.xlsx, 중간관리매장_매입세금계산서_기재사항정보.xlsx**
  같은 "월별 스냅샷이 따로 없던" 마스터성 파일들도 이제 해당 월 폴더 안에
  `사원마스터_202606.xlsx`처럼 연월 접미사를 붙여 그 달 시점 버전으로 관리한다. 즉 과거
  달을 다시 열어도(`run.py`에 예전 YYYY-MM 입력) 그 달 폴더에 저장해둔 시점의 마스터 내용을
  그대로 쓴다 — 단, 2026-07 이전 달(2026-06 이전)은 애초에 월별 스냅샷이 없었으므로
  과거 재조회 시 그 달 당시의 정확한 마스터 내용을 보장하지 못한다.
- 필수 파일(**사원마스터, 영판매(매출)**)이 해당 월 폴더에 없으면 무엇이 빠졌는지 나열하고
  즉시 중단한다(조용히 매출 0 등으로 대체하지 않음). 그 외 파일(경비내역서, 매장공제건집계,
  재고실사, 아르바이트 3종, 매장전화요금내역, 중간관리매장_매입세금계산서_기재사항정보,
  로젠고객직배, 매장구분_행사_신규_폐점)은 없으면 경고만 찍고 0/건너뜀으로 계속 진행한다.
  전체 목록은 `common.py`의 `INPUT_FILE_SPECS` 참고.
- **NAS(`Y:\input\`, `\\nas\CN_AMD\판매수수료 관련자료\salary_system_claude\input\`)도
  이미 동일하게 `input/2026-05/`, `input/2026-06/` 형태의 정산월별 폴더 구조로 자료가
  들어가 있다** (로컬과 별개로 NAS 쪽이 먼저 이 구조를 쓰고 있었음, 2026-07-09 확인).
  앞으로 팀원들이 NAS에 자료를 올릴 때도 이 구조(`input/YYYY-MM/폴더명_YYYYMM.xlsx`)를
  그대로 따르면 되고, 로컬 저장소의 `input/` 구조와 맞춰서 관리한다.

## 알려진 과거 버그 (재발 방지용 기록)
- `management_report.py`: 제목(_title) 작성 직후 `ws.append([])`가 한 번 더 호출되어
  헤더 텍스트가 파란 헤더 배경(fillId=2, #BFE8F5)보다 한 행 아래로 밀리는 버그가 있었음.
  2026-07-08 수정 완료 (PR #5). 비슷한 패턴의 헤더 생성 코드를 새로 작성할 때 주의.
- 판매수수료 계산 로직: 매장코드 중복 정의, 미입점행사(ET_) 매장 분류 회귀,
  직배비공제 3가지 규칙, 롯데아울렛서울역점 누락 등 총 7건 수정 완료 (2026-07-08).

## 보고서 생성 스크립트 목록
- `run.py`: 전체 파이프라인 진입점. 정산 월(YYYY-MM) 입력받아 아래 스크립트들을 순차 실행
- `common.py`: 공통 스타일/색상/매장코드 매핑 모듈 (다른 스크립트들이 import)
- `amd_report.py`: 판매수수료작업시트(AMD) 생성 → `output/판매수수료작업시트_AMD.xlsx`
- `arba_process.py`: 아르바이트 급여 자동 취합 → `output/YYYY-MM/아르바이트_정산서.xlsx`
- `expense_report.py`: 경비지원 및 공제 집계 보고서 → `output/경비지원및공제_집계.xlsx`
- `sales_report.py`: 매장별 매출집계 보고서 → `output/YYYY-MM/매출집계.xlsx`
- `sales_analysis.py`: 매출집계(분석) 보고서 → `output/YYYY-MM/매출집계_분석.xlsx`
- `sales_summary.py`: 판매수수료집계 보고서 → `output/판매수수료집계.xlsx`
- `transfer_report.py`: 중간관리 판매수수료 이체내역 → `output/중간관리판매수수료이체내역.xlsx`
- `tax_invoice_report.py`: 세금계산서 출력 → `output/세금계산서_YYYYMM.xlsx`
- `management_report.py`: 경영분석보고서 생성 → `output/경영분석보고서.xlsx`
- `year_compare_report.py`: 전년대비 판매수수료/손익 보고서 → `output/판매수수료_전년대비보고서.xlsx`
- `history_report.py`: 월별손익DB 기준 연도/분기/시즌 분석 보고서 → `output/연도별_분기별_시즌별_손익분석.xlsx`
- `history_update.py`: `DB/월별손익DB.xlsx` 업데이트 (2026-06부터 매출집계_분석.xlsx 자료 반영)
- 출력 파일은 output/ 폴더에 생성됨 (.gitignore로 git 추적 제외)

## 작업 스타일
- 파일 수정 후에는 openpyxl 등으로 재계산/검증하고, 수식 오류 0건 확인 후 완료 보고할 것
- 커밋 메시지와 PR 설명은 한국어로 작성
