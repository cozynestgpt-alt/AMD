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
- **예외**: 작업요약_*.md 같은 순수 기록용 문서 파일만 변경하는 경우, 브랜치→PR
  절차는 유지하되 병합 전 사용자 확인 없이 바로 병합까지 진행한다. 코드/수식/차트
  등 기능에 영향을 주는 변경이 포함된 PR은 기존대로 반드시 diff 확인 후 병합한다.
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

## NAS 운영 방식 (중요 — 코드 배포 규칙)
- **코드의 진짜 버전(원본)은 이 PC(로컬 저장소) + GitHub main 브랜치에만 있다.**
  NAS(`Y:\`, UNC 경로 `\\nas\CN_AMD\판매수수료 관련자료\salary_system_claude`)는
  **git 저장소가 아니다.** 2026-07-09 이전에는 NAS에 `.git` 폴더가 통째로 남아있었고,
  이미 정리된 오래된 `master` 브랜치(afe85da 이전 계열)에 체크아웃된 채 방치되어 있었다.
  확인 결과 origin/main 대비 잃어버릴 새 작업이 전혀 없어 `.git` 폴더를 완전히
  제거했다. NAS를 다시 git으로 되돌리지 말 것.
- **NAS는 순수 실행 전용 폴더다.** 팀원들은 NAS의 `input/YYYY-MM/` 폴더에 그 달 자료만
  넣고, `실행.bat`/`DB업데이트만.bat`/`보고서만생성.bat`/`전년대비보고서만생성.bat` 등
  배치 파일을 더블클릭해 결과(`output/`)만 받는다. **팀원은 코드(.py)를 절대 건드리지
  않는다.**
- **input/output 실무 데이터의 실제 위치는 NAS다.** 로컬 저장소의 input/output은
  Claude Code 세션에서 코드 변경을 격리 테스트하거나 실제 실행을 검증할 때 쓰는 것이고,
  팀이 매달 실제로 쓰는 자료·산출물은 NAS 기준이다.
- **코드 수정 흐름**: 이 PC에서 코드 수정 → 브랜치 생성 → PR → main에 병합 → 병합된
  최신 코드를 저장소 루트의 `1_배포_NAS로_동기화.bat` 실행으로 NAS에 반영한다. 이 배치는
  git이 아니라 **순수 파일 복사(robocopy)**로 `*.py`, `*.bat`, `requirements.txt`,
  `templates/`만 NAS에 덮어쓰고, `input/output/DB/backup`은 절대 건드리지 않는다
  (2026-07-09 실제 배포 실행으로 데이터 폴더 미접근 확인 완료).

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
- `history_update.py`의 `read_analysis()`: `output/YYYY-MM/매출집계_분석.xlsx`에서
  매장별 데이터를 읽을 때 요약행(합계행) 제외 필터가 매장명(B열)에 '합계'/'총계' 문자열이
  있는지만 검사했는데, `sales_analysis.py`가 만드는 실제 합계행은 매장코드(A열)="합 계"
  (중간에 공백), 매장명(B열)="(65개 매장)" 형태라 필터를 통과해 `DB/월별손익DB.xlsx`에
  그대로 섞여 들어가는 버그가 있었음. 매장코드는 항상 5자리 숫자라는 불변조건으로 필터를
  보강해 2026-07-09 수정 완료 (PR #20). DB에서 오염된 행(9846행, 2026-06 합계행)도 백업
  후 제거함. 앞으로 매출집계/DB 관련 요약행 필터링 코드를 새로 작성할 때, 매장명뿐 아니라
  매장코드(항상 숫자)도 함께 검증하는 방식을 우선 고려할 것.

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
- `management_report.py`: 경영분석보고서 생성 → `output/경영분석보고서.xlsx` (+ `output/YYYY-MM/`에도 월별 사본)
- `year_compare_report.py`: 전년대비 판매수수료/손익 보고서 → `output/판매수수료_전년대비보고서.xlsx`
- `history_report.py`: 월별손익DB 기준 연도/분기/시즌 분석 보고서 → `output/연도별_분기별_시즌별_손익분석.xlsx` (+ `output/YYYY-MM/`에도 월별 사본)
- `history_update.py`: `DB/월별손익DB.xlsx` 업데이트 (2026-06부터 매출집계_분석.xlsx 자료 반영)
- 출력 파일은 output/ 폴더에 생성됨 (.gitignore로 git 추적 제외)

### run.py 호출 체인 (중요 — management_report.py 등은 "직접"이 아니라 2단 간접 호출됨)
`run.py`의 `main()`은 amd_report/arba_process/expense_report/sales_report/
sales_analysis/sales_summary/transfer_report/tax_invoice_report는 **직접** 호출하지만,
`management_report.py`/`year_compare_report.py`/`history_report.py` 3개는 run.py
본문에 이름이 등장하지 않는다. 대신 run.py 맨 마지막에 호출하는
`history_update.run(ym, BASE)` **내부에서** 다음 순서로 체인 호출된다
(`history_update.py`의 `run()` 함수, V8 최초 커밋 824a3e3부터 있던 원래 설계 —
빠뜨린 버그 아님):

```
run.py: history_update.run(ym, BASE)
  └─ update_history(ym, base_dir)          # DB 업데이트. try 밖 — 여기서 실패하면 아래 전부 스킵
  └─ (성공 시) history_report.run(base_dir, ym)     # 실패하면 아래 두 개도 같이 스킵됨(같은 try 블록)
       └─ management_report.run(base_dir, ym)       # 독립 try/except — 실패해도 아래는 실행됨
       └─ year_compare_report.run(ym, base_dir)     # 독립 try/except
```

즉 `실행.bat`(`python run.py`) 한 번으로 경영분석보고서/전년대비보고서/손익분석
보고서까지 전부 갱신된다. 단, DB 업데이트(`update_history`)나 `history_report.run()`
자체가 예외를 던지면 그 아래 체인이 통째로 스킵되므로, "실행.bat을 돌렸는데
경영분석보고서.xlsx가 그대로다"라면 콘솔에 `⚠️ 연도별/분기/시즌 보고서 생성 오류`
같은 메시지가 없었는지부터 확인할 것. management_report.py 소스만 보고 "run.py가
호출 안 하니 반영 안 됨"이라고 단정하지 말 것 — run.py 본문 텍스트에 이름이
없다고 호출되지 않는 게 아니라, history_update.py를 거쳐 간접 호출된다
(2026-07-10 이 문서와 실제 코드가 불일치하는 것처럼 보여 조사한 결과 확인).

### .bat 파일별 실제 동작 (경영분석보고서.xlsx 갱신 여부 포함)
| 파일 | 실행 내용 | 경영분석보고서.xlsx 갱신? |
|---|---|---|
| `실행.bat` | `python run.py` (전체 파이프라인, 위 체인 포함) | 갱신됨 |
| `DB업데이트만.bat` | `python history_update.py` (ym 입력받아 위 체인 중 update_history~year_compare_report 실행) | 갱신됨 — 이름은 "DB만"이지만 실제로는 3개 보고서도 같이 재생성됨 |
| `보고서만생성.bat` | `history_report.py`→`management_report.py`→`year_compare_report.py`를 각각 독립 실행 (DB 업데이트 없이 보고서만 재생성하고 싶을 때) | 갱신됨 |
| `전년대비보고서만생성.bat` | `python year_compare_report.py`만 | 갱신 안 됨 (전년대비보고서만 갱신) |
| `1_배포_NAS로_동기화.bat` | robocopy로 `*.py *.bat requirements.txt templates/`를 NAS로 복사 (코드 배포 전용, input/output/DB 미접근) | 해당 없음 |

## 작업 스타일
- 파일 수정 후에는 openpyxl 등으로 재계산/검증하고, 수식 오류 0건 확인 후 완료 보고할 것
- 커밋 메시지와 PR 설명은 한국어로 작성
