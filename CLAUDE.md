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

## NAS 운영 방식 (중요 — 코드 배포 규칙, 2026-07-13 git pull 방식으로 재전환)
- **코드의 진짜 버전(원본)은 GitHub main 브랜치다.** NAS(`Y:\`, UNC 경로
  `\\nas\CN_AMD\판매수수료 관련자료\salary_system_claude`)는 이제 **다시 git
  저장소다** — 아래 이력/이유 참고. `origin`은 `https://github.com/cozynestgpt-alt/AMD.git`,
  `main` 브랜치를 체크아웃한 상태로 유지한다.
- **이력 (왜 "git 아님" → "다시 git"으로 바뀌었나)**: 2026-07-09 이전에는 NAS에
  `.git` 폴더가 통째로 남아있었지만, 이미 정리된 오래된 `master` 브랜치
  (afe85da 이전 계열)에 체크아웃된 채 방치되어 있었고 아무도 관리하지 않아서
  당시(PR #21) `.git` 폴더를 완전히 제거하고 "NAS는 git 저장소가 아니다,
  robocopy로만 배포한다"로 정책을 정했다. 그런데 robocopy 배포는 매번 사람이
  이 PC에서 `1_배포_NAS로_동기화.bat`을 수동 실행해야 하는 번거로움이 있어서,
  2026-07-13(PR #36 및 그 직전 커밋 `0a659c7`)에 이 결정을 의도적으로
  재검토해서 뒤집었다. 팀원용 실행 bat들이 python을 호출하기 전에 내부 헬퍼
  `_auto_git_pull.bat`을 통해 **매번 자동으로 `git pull --ff-only`** 를 실행하는
  방식으로 바뀌었고, 아래 안전장치를 설계해 이번엔 방치될 위험을 줄였다:
  - `.pulling.lock` 파일로 동시 실행(여러 팀원이 동시에 bat 실행) 시 pull이
    중복/충돌하지 않도록 막는다(`.gitignore`에도 등록되어 있어 커밋되지 않음).
  - `git pull --ff-only`만 사용 — fast-forward가 안 되는 상황(NAS 쪽에 로컬
    커밋이 쌓이는 등 예상 밖의 분기)이면 그냥 실패하고 멈춘다(강제 병합·리셋
    안 함).
  - pull이 실패해도(네트워크 문제 등) 팀원의 작업 자체를 막지 않는다 — 기존
    코드로 계속 진행되고, 실패 알림만 PowerShell `-EncodedCommand`로 한글
    표시한다(`_auto_git_pull.bat`, [[cp949_조사결과_및_결정]] 참고).
  - `input/output/DB/backup`과 NAS 로컬 전용 실무 파일(예: `templates/` 안의
    산출물 2건)은 `.gitignore`에 등록되어 있어 `git pull`이 절대 건드리지 않는다.
  - 2026-07-13 실제로 NAS에서 `git pull --ff-only` 수동 실행해 `ea7ffb2` →
    `37acf8e`(PR #36 포함)까지 fast-forward 성공, `input/output/DB/backup`
    미접근 확인 완료.
- **NAS는 여전히 순수 실행 전용 폴더다.** 팀원들은 NAS의 `input/YYYY-MM/` 폴더에
  그 달 자료만 넣고, `실행.bat`/`DB업데이트만.bat`/`보고서만생성.bat`/
  `전년대비보고서만생성.bat` 등 배치 파일을 더블클릭해 결과(`output/`)만 받는다.
  **팀원은 코드(.py)를 절대 건드리지 않는다** — git 저장소가 됐다고 해서 팀원이
  직접 git 명령을 쓰거나 커밋하는 것은 아니다. 자동 pull은 bat 실행 시 내부적으로만
  일어난다.
- **input/output 실무 데이터의 실제 위치는 NAS다.** 로컬 저장소의 input/output은
  Claude Code 세션에서 코드 변경을 격리 테스트하거나 실제 실행을 검증할 때 쓰는 것이고,
  팀이 매달 실제로 쓰는 자료·산출물은 NAS 기준이다.
- **코드 수정 흐름**: 이 PC에서 코드 수정 → 브랜치 생성 → PR → main에 병합. 이후
  NAS 쪽은 팀원이 실행 bat을 더블클릭하는 순간 `_auto_git_pull.bat`이 자동으로
  최신 main을 반영한다(수동 배포 스크립트 불필요해짐). `1_배포_NAS로_동기화.bat`
  (robocopy 방식)은 `docs/legacy/1_배포_NAS로_동기화.bat.old`로 이관되어 더 이상
  쓰지 않는다.
- **✅ 실기기 검증 완료 (2026-07-13)**: `_auto_git_pull.bat`(현재는
  `_pull_retry.ps1`)의 한글 알림 메시지가 실제 한국어 Windows 콘솔 화면에
  깨지지 않고 정상 출력되는 것을, NAS를 통해 실제 팀원 PC에서 `실행.bat`
  더블클릭으로 확인했다. 자세한 내용은 `docs/legacy/cp949_조사결과_및_결정.md`
  5번 항목 참고.

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

## .bat 파일 작성 규칙 (중요 — 2026-07-13 CP949 버그 조사 이후)
- **모든 .bat 파일 본문은 ASCII만 사용한다. 한글 주석/echo 텍스트를 파일에 직접
  넣지 않는다.** cmd.exe는 배치 파일을 읽을 때 실행 시점 콘솔의 코드페이지(한국어
  Windows 기본값 CP949)로 바이트를 해석하는데, UTF-8로 저장된 한글 텍스트가 이와
  어긋나면 `if (...)`/`for (...)` 같은 괄호 블록 안에서 한 줄이 엉뚱하게 여러
  명령으로 쪼개져 실행되는 버그가 있다. `chcp 65001`을 넣어도 괄호 블록은 파서가
  통째로 룩어헤드로 읽기 때문에 근본적으로 막지 못한다. 원인과 재현 테스트 전체
  기록은 `docs/legacy/cp949_조사결과.md` 참고.
- **팀원에게 보여줄 한글 메시지가 필요하면** `.bat` 안에 한글을 직접 쓰지 말고,
  PowerShell `-EncodedCommand`(Base64/UTF-16LE로 인코딩한 명령)를 통해 출력한다.
  `.bat` 파일 자체는 순수 ASCII로 유지되고, 한글 디코딩은 PowerShell이 담당해서
  cmd.exe의 코드페이지 문제를 원천적으로 피한다. `_auto_git_pull.bat`이 이 패턴의
  참고 예시.
- **내부 전용 헬퍼(팀원이 직접 더블클릭하지 않는 파일)는 한글 파일명 대신 ASCII
  파일명을 쓴다.** `_최신코드자동확인.bat`을 `_auto_git_pull.bat`으로 개명한 것이
  예시. 다만 `실행.bat`/`DB업데이트만.bat`/`보고서만생성.bat`/
  `전년대비보고서만생성.bat`처럼 팀원에게 이미 익숙한 한글 파일명은 개명 대상이
  아니다 — 파일명 자체는 NTFS가 관리하는 유니코드 메타데이터라 이 버그와 무관하고,
  문제는 배치 "본문" 안의 non-ASCII 바이트다.
- **✅ 실기기 검증 완료 (2026-07-13).** EncodedCommand/`_pull_retry.ps1`의
  한글 알림 메시지("[알림] 최신 코드 확인 실패, ...")가 실제 한국어 Windows
  콘솔 화면에서 깨지지 않고 정상 출력되는 것을 실기기(NAS `실행.bat` 더블클릭,
  마침 dubious ownership 실패 상황)에서 육안으로 확인했다. 아래 "dubious
  ownership" 건도 실제 팀원 PC에서 재현되어 원인 확인과 수정, NAS 재검증까지
  완료됨 — 자세한 내용은 `docs/legacy/cp949_조사결과_및_결정.md` 5번·6번
  항목 참고.
- **여러 줄짜리 PowerShell 로직은 `-EncodedCommand`(Base64)로 욱여넣지 말고
  별도 `.ps1` 파일로 분리한다.** `_pull_retry.ps1`이 예시 — 조건문 여러 개가
  섞인 로직을 Base64 한 줄에 넣으면 나중에 사람이 감사/디버깅할 때 다시
  디코딩해야만 내용을 볼 수 있어 실용성이 떨어진다. `.ps1` 파일은 UTF-8(BOM)로
  저장하면 cmd.exe 배치 파서를 거치지 않고 PowerShell이 직접 읽으므로 한글을
  써도 코드페이지 문제가 없다. `.bat`은 `powershell ... -File "%~dp0_xxx.ps1"`
  한 줄로 호출만 한다(본문은 여전히 ASCII 유지).
- **PowerShell에서 git config 값이 이미 등록됐는지 확인할 땐 PowerShell 문자열
  비교(`-contains`/`-notcontains`)로 직접 비교하지 말고 git 자신에게 물어본다**
  (`git config --fixed-value --get-all <key> <value>`의 종료 코드로 판단).
  한글처럼 유니코드 문자열을 비교할 때 `-notcontains`가 육안상 동일한 문자열도
  다르다고 판단해 값을 중복 추가하는 버그가 실제로 발견됐다(정규화 형태 차이로
  추정, `docs/legacy/cp949_조사결과_및_결정.md` 6번 항목 참고).
- **NAS `git pull` 실패의 흔한 원인 — dubious ownership (CVE-2022-24765).**
  Git 2.35.2+는 저장소 소유자 SID와 로그인 사용자 SID가 다르면 명령을 거부하는데,
  NAS(SMB 공유)는 소유자 SID가 로컬 계정과 구조적으로 안 맞는 경우가 흔해서
  네트워크 경로 저장소에서 특히 잘 걸린다. 실제 팀원 PC에서 `실행.bat` 더블클릭 시
  "[알림] 최신 코드 확인 실패"로 재현됐던 원인이 바로 이것. `_pull_retry.ps1`이
  `git config --global --add safe.directory`로 **이 NAS 경로 하나만** (`*`
  전체 신뢰 아님) 등록한 뒤 pull을 한 번 더 재시도하는 자가 치유(self-heal)
  로직으로 이미 수정됨 — 실제 NAS UNC 경로에서 최초 등록/재실행 시 중복
  미등록까지 검증 완료.

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
- `amd_report.py`의 `load_delivery_fee()`: 2026-05 정산 검증 중 개별통지문의
  "직배비공제(V+)"가 전체 매장 0원으로 나오는 문제 발견. 원인은 코드 버그가 아니라
  입력파일 문제 — `input/2026-05/`에 처음 받은 로젠 고객직배 원본 파일이 "물품명",
  "제주운임/산간료" 컬럼이 아예 없는 축소 양식이었음. `load_delivery_fee()`는
  집배구분/운송장번호/물품옵션/물품명/신용/제주 6개 컬럼을 헤더에서 전부 찾아야
  하고 하나라도 없으면 함수 전체가 빈 딕셔너리를 반환하도록 설계되어 있어(직배비
  3규칙 커밋 c1bcc0a, 2026-07-07 — 신용+제주운임 합산, 요청반품건 제외, 로젠
  약칭→정식매장명 매핑), 매장 필터링 이전 단계에서 이미 전 매장이 조용히 0으로
  처리됨. 정식 양식 파일로 교체 후 2026-05 전체 파이프라인 재실행해 정상화 확인
  (예: 롯데동래점 22,110원, 롯데광주점 91,795원, 현대천호점 383,625원, 24개
  중간관리 매장 합계 3,834,875원 — 2026-07-15). 향후 이런 "전체 매장 공통 0"
  증상은 특정 매장 로직보다 입력파일 헤더 양식 변경을 먼저 의심할 것.
- **2026-07-15 정산 검증 작업 요약** (PR #47~#51, 전부 main 병합·NAS 배포 완료):
  - PR #47: `load_delivery_fee()` 요청반품 제외 조건을 "요청반품&입금완료"에서
    "요청반품"(정산완료 여부 무관)으로 변경.
  - PR #48/#49: 매출이 아예 없는 "고아 매장"(예: 롯데아울렛서울역점)이
    `STORE_CODE_MAP` 전체 순회 로직 때문에 POS환급 균등배분 등에 잘못
    포함되던 버그 수정 — amd_report.py/expense_report.py/sales_summary.py/
    run.py/sales_analysis.py 5곳에 "이번 달 매출 있는 매장만 리스트업" 게이트
    통일 적용.
  - PR #50: `output/경영분석보고서.xlsx` 등 3개 파일이 월별 폴더와 별개로
    최상위에도 "최신본 바로가기"로 계속 생성되던 것을 완전 제거 — 격리
    테스트/과거월 재실행 한 번에 NAS 운영 파일이 조용히 옛 데이터로
    덮어써지는 사고가 실제 발생해 재발 방지 차원에서 결정.
  - PR #51: `판매수수료집계.xlsx` 인쇄영역이 템플릿에 `B2:G34`로 고정돼
    있어 행사매장 구분 행(35행)이 있는 달에 마지막 줄이 잘리던 것을
    마지막 데이터 행 기준 동적 계산으로 수정.

## 매뉴얼 문서 보관 원칙
- 관리자용(Git/PR/브랜치 보호 등 개발 워크플로우) 매뉴얼과 팀원용 실행 매뉴얼
  (입력파일 준비, 배치파일 사용법 등)을 원래는 "관리자용은 GitHub만, 팀원용은
  NAS에도"로 구분해 관리하려던 의도였다.
- **2026-07-13부터 이 구분은 더 이상 물리적으로 지켜지지 않는다.** NAS 배포
  방식이 robocopy(선택 복사)에서 git pull(전체 동기화) 기반으로 바뀌면서,
  NAS(`salary_system_claude`)의 `docs/`는 이제 이 저장소 `docs/` 전체와
  (관리자용 문서 포함) 자동으로 동일하게 유지된다 — git pull은 파일을 골라서
  받지 않는다. 예전엔 robocopy 배포 스크립트가 애초에 `docs/`를 복사 대상에
  넣지 않아서(`*.py`/`*.bat`/`requirements.txt`/`templates/`만 복사) NAS의
  `docs/`가 수동 관리되던 별개의 폴더였고, 그래서 "Git/PR 워크플로우 매뉴얼
  2건은 GitHub에만 있고 NAS엔 없음"이 실제로 성립했다. 지금은 그 2건
  (`Claude_Code_GitHub_시작_매뉴얼.docx`, `main_브랜치_보호_PR_워크플로우_매뉴얼.docx`)도
  포함해 `docs/`의 모든 파일이 NAS에 그대로 존재한다(2026-07-13 실제 pull로
  확인). 팀원은 어차피 실행 매뉴얼만 찾아보면 되므로 실무에 지장은 없지만,
  "관리자 전용 문서는 NAS에 안 보인다"고 가정하고 민감한 내용을 `docs/`에
  넣지 않도록 주의할 것 — 이제는 전부 NAS에도 물리적으로 존재한다.
- NAS(`salary_system_claude`) `docs/` 폴더 목록 (2026-07-10 기준 스냅샷, 대부분
  V4~V8 개발 이력 중 운영기준으로 남겨둔 문서들이고 GitHub `docs/`에도 동일
  파일이 있음. 새 파일을 저장소 `docs/`에 추가하면 다음 NAS git pull 때 자동
  반영되므로 이 목록은 참고용 스냅샷일 뿐 유지보수 대상이 아니다):
  - `00_V4_점검결과.txt` — V4 감사(audit) 결과
  - `01_템플릿_사용현황.txt` — 템플릿 사용 현황
  - `02_입력파일_사용현황.txt` — 입력 파일 사용 현황
  - `03_출력파일_생성흐름.txt` — 출력 파일 생성 흐름
  - `04_모듈_의존성.txt` — 모듈 의존성 정리
  - `05_삭제가능_파일목록.txt` — 삭제 가능/보관용 파일 목록
  - `06_필수파일목록.txt` — 필수 파일 목록
  - `07_DB_기준수치_운영기준.txt` — 월별손익DB 기준 수치 운영기준
  - `08_V5_경영분석보고서_추가.txt` — V5 경영분석 보고서 추가본
  - `09_매장구분파일_운영기준.txt` — 매장구분 파일 운영 기준
  - `10_KPI_Vminus_계산기준.txt` — KPI 계산 기준 변경 안내(V6)
  - `10_전년대비보고서_운영기준.txt` — 전년대비 보고서 운영 기준
  - `11_V7_최종운영안.txt` — V7 최종 운영안
  - `2026-05_기존DB_수치_검증.txt` — 2026-05 기존 DB 수치 반영 검증
  - `CONFIG_POLICY.md` — AMD V8.0.1 Git 관리 정책
  - `V8_Final_Release_검증결과.txt` — V8 Final Release 검증 결과
  - `저장소_구조.md` — AMD 저장소 구조 설명
  - `Claude_Code_GitHub_시작_매뉴얼.docx`, `main_브랜치_보호_PR_워크플로우_매뉴얼.docx`
    — 원래 "GitHub 전용 관리자 매뉴얼"로 의도했던 2건. git pull 방식으로 바뀐 뒤로는
    다른 파일들과 마찬가지로 NAS에도 존재한다(2026-07-13 확인, 위 안내 참고).
  - `legacy/` 하위 폴더도 git pull로 함께 동기화된다(`cp949_조사결과_및_결정.md`,
    `1_배포_NAS로_동기화.bat.old` 등 — 목록에 개별 나열하지 않음, 저장소 `docs/legacy/`
    참고).

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
