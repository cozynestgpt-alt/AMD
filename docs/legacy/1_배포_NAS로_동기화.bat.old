@echo off
setlocal

rem ============================================================
rem  코드를 NAS로 배포합니다 (git이 아니라 순수 파일 복사)
rem  - NAS(Y:\)는 git 저장소가 아닙니다. 팀원은 NAS에서 input 폴더에
rem    자료만 넣고 .bat 실행 결과만 받으며, 코드는 절대 건드리지 않습니다.
rem  - 코드의 진짜 버전은 이 PC + GitHub(main 브랜치)에만 있습니다.
rem  - input/output/DB/backup 은 이 스크립트가 절대 건드리지 않습니다.
rem ============================================================

set "SRC=%~dp0"
set "DST=\\nas\CN_AMD\판매수수료 관련자료\salary_system_claude"

echo ============================================================
echo   코드를 NAS로 배포합니다
echo   원본: %SRC%
echo   대상: %DST%
echo   (input/output/DB/backup 은 건드리지 않습니다)
echo ============================================================
echo.

if not exist "%DST%" (
    echo   [오류] NAS 경로에 접근할 수 없습니다: %DST%
    pause
    exit /b 1
)

echo   [1/2] Python 스크립트 / 배치 파일 / requirements.txt 복사 중...
robocopy "%SRC%." "%DST%" *.py *.bat requirements.txt /R:2 /W:2 /NJH
set RC1=%ERRORLEVEL%

echo.
echo   [2/2] templates 폴더 복사 중...
robocopy "%SRC%templates" "%DST%\templates" /E /R:2 /W:2 /NJH
set RC2=%ERRORLEVEL%

echo.
if %RC1% GEQ 8 (
    echo   [오류] 스크립트 파일 복사 중 문제가 발생했습니다 ^(코드 %RC1%^)
) else if %RC2% GEQ 8 (
    echo   [오류] templates 폴더 복사 중 문제가 발생했습니다 ^(코드 %RC2%^)
) else (
    echo   완료: NAS 배포가 정상적으로 끝났습니다.
)

echo.
pause
