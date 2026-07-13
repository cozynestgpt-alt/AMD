@echo off
chcp 65001 > nul
rem 실행용 .bat들이 python을 호출하기 전에 공통으로 call 하는 내부 헬퍼입니다.
rem 팀원이 직접 더블클릭할 필요는 없습니다.
setlocal
set "LOCKFILE=%~dp0.pulling.lock"

if exist "%LOCKFILE%" (
    powershell -NoProfile -Command "if ((Get-Item '%LOCKFILE%').LastWriteTime -lt (Get-Date).AddMinutes(-10)) { exit 1 } else { exit 0 }" >nul 2>&1
    if errorlevel 1 del /f /q "%LOCKFILE%" >nul 2>&1
)

if not exist "%LOCKFILE%" (
    echo. > "%LOCKFILE%" 2>nul
    git -C "%~dp0." pull --ff-only >nul 2>&1
    if errorlevel 1 (
        echo [알림] 최신 코드 확인 실패, 기존 코드로 진행합니다.
    )
    del /f /q "%LOCKFILE%" >nul 2>&1
)

exit /b 0
