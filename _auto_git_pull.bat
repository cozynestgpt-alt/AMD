@echo off
rem Internal helper called by the run bat files before invoking python.
rem Team members do not need to double-click this file directly.
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
        powershell -NoProfile -EncodedCommand VwByAGkAdABlAC0ASABvAHMAdAAgACcAWwBMxby5XQAgAFzN4MIgAFTP3LQgAFXWeMcgAOTCKNMsACAAMK50yCAAVM/ctFy4IADEyYnVadXIsuSyLgAnAA==
    )
    del /f /q "%LOCKFILE%" >nul 2>&1
)

exit /b 0
