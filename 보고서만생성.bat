@echo off
set SCRIPT_DIR=%~dp0
call "%SCRIPT_DIR%_최신코드자동확인.bat"

python "%SCRIPT_DIR%history_report.py"
if errorlevel 1 goto ERROR

python "%SCRIPT_DIR%management_report.py"
if errorlevel 1 goto ERROR

python "%SCRIPT_DIR%year_compare_report.py"
if errorlevel 1 goto ERROR

pause
exit /b 0

:ERROR
pause
exit /b 1
