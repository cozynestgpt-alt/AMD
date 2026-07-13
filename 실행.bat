@echo off
set SCRIPT_DIR=%~dp0
call "%SCRIPT_DIR%_최신코드자동확인.bat"
python "%SCRIPT_DIR%run.py"
if errorlevel 1 pause
exit /b %errorlevel%
