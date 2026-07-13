@echo off
set SCRIPT_DIR=%~dp0
call "%SCRIPT_DIR%_auto_git_pull.bat"
python "%SCRIPT_DIR%run.py"
if errorlevel 1 pause
exit /b %errorlevel%
