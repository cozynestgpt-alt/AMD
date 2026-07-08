@echo off
set SCRIPT_DIR=%~dp0
python "%SCRIPT_DIR%run.py"
if errorlevel 1 pause
exit /b %errorlevel%
