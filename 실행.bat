@echo off
cd /d "%~dp0"

echo =================================================
echo   AMD - Sales Commission System  V8.0.0
echo =================================================
echo.

python run.py
exit /b %errorlevel%
