@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ===============================================
echo Year compare report only V8 Final Release
echo ===============================================
echo.

python year_compare_report.py
if errorlevel 1 goto ERROR

echo.
echo Year compare report has been created successfully.
echo.
pause
exit /b 0

:ERROR
echo.
echo ERROR occurred. Please check the message above.
echo.
pause
exit /b 1
