@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ===============================================
echo Report generation V8 Final Release
echo ===============================================
echo.

python history_report.py
if errorlevel 1 goto ERROR

python management_report.py
if errorlevel 1 goto ERROR

python year_compare_report.py
if errorlevel 1 goto ERROR

echo.
echo ===============================================
echo All reports have been created successfully.
echo ===============================================
echo.
pause
exit /b 0

:ERROR
echo.
echo ===============================================
echo ERROR occurred. Please check the message above.
echo ===============================================
echo.
pause
exit /b 1
