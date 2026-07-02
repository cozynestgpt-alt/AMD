@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ===============================================
echo DB update only V8 Final Release
echo ===============================================
echo.

if exist requirements.txt (
    echo Installing Python dependencies...
    python -m pip install -r requirements.txt --quiet
    if errorlevel 1 goto ERROR
    echo.
)

python history_update.py
if errorlevel 1 goto ERROR

echo.
echo DB update has been completed successfully.
echo.
pause
exit /b 0

:ERROR
echo.
echo ERROR occurred. Please check the message above.
echo.
pause
exit /b 1
