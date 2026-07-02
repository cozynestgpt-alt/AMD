@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ===============================================
echo Sales commission system V8 Final Release
echo ===============================================
echo.

if exist requirements.txt (
    echo Installing Python dependencies...
    python -m pip install -r requirements.txt --quiet
    if errorlevel 1 goto ERROR
    echo.
)

python run.py
if errorlevel 1 goto ERROR

echo.
echo ===============================================
echo All tasks have been completed successfully.
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
