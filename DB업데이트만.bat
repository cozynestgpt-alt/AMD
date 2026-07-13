@echo off
cd /d "%~dp0"

call "%~dp0_auto_git_pull.bat"

echo ===============================================
echo DB update only V8 Final Release
echo ===============================================
echo.

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
