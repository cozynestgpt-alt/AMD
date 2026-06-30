@echo off
chcp 65001 >nul
title 판매수수료 전년대비보고서 생성
cd /d "%~dp0"
echo ==========================================
echo   판매수수료 전년대비보고서 생성
echo ==========================================
echo.
python year_compare_report.py
pause
