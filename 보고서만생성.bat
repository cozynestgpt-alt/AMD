@echo off
chcp 65001 > nul
python history_report.py
python management_report.py
pause
