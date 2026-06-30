@echo off
chcp 65001 >nul
cd /d "E:\88. Claude\02_reminder"
powershell -NoProfile -ExecutionPolicy Bypass -File "setup_schedules.ps1"
pause
