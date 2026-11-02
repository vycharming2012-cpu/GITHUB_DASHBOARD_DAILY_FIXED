@echo off
cd /d "%~dp0"
python build_daily.py
if errorlevel 1 pause & exit /b 1
echo Dashboard ready: docs\index.html
pause
