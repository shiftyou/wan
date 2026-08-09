@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" update_xlsx.py
echo.
pause
