@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" ocr_classifier.py
echo.
pause
