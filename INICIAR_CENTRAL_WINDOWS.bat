@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" "app\central_mobyan.py"
if errorlevel 1 (
    echo.
    echo A Central terminou com erro.
    pause
)
