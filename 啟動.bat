@echo off
cd /d "%~dp0"

python -c "import flask" 2>nul
if errorlevel 1 (
    echo [ERROR] Flask not found. Please run:
    echo    pip install -r requirements.txt
    pause
    exit /b 1
)

echo Starting server...
start /b python app.py
timeout /t 2 /nobreak > nul
start http://127.0.0.1:5001
echo Server is running. Close this window to stop.
pause
