@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo 啟動經費結算系統...

python -c "import flask" 2>nul
if errorlevel 1 (
    echo.
    echo 找不到已安裝 flask 的 Python，請先執行：
    echo    pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

echo 使用 Python 啟動中，請稍候...
start /b python app.py
timeout /t 2 /nobreak > nul
start http://127.0.0.1:5001
echo.
echo 程式執行中。關閉此視窗即可停止程式。
pause
