@echo off
setlocal
cd /d "%~dp0"

echo Starting SharePoint Permission Manager...
python -m pip install -r config/requirements.txt
echo Starting Python Backend Server...
start "SharePoint Permission Manager" python app\main.py
echo Waiting for server to start...
timeout /t 3 /nobreak >nul
start http://127.0.0.1:5000