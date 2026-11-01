@echo off
rem 手動啟動儀表板伺服器（背景執行，不會留下視窗），然後打開儀表板
cd /d "%~dp0"
start "" "%LOCALAPPDATA%\Programs\Python\Python313\pythonw.exe" "%~dp0server.py"
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8765/"
