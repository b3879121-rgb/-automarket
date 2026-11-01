@echo off
rem 每日盤後自動執行：抓資料（資料不齊會重試到 18:30）→ 產生儀表板
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
"%PY%" fetch_after_close.py --wait
"%PY%" build_dashboard.py