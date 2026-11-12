@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .env copy .env.example .env >nul
if not exist .venv (
  python -m venv .venv
  .venv\Scripts\python -m pip install -r requirements.txt
)
.venv\Scripts\python app.py
pause
