@echo off
chcp 65001 >nul
cd /d "%~dp0"
title AI Ofis
where py >nul 2>nul && (set PY=py -3) || (set PY=python)
if not exist .venv (
  echo [AI Ofis] Birinchi ishga tushirish: muhit tayyorlanmoqda...
  %PY% -m venv .venv || (echo Python topilmadi. python.org dan Python 3.11+ o'rnating & pause & exit /b 1)
)
call .venv\Scripts\activate.bat
python -m pip install -q --upgrade pip >nul
python -m pip install -q -r requirements.txt || (echo Kutubxonalarni o'rnatib bo'lmadi & pause & exit /b 1)
if not exist .env (
  copy .env.example .env >nul
  echo [AI Ofis] .env fayli yaratildi. Unga ANTHROPIC_API_KEY yozing.
  notepad .env
)
echo [AI Ofis] Server ishga tushdi: http://localhost:8000  (to'xtatish: Ctrl+C)
start "" http://localhost:8000
python -m uvicorn server.app:app --host 127.0.0.1 --port 8000
pause
