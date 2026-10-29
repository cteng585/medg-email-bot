@echo off
rem Double-click this file (Windows) to start Email Bot.
cd /d "%~dp0"
where py >nul 2>nul && (set PY=py -3) || (set PY=python)
if not exist .venv\Scripts\python.exe (
  echo First-time setup: installing what Email Bot needs ^(this takes a minute^)...
  %PY% -m venv .venv || goto :nopython
  .venv\Scripts\python -m pip install -q --upgrade pip
  .venv\Scripts\python -m pip install -q -r requirements.txt || goto :failed
)
.venv\Scripts\python app.py
pause
exit /b
:nopython
echo Python 3 is needed. Download it from https://www.python.org/downloads/ (tick "Add Python to PATH") and try again.
pause
exit /b 1
:failed
echo Setup failed.
pause
exit /b 1
