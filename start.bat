@echo off
rem Double-click this file to start Email Bot on Windows.
rem
rem The first time, it sets up everything Email Bot needs inside a ".venv" folder
rem next to this file. Nothing is installed anywhere else on the PC. Later starts
rem skip straight to the app. Sending needs classic Outlook (not the new Outlook).

setlocal
cd /d "%~dp0"
set "VENV_PY=.venv\Scripts\python.exe"

if not exist "%VENV_PY%" goto create
"%VENV_PY%" -c "import sys" >nul 2>&1
if not errorlevel 1 goto install
echo The previous setup no longer works (Python may have been updated). Setting up again...
rmdir /s /q .venv

:create
echo First-time setup: getting Email Bot ready. This takes a minute or two.
echo.
rem Prefer the "py" launcher from python.org; plain "python" may be the Microsoft Store
rem placeholder, which fails this check.
set "PY="
py -3 -c "import sys, venv, ensurepip; sys.exit(sys.version_info < (3, 9))" >nul 2>&1 && set "PY=py -3"
if not defined PY python -c "import sys, venv, ensurepip; sys.exit(sys.version_info < (3, 9))" >nul 2>&1 && set "PY=python"
if not defined PY goto no_python
%PY% -m venv .venv >nul 2>&1
if errorlevel 1 goto venv_failed

:install
rem Installs anything missing, one package at a time (see emailbot\setup_deps.py).
rem Exit code 1 means something required is missing; 2 means only extras are.
"%VENV_PY%" -m emailbot.setup_deps
if errorlevel 2 goto run
if errorlevel 1 goto finish

:run
"%VENV_PY%" app.py
goto finish

:no_python
echo Email Bot needs Python (version 3.9 or newer), and this PC doesn't have it yet.
echo.
echo Install it from https://www.python.org/downloads/windows/
echo During installation, tick "Add python.exe to PATH".
echo If it asks for an administrator password you don't have, ask whoever manages this PC.
echo Then double-click start.bat again.
goto finish

:venv_failed
if exist .venv rmdir /s /q .venv
echo Couldn't set up Email Bot's folder. Make sure the Email Bot folder isn't read-only
echo (for example, still inside a zip file: right-click it and choose "Extract All" first),
echo then try again.

:finish
echo.
pause
