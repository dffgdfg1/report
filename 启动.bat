@echo off
cd /d "%~dp0app"
set PY=
where python >nul 2>nul && set PY=python
if "%PY%"=="" (where py >nul 2>nul && set PY=py)
if "%PY%"=="" if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set PY="%LocalAppData%\Programs\Python\Python311\python.exe"
if "%PY%"=="" (
  echo.
  echo [!] Python not found. Please install Python 3.11+ first.
  echo     Download: https://www.python.org/downloads/
  echo     During install, tick "Add python.exe to PATH".
  echo.
  pause
  exit /b
)
%PY% bootstrap.py
%PY% app.py
echo.
echo (server stopped - you can close this window)
pause
