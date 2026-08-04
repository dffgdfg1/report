@echo off
chcp 65001 >nul
title TOC Update Service - keep this window open
cd /d "%~dp0app"

set TOC_SERVICE_PORT=8765

set PY=
where python >nul 2>nul && set PY=python
if "%PY%"=="" (where py >nul 2>nul && set PY=py)
if "%PY%"=="" if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set PY="%LocalAppData%\Programs\Python\Python311\python.exe"
if "%PY%"=="" (
  echo.
  echo [X] Python not found. Install Python 3.11+ first:
  echo     https://www.python.org/downloads/
  echo     Check "Add python.exe to PATH" during install.
  echo.
  pause
  exit /b
)

echo Checking components (first run may take a moment)...
%PY% -c "import flask" 2>nul || %PY% -m pip install flask -q
%PY% -c "import win32com.client" 2>nul || %PY% -m pip install pywin32 -q

echo.
%PY% toc_service.py --port %TOC_SERVICE_PORT%
echo.
echo (Service stopped - you can close this window)
pause
