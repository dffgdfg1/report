@echo off
chcp 65001 >nul
setlocal

rem find python / pythonw
set "PY="
for /f "delims=" %%i in ('where python 2^>nul') do if not defined PY set "PY=%%i"
if not defined PY if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set "PY=%LocalAppData%\Programs\Python\Python311\python.exe"
set "PYW="
for /f "delims=" %%i in ('where pythonw 2^>nul') do if not defined PYW set "PYW=%%i"
if not defined PYW if exist "%LocalAppData%\Programs\Python\Python311\pythonw.exe" set "PYW=%LocalAppData%\Programs\Python\Python311\pythonw.exe"
if not defined PYW (
  echo [X] pythonw.exe not found. Install Python 3.11+ first:
  echo     https://www.python.org/downloads/
  pause & exit /b
)

echo Checking components (first run may take a moment)...
"%PY%" -c "import flask" 2>nul || "%PY%" -m pip install flask -q
"%PY%" -c "import win32com.client" 2>nul || "%PY%" -m pip install pywin32 -q

set "SCRIPT=%~dp0app\toc_service.py"

rem register auto-start at logon (HKCU Run key: no admin needed, not blocked by AV)
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v TOCUpdateService /t REG_SZ /d "\"%PYW%\" \"%SCRIPT%\" --port 8765" /f >nul
if errorlevel 1 (
  echo [X] Failed to write auto-start entry.
  pause & exit /b
)

rem start it right now, hidden (pythonw = no console window)
start "" "%PYW%" "%SCRIPT%" --port 8765

echo.
echo [OK] Auto-start installed. The TOC service is now running HIDDEN in
echo      the background, and will start automatically after every logon.
echo      Port: 8765
echo.
echo      Check it works:  open  http://192.168.24.68:8765/health  (shows "ok")
echo      To remove it:    double-click  "取消开机自启.bat"
echo.
pause
