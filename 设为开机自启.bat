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

set "APPDIR=%~dp0app"
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "VBS=%STARTUP%\TOCService.vbs"
> "%VBS%" echo Set sh = CreateObject("WScript.Shell")
>> "%VBS%" echo sh.Run """%PYW%"" ""%APPDIR%\toc_service.py"" --port 8765", 0, False

wscript "%VBS%"

echo.
echo [OK] Auto-start installed. The TOC service now runs HIDDEN in the
echo      background, and starts automatically every time you log in.
echo      Port: 8765
echo.
echo      To stop / remove it later, double-click the file named:
echo      qu-xiao-kai-ji-zi-qi  ("取消开机自启.bat")
echo.
pause
