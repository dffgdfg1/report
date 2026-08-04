@echo off
chcp 65001 >nul
set "VBS=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\TOCService.vbs"
if exist "%VBS%" del "%VBS%"
wmic process where "name='pythonw.exe' and commandline like '%%toc_service%%'" delete >nul 2>nul
echo.
echo [OK] Auto-start removed, and the running service was stopped.
echo.
pause
