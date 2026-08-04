@echo off
chcp 65001 >nul
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v TOCUpdateService /f >nul 2>nul
wmic process where "name='pythonw.exe' and commandline like '%%toc_service%%'" delete >nul 2>nul
echo.
echo [OK] Auto-start removed, and the running service was stopped.
echo.
pause
