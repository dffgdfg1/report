@echo off
chcp 65001 >nul
cd /d "%~dp0app"

rem ============ 目录更新服务（在装了 WPS/Word 的 Windows 机器上跑） ============
rem 虚拟机生成报告后会把 docx 发到这里，用 WPS 刷新目录页码再传回。
rem 让这个窗口一直开着。关了它，虚拟机就只能退回“手动更新目录”。
rem
rem ★ 口令：把下面 TOKEN 改成你自己的（和虚拟机 toc_service.json 里的 token 一致）
set TOC_SERVICE_TOKEN=change-me-8f3a
set TOC_SERVICE_PORT=8765
rem ============================================================================

set PY=
where python >nul 2>nul && set PY=python
if "%PY%"=="" (where py >nul 2>nul && set PY=py)
if "%PY%"=="" if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set PY="%LocalAppData%\Programs\Python\Python311\python.exe"
if "%PY%"=="" (
  echo.
  echo [!] 没找到 Python，请先装 Python 3.11+
  echo     下载: https://www.python.org/downloads/
  echo     安装时勾选 "Add python.exe to PATH"
  echo.
  pause
  exit /b
)

%PY% toc_service.py --port %TOC_SERVICE_PORT% --token %TOC_SERVICE_TOKEN%
echo.
echo (服务已停止 - 可以关闭此窗口)
pause
