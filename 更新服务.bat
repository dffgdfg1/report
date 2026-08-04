@echo off
chcp 65001 >nul
title 目录更新服务（保持开着）
cd /d "%~dp0app"

rem ============ 目录更新服务（在装了 WPS/Word 的 Windows 机器上跑） ============
rem 用法：双击本文件即可。第一次会自动装好需要的组件。
rem 虚拟机生成报告后会把 docx 发到这里，用 WPS 刷新目录页码再传回。
rem 让这个窗口一直开着。关了它，虚拟机就只能退回“手动更新目录”。
rem ============================================================================
set TOC_SERVICE_PORT=8765

rem --- 找 Python ---
set PY=
where python >nul 2>nul && set PY=python
if "%PY%"=="" (where py >nul 2>nul && set PY=py)
if "%PY%"=="" if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set PY="%LocalAppData%\Programs\Python\Python311\python.exe"
if "%PY%"=="" (
  echo.
  echo [!] 没找到 Python，请先装 Python 3.11+
  echo     下载: https://www.python.org/downloads/
  echo     安装时务必勾选 "Add python.exe to PATH"
  echo.
  pause
  exit /b
)

rem --- 自动补齐依赖（装过就秒过） ---
echo 正在检查运行组件（首次稍等）...
%PY% -c "import flask" 2>nul || %PY% -m pip install flask -q
%PY% -c "import win32com.client" 2>nul || %PY% -m pip install pywin32 -q

echo.
%PY% toc_service.py --port %TOC_SERVICE_PORT%
echo.
echo (服务已停止 - 可以关闭此窗口)
pause
