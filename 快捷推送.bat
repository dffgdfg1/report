@echo off
chcp 65001 >nul
setlocal EnableExtensions DisableDelayedExpansion

rem Always run from the directory containing this script.
cd /d "%~dp0"

echo ================================================
echo   Report Generator - Quick GitHub Push
echo ================================================
echo.

rem Verify that Git for Windows is installed.
where git >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Git was not found. Install Git for Windows first.
    goto :fail
)

rem Avoid running Git commands from the wrong directory.
git rev-parse --is-inside-work-tree >nul 2>nul
if errorlevel 1 (
    echo [ERROR] This directory is not a Git repository:
    echo %CD%
    goto :fail
)

rem Verify that the GitHub remote exists.
git remote get-url origin >nul 2>nul
if errorlevel 1 (
    echo [ERROR] The remote named origin does not exist.
    goto :fail
)

rem The verified .gitignore protects runtime data and the large template.
if not exist ".gitignore" (
    echo [ERROR] .gitignore is missing. Push stopped for safety.
    goto :fail
)

echo Current branch and changes:
echo ------------------------------------------------
git status --short --branch
echo ------------------------------------------------
echo.

rem Detect whether the working tree contains changes.
set "HAS_CHANGES="
for /f "delims=" %%F in ('git status --porcelain') do set "HAS_CHANGES=1"

if not defined HAS_CHANGES (
    echo No new file changes. Checking existing local commits.
    goto :sync_and_push
)

rem Show changes before staging so an accidental file can be noticed.
choice /C YN /N /M "Commit the changes shown above and push? [Y/N] "
if errorlevel 2 goto :cancelled

set "DEFAULT_MSG=Update %date% %time:~0,8%"
set "COMMIT_MSG="
set /p "COMMIT_MSG=Commit message, or press Enter to use the current time: "
if not defined COMMIT_MSG set "COMMIT_MSG=%DEFAULT_MSG%"

rem Stage all changes except paths protected by .gitignore.
git add -A -- .
if errorlevel 1 (
    echo [ERROR] git add failed.
    goto :fail
)

echo.
echo Files to commit:
echo ------------------------------------------------
git status --short
echo ------------------------------------------------

rem Skip the commit if staging produced no changes.
git diff --cached --quiet
if not errorlevel 1 goto :sync_and_push

rem Create a local commit.
git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo [ERROR] git commit failed.
    goto :fail
)

:sync_and_push
echo.
echo Pulling the latest GitHub commits...

rem Rebase before pushing to avoid an unnecessary merge commit.
git pull --rebase origin main
if errorlevel 1 (
    echo [ERROR] Pull failed or a conflict occurred. Nothing was pushed.
    echo Resolve the Git message above, then run this script again.
    goto :fail
)

echo.
echo Pushing to GitHub...
git push -u origin main
if errorlevel 1 (
    echo [ERROR] Push failed. Check the network and GitHub sign-in.
    goto :fail
)

echo.
echo [SUCCESS] The code is now on GitHub.
git status --short --branch
echo.
pause
exit /b 0

:cancelled
echo.
echo Cancelled. No files were staged, committed, or pushed.
echo.
pause
exit /b 0

:fail
echo.
echo The operation did not complete. Review the message above.
echo.
pause
exit /b 1
