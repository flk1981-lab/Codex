@echo off
setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0"

echo Repository: %CD%
echo.

where git >nul 2>&1
if errorlevel 1 (
    echo Git was not found in PATH.
    pause
    exit /b 1
)

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
    echo This folder is not a Git repository.
    pause
    exit /b 1
)

set "COMMIT_MSG="
set /p COMMIT_MSG=Commit message [update]: 
call :trim_trailing_spaces COMMIT_MSG
set "COMMIT_MSG_CHECK=%COMMIT_MSG: =%"
if not defined COMMIT_MSG_CHECK set "COMMIT_MSG=update"

echo.
echo Running: git add -A
git add -A
if errorlevel 1 (
    echo git add failed.
    pause
    exit /b 1
)

set "HAS_CHANGES="
for /f %%A in ('git status --porcelain') do (
    set "HAS_CHANGES=1"
    goto :changes_found
)

if not defined HAS_CHANGES (
    echo.
    echo No changes to commit.
    git status --short --branch
    pause
    exit /b 0
)

:changes_found

echo.
echo Running: git commit -m "%COMMIT_MSG%"
git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo git commit failed.
    pause
    exit /b 1
)

echo.
echo Running: git push
git push
if errorlevel 1 (
    echo git push failed.
    pause
    exit /b 1
)

echo.
echo Sync complete.
git status --short --branch
pause
exit /b 0

:trim_trailing_spaces
set "VALUE=!%~1!"
:trim_loop
if defined VALUE if "!VALUE:~-1!"==" " (
    set "VALUE=!VALUE:~0,-1!"
    goto :trim_loop
)
set "%~1=%VALUE%"
exit /b 0
