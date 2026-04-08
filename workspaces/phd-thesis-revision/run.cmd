@echo off
setlocal

set "ROOT=%~dp0"
set "ACTION=%~1"
set "DOCX=%~2"
set "LOCAL_APPDATA=%LOCALAPPDATA%"

if "%ACTION%"=="" set "ACTION=check"
if "%DOCX%"=="" set "DOCX=%ROOT%input\working\thesis.docx"
if not defined LOCAL_APPDATA set "LOCAL_APPDATA=%USERPROFILE%\AppData\Local"

set "PYTHON_EXE="
if exist "%LOCAL_APPDATA%\Python\bin\python.exe" set "PYTHON_EXE=%LOCAL_APPDATA%\Python\bin\python.exe"
if not defined PYTHON_EXE if exist "%LOCAL_APPDATA%\Python\pythoncore-3.14-64\python.exe" set "PYTHON_EXE=%LOCAL_APPDATA%\Python\pythoncore-3.14-64\python.exe"
if not defined PYTHON_EXE if exist "%LOCAL_APPDATA%\Programs\Python\Python314\python.exe" set "PYTHON_EXE=%LOCAL_APPDATA%\Programs\Python\Python314\python.exe"
if not defined PYTHON_EXE if exist "%LOCAL_APPDATA%\Programs\Python\Python313\python.exe" set "PYTHON_EXE=%LOCAL_APPDATA%\Programs\Python\Python313\python.exe"
if not defined PYTHON_EXE set "PYTHON_EXE=python"

set "TOOL=%ROOT%tools\thesis_tools.py"
set "TERMS=%ROOT%config\terms.csv"
set "CHAPTERS=%ROOT%output\chapters"
set "OUTLINE=%ROOT%output\outline\outline.md"
set "STATS=%ROOT%output\reports\docx-stats.md"
set "TERM_REPORT=%ROOT%output\reports\term-check.md"
set "BACKUPS=%ROOT%output\backups"

if /I "%ACTION%"=="check" goto check
if /I "%ACTION%"=="stats" goto stats
if /I "%ACTION%"=="outline" goto outline
if /I "%ACTION%"=="split" goto split
if /I "%ACTION%"=="terms" goto terms
if /I "%ACTION%"=="backup" goto backup

echo Unsupported action: %ACTION%
exit /b 1

:check
"%PYTHON_EXE%" "%TOOL%" env-check --root "%ROOT%"
exit /b %ERRORLEVEL%

:stats
"%PYTHON_EXE%" "%TOOL%" docx-stats "%DOCX%" --output "%STATS%"
exit /b %ERRORLEVEL%

:outline
"%PYTHON_EXE%" "%TOOL%" outline "%DOCX%" --output "%OUTLINE%"
exit /b %ERRORLEVEL%

:split
"%PYTHON_EXE%" "%TOOL%" split-docx "%DOCX%" --outdir "%CHAPTERS%"
exit /b %ERRORLEVEL%

:terms
"%PYTHON_EXE%" "%TOOL%" term-check "%DOCX%" --terms "%TERMS%" --output "%TERM_REPORT%"
exit /b %ERRORLEVEL%

:backup
"%PYTHON_EXE%" "%TOOL%" backup "%DOCX%" --backup-dir "%BACKUPS%"
exit /b %ERRORLEVEL%
