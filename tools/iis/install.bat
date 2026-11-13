@echo off
setlocal
cd /d "%~dp0"
set "LOG=%~dp0logs\install.log"
if not exist "logs\" mkdir "logs"
if not exist "logs\" (
    echo Cannot create "%~dp0logs". Run from a writable site directory.
    set "RESULT=1"
    goto :FAILED
)
echo Installing the DataMerge IIS environment. Details: "%LOG%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %* > "%LOG%" 2>&1
set "RESULT=%errorlevel%"
if exist "%LOG%" type "%LOG%"
if not "%RESULT%"=="0" goto :FAILED
exit /b 0

:FAILED
echo IIS installation failed. See "%LOG%".
if not defined DATAMERGE_NO_PAUSE pause
exit /b %RESULT%
