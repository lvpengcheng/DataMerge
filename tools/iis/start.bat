@echo off
setlocal
cd /d "%~dp0"
set "PYTHON=%~dp0venv-iis\Scripts\python.exe"
set "LOG=%~dp0logs\startup.log"
set "DOTNET_ROOT=%~dp0runtime\dotnet-9.0.21"
set "DOTNET_ROOT_X64=%DOTNET_ROOT%"
set "PATH=%DOTNET_ROOT%;%PATH%"
if not exist "logs\" mkdir "logs"
if not exist "logs\" (
    echo Cannot create "%~dp0logs". Run from a writable site directory.
    set "RESULT=1"
    goto :FAILED
)
if not exist "%PYTHON%" goto :NOT_INSTALLED
if not exist "%DOTNET_ROOT%\shared\Microsoft.NETCore.App\9.0.21\coreclr.dll" goto :NOT_INSTALLED
echo Starting DataMerge on port 8000. Details: "%LOG%"
"%PYTHON%" -u -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000 > "%LOG%" 2>&1
set "RESULT=%errorlevel%"
if not "%RESULT%"=="0" goto :FAILED
exit /b 0

:NOT_INSTALLED
> "%LOG%" echo The IIS runtime is incomplete. Run install.bat and check logs\install.log.
set "RESULT=1"

:FAILED
if exist "%LOG%" type "%LOG%"
echo DataMerge failed to start. See "%LOG%".
if not defined DATAMERGE_NO_PAUSE pause
exit /b %RESULT%
