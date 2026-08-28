@echo off
chcp 65001 >nul
echo 重启FastAPI服务...
echo.

REM 查找并停止现有的uvicorn进程
echo 停止现有服务...
taskkill /F /IM uvicorn.exe 2>nul
taskkill /F /IM python.exe 2>nul
timeout /t 2 >nul

echo 启动新服务...
cd /d "%~dp0"
set "APP_PORT=18000"
if exist ".env" (
    for /f "usebackq tokens=1,* delims==" %%A in (`findstr /B /C:"APP_PORT=" ".env"`) do set "APP_PORT=%%B"
)
start "FastAPI Server" cmd /k "uvicorn backend.app.main:app --reload --host 0.0.0.0 --port %APP_PORT%"

echo.
echo 服务已启动，请访问: http://localhost:%APP_PORT%/docs
echo 按任意键退出...
pause >nul
