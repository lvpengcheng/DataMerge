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
start "FastAPI Server" cmd /k "uvicorn backend.app.main:app --reload --host 0.0.0.0 --port 8000"

echo.
echo 服务已启动，请访问: http://localhost:8000/docs
echo 按任意键退出...
pause >nul