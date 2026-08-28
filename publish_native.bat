@echo off
chcp 65001 >nul 2>nul
setlocal
cd /d "%~dp0"

set "PY_CMD="
where py >nul 2>nul && set "PY_CMD=py -3"
if not defined PY_CMD where python >nul 2>nul && set "PY_CMD=python"
if not defined PY_CMD (
    echo [错误] 未找到 Python 3，请先安装 Python 3.11+
    pause
    exit /b 1
)

if not exist "deploy.native.config.json" (
    copy /Y "deploy.native.config.example.json" "deploy.native.config.json" >nul
    echo [首次使用] 已生成 deploy.native.config.json
    echo 请填写 Ubuntu 服务器地址、SSH用户、远程目录后，再次运行本脚本。
    start "" notepad "deploy.native.config.json"
    pause
    exit /b 2
)

%PY_CMD% -c "import paramiko" >nul 2>nul
if errorlevel 1 (
    echo [初始化] 安装本机发布依赖 paramiko...
    %PY_CMD% -m pip install paramiko
    if errorlevel 1 (
        echo [错误] paramiko 安装失败
        pause
        exit /b 1
    )
)

%PY_CMD% publish_native.py %*
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
    echo.
    echo [失败] 发布退出码 %RC%
    pause
)
exit /b %RC%
