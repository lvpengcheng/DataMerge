@echo off
REM DataMerge 生产环境打包脚本 (Windows)
REM 生成可直接部署的完整包

echo ==========================================
echo DataMerge 生产环境打包脚本
echo ==========================================

REM 设置版本号
set VERSION=1.0.0
set RELEASE_NAME=DataMerge-v%VERSION%-production

echo.
echo 版本: %VERSION%
echo 发布名称: %RELEASE_NAME%
echo.

REM 1. 清理旧文件
echo [1/7] 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist releases\%RELEASE_NAME% rmdir /s /q releases\%RELEASE_NAME%
for /d /r %%d in (__pycache__) do @if exist "%%d" rmdir /s /q "%%d"
del /s /q *.pyc 2>nul

REM 2. 创建发布目录
echo.
echo [2/7] 创建发布目录...
if not exist releases mkdir releases
mkdir releases\%RELEASE_NAME%

REM 3. 复制应用文件
echo.
echo [3/7] 复制应用文件...
xcopy /s /e /y backend releases\%RELEASE_NAME%\backend\
xcopy /s /e /y frontend releases\%RELEASE_NAME%\frontend\
copy excel_parser.py releases\%RELEASE_NAME%\
copy run.py releases\%RELEASE_NAME%\
copy requirements.txt releases\%RELEASE_NAME%\
copy README.md releases\%RELEASE_NAME%\
if exist .env.example copy .env.example releases\%RELEASE_NAME%\
if exist DEPLOYMENT.md copy DEPLOYMENT.md releases\%RELEASE_NAME%\
if exist Dockerfile copy Dockerfile releases\%RELEASE_NAME%\
if exist docker-compose.yml copy docker-compose.yml releases\%RELEASE_NAME%\

REM 4. 清理不需要的文件
echo.
echo [4/7] 清理不需要的文件...
for /d /r releases\%RELEASE_NAME% %%d in (__pycache__) do @if exist "%%d" rmdir /s /q "%%d"
del /s /q releases\%RELEASE_NAME%\*.pyc 2>nul
del /s /q releases\%RELEASE_NAME%\*.log 2>nul

REM 5. 创建启动脚本
echo.
echo [5/7] 创建启动脚本...

REM Windows启动脚本
echo @echo off > releases\%RELEASE_NAME%\start.bat
echo echo ========================================== >> releases\%RELEASE_NAME%\start.bat
echo echo DataMerge 启动脚本 >> releases\%RELEASE_NAME%\start.bat
echo echo ========================================== >> releases\%RELEASE_NAME%\start.bat
echo. >> releases\%RELEASE_NAME%\start.bat
echo REM 检查Python >> releases\%RELEASE_NAME%\start.bat
echo python --version ^>nul 2^>^&1 >> releases\%RELEASE_NAME%\start.bat
echo if errorlevel 1 ( >> releases\%RELEASE_NAME%\start.bat
echo     echo 错误: 未找到Python，请先安装Python 3.8+ >> releases\%RELEASE_NAME%\start.bat
echo     pause >> releases\%RELEASE_NAME%\start.bat
echo     exit /b 1 >> releases\%RELEASE_NAME%\start.bat
echo ^) >> releases\%RELEASE_NAME%\start.bat
echo. >> releases\%RELEASE_NAME%\start.bat
echo REM 检查虚拟环境 >> releases\%RELEASE_NAME%\start.bat
echo if not exist venv ( >> releases\%RELEASE_NAME%\start.bat
echo     echo 首次运行，创建虚拟环境... >> releases\%RELEASE_NAME%\start.bat
echo     python -m venv venv >> releases\%RELEASE_NAME%\start.bat
echo     call venv\Scripts\activate >> releases\%RELEASE_NAME%\start.bat
echo     echo 安装依赖... >> releases\%RELEASE_NAME%\start.bat
echo     pip install -r requirements.txt >> releases\%RELEASE_NAME%\start.bat
echo ^) else ( >> releases\%RELEASE_NAME%\start.bat
echo     call venv\Scripts\activate >> releases\%RELEASE_NAME%\start.bat
echo ^) >> releases\%RELEASE_NAME%\start.bat
echo. >> releases\%RELEASE_NAME%\start.bat
echo echo 启动DataMerge服务... >> releases\%RELEASE_NAME%\start.bat
echo python run.py --start >> releases\%RELEASE_NAME%\start.bat

REM Linux启动脚本
echo #!/bin/bash > releases\%RELEASE_NAME%\start.sh
echo echo "==========================================" >> releases\%RELEASE_NAME%\start.sh
echo echo "DataMerge 启动脚本" >> releases\%RELEASE_NAME%\start.sh
echo echo "==========================================" >> releases\%RELEASE_NAME%\start.sh
echo. >> releases\%RELEASE_NAME%\start.sh
echo # 检查Python >> releases\%RELEASE_NAME%\start.sh
echo if ! command -v python3 ^&^> /dev/null; then >> releases\%RELEASE_NAME%\start.sh
echo     echo "错误: 未找到Python，请先安装Python 3.8+" >> releases\%RELEASE_NAME%\start.sh
echo     exit 1 >> releases\%RELEASE_NAME%\start.sh
echo fi >> releases\%RELEASE_NAME%\start.sh
echo. >> releases\%RELEASE_NAME%\start.sh
echo # 检查虚拟环境 >> releases\%RELEASE_NAME%\start.sh
echo if [ ! -d "venv" ]; then >> releases\%RELEASE_NAME%\start.sh
echo     echo "首次运行，创建虚拟环境..." >> releases\%RELEASE_NAME%\start.sh
echo     python3 -m venv venv >> releases\%RELEASE_NAME%\start.sh
echo     source venv/bin/activate >> releases\%RELEASE_NAME%\start.sh
echo     echo "安装依赖..." >> releases\%RELEASE_NAME%\start.sh
echo     pip install -r requirements.txt >> releases\%RELEASE_NAME%\start.sh
echo else >> releases\%RELEASE_NAME%\start.sh
echo     source venv/bin/activate >> releases\%RELEASE_NAME%\start.sh
echo fi >> releases\%RELEASE_NAME%\start.sh
echo. >> releases\%RELEASE_NAME%\start.sh
echo echo "启动DataMerge服务..." >> releases\%RELEASE_NAME%\start.sh
echo python run.py --start >> releases\%RELEASE_NAME%\start.sh

REM 6. 创建README
echo.
echo [6/7] 创建部署说明...
echo # DataMerge v%VERSION% 生产环境部署包 > releases\%RELEASE_NAME%\INSTALL.md
echo. >> releases\%RELEASE_NAME%\INSTALL.md
echo ## 快速开始 >> releases\%RELEASE_NAME%\INSTALL.md
echo. >> releases\%RELEASE_NAME%\INSTALL.md
echo ### Windows >> releases\%RELEASE_NAME%\INSTALL.md
echo 1. 解压此文件夹到目标位置 >> releases\%RELEASE_NAME%\INSTALL.md
echo 2. 双击运行 `start.bat` >> releases\%RELEASE_NAME%\INSTALL.md
echo 3. 访问 http://localhost:8000 >> releases\%RELEASE_NAME%\INSTALL.md
echo. >> releases\%RELEASE_NAME%\INSTALL.md
echo ### Linux/Mac >> releases\%RELEASE_NAME%\INSTALL.md
echo ```bash >> releases\%RELEASE_NAME%\INSTALL.md
echo chmod +x start.sh >> releases\%RELEASE_NAME%\INSTALL.md
echo ./start.sh >> releases\%RELEASE_NAME%\INSTALL.md
echo ``` >> releases\%RELEASE_NAME%\INSTALL.md
echo. >> releases\%RELEASE_NAME%\INSTALL.md
echo ## 配置 >> releases\%RELEASE_NAME%\INSTALL.md
echo. >> releases\%RELEASE_NAME%\INSTALL.md
echo 复制 `.env.example` 为 `.env` 并配置: >> releases\%RELEASE_NAME%\INSTALL.md
echo - AI_PROVIDER: AI提供商 (openai/claude/deepseek) >> releases\%RELEASE_NAME%\INSTALL.md
echo - OPENAI_API_KEY: OpenAI API密钥 >> releases\%RELEASE_NAME%\INSTALL.md
echo. >> releases\%RELEASE_NAME%\INSTALL.md
echo ## Docker部署 >> releases\%RELEASE_NAME%\INSTALL.md
echo. >> releases\%RELEASE_NAME%\INSTALL.md
echo ```bash >> releases\%RELEASE_NAME%\INSTALL.md
echo docker-compose up -d >> releases\%RELEASE_NAME%\INSTALL.md
echo ``` >> releases\%RELEASE_NAME%\INSTALL.md

REM 7. 打包
echo.
echo [7/7] 打包为ZIP...
cd releases
powershell Compress-Archive -Path "%RELEASE_NAME%" -DestinationPath "%RELEASE_NAME%.zip" -Force
cd ..

echo.
echo ==========================================
echo 打包完成！
echo ==========================================
echo.
echo 发布包位置: releases\%RELEASE_NAME%.zip
echo 大小:
dir releases\%RELEASE_NAME%.zip | findstr ".zip"
echo.
echo 部署方式:
echo 1. 解压ZIP文件
echo 2. 运行 start.bat (Windows) 或 start.sh (Linux)
echo 3. 访问 http://localhost:8000
echo.
echo ==========================================
pause
