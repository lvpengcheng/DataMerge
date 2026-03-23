@echo off
chcp 65001 >nul 2>nul
REM ==========================================
REM  DataMerge 打包脚本（用于 Ubuntu Docker 部署）
REM  在 Windows 上运行，生成 .tar.gz 部署包
REM ==========================================

echo ==========================================
echo   DataMerge 部署包打包
echo ==========================================

set VERSION=1.0.0
set PACKAGE_NAME=datamerge-%VERSION%
set RELEASES_DIR=releases
set STAGE_DIR=%RELEASES_DIR%\%PACKAGE_NAME%

echo.
echo 版本: %VERSION%
echo 输出: %RELEASES_DIR%\%PACKAGE_NAME%.tar.gz
echo.

REM 清理旧的暂存目录
if exist "%STAGE_DIR%" rd /s /q "%STAGE_DIR%"
mkdir "%STAGE_DIR%"

echo [1/5] 复制后端代码...
xcopy /E /I /Q /Y backend "%STAGE_DIR%\backend" >nul
REM 清理 __pycache__
for /d /r "%STAGE_DIR%\backend" %%d in (__pycache__) do @if exist "%%d" rd /s /q "%%d"
REM 清理 test 文件
del /q "%STAGE_DIR%\backend\test_*.py" 2>nul

echo [2/5] 复制前端和静态资源...
xcopy /E /I /Q /Y frontend "%STAGE_DIR%\frontend" >nul
if exist global_assets xcopy /E /I /Q /Y global_assets "%STAGE_DIR%\global_assets" >nul

echo [3/5] 复制 libs（Aspose.Cells .NET DLL）...
xcopy /E /I /Q /Y libs "%STAGE_DIR%\libs" >nul

echo [4/5] 复制配置文件...
copy /Y excel_parser.py "%STAGE_DIR%\" >nul
copy /Y aspose_init.py "%STAGE_DIR%\" >nul
copy /Y run.py "%STAGE_DIR%\" >nul
copy /Y requirements.txt "%STAGE_DIR%\" >nul
copy /Y Dockerfile "%STAGE_DIR%\" >nul
copy /Y docker-compose.yml "%STAGE_DIR%\" >nul
copy /Y .dockerignore "%STAGE_DIR%\" >nul
if exist .env.example copy /Y .env.example "%STAGE_DIR%\" >nul

echo [5/5] 打包 tar.gz...
cd "%RELEASES_DIR%"
tar -czf "%PACKAGE_NAME%.tar.gz" "%PACKAGE_NAME%"
cd ..

REM 清理暂存目录
rd /s /q "%STAGE_DIR%"

echo.
echo ==========================================
echo   打包完成!
echo ==========================================
echo.
for %%A in (%RELEASES_DIR%\%PACKAGE_NAME%.tar.gz) do echo 文件: %%~fA
for %%A in (%RELEASES_DIR%\%PACKAGE_NAME%.tar.gz) do echo 大小: %%~zA bytes
echo.
echo ---- Ubuntu 部署步骤 ----
echo.
echo   1. 上传到服务器:
echo      scp %RELEASES_DIR%\%PACKAGE_NAME%.tar.gz user@server:~/
echo.
echo   2. 在服务器上解压:
echo      tar -xzf %PACKAGE_NAME%.tar.gz
echo      cd %PACKAGE_NAME%
echo.
echo   3. 配置环境变量:
echo      cp .env.example .env
echo      vi .env   # 填入 AI_PROVIDER 和 API Key
echo.
echo   4. 构建并启动:
echo      docker-compose up -d --build
echo.
echo   5. 访问: http://服务器IP:8000
echo.
echo ==========================================
pause
