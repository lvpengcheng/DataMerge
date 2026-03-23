@echo off
chcp 65001 >nul 2>nul
setlocal enabledelayedexpansion

REM ==========================================
REM  DataMerge 统一打包脚本
REM  支持: Windows IIS / Docker (Ubuntu)
REM ==========================================

set VERSION=1.0.0
set RELEASES_DIR=releases

echo.
echo ==========================================
echo   DataMerge 打包工具 v%VERSION%
echo ==========================================
echo.
echo   [1] Docker 部署包 (Ubuntu/Linux)
echo   [2] Windows IIS 部署包
echo   [3] 全部打包
echo   [0] 退出
echo.
set /p CHOICE=请选择打包目标:

if "%CHOICE%"=="1" goto :BUILD_DOCKER
if "%CHOICE%"=="2" goto :BUILD_IIS
if "%CHOICE%"=="3" goto :BUILD_ALL
if "%CHOICE%"=="0" exit /b 0
echo [错误] 无效选择
goto :EOF

:BUILD_ALL
call :BUILD_DOCKER
call :BUILD_IIS
goto :DONE

REM ==========================================
REM  Docker 打包
REM ==========================================
:BUILD_DOCKER
echo.
echo ------------------------------------------
echo   打包 Docker 部署包...
echo ------------------------------------------

set DOCKER_PKG=datamerge-%VERSION%
set DOCKER_DIR=%RELEASES_DIR%\%DOCKER_PKG%

if exist "%DOCKER_DIR%" rd /s /q "%DOCKER_DIR%"
mkdir "%DOCKER_DIR%"

echo [1/6] 复制后端代码...
xcopy /E /I /Q /Y backend "%DOCKER_DIR%\backend" >nul
for /d /r "%DOCKER_DIR%\backend" %%d in (__pycache__) do @if exist "%%d" rd /s /q "%%d"
del /q "%DOCKER_DIR%\backend\test_*.py" 2>nul

echo [2/6] 复制前端和静态资源...
xcopy /E /I /Q /Y frontend "%DOCKER_DIR%\frontend" >nul
if exist global_assets xcopy /E /I /Q /Y global_assets "%DOCKER_DIR%\global_assets" >nul

echo [3/6] 复制 libs (Aspose.Cells .NET DLL)...
xcopy /E /I /Q /Y libs "%DOCKER_DIR%\libs" >nul

echo [4/6] 复制中文字体...
if exist fonts (
    xcopy /E /I /Q /Y fonts "%DOCKER_DIR%\fonts" >nul
) else (
    echo [警告] fonts 目录不存在，Docker 镜像可能缺少中文字体
)

echo [5/6] 复制配置文件...
for %%f in (excel_parser.py aspose_init.py run.py requirements.txt) do (
    copy /Y "%%f" "%DOCKER_DIR%\" >nul
)
copy /Y Dockerfile "%DOCKER_DIR%\" >nul
copy /Y docker-compose.yml "%DOCKER_DIR%\" >nul
copy /Y .dockerignore "%DOCKER_DIR%\" >nul
copy /Y build_docker.sh "%DOCKER_DIR%\" >nul
if exist .env.example copy /Y .env.example "%DOCKER_DIR%\" >nul

echo [6/6] 打包 tar.gz...
if not exist "%RELEASES_DIR%" mkdir "%RELEASES_DIR%"
cd "%RELEASES_DIR%"
tar -czf "%DOCKER_PKG%.tar.gz" "%DOCKER_PKG%"
cd ..
rd /s /q "%DOCKER_DIR%"

for %%A in (%RELEASES_DIR%\%DOCKER_PKG%.tar.gz) do (
    echo.
    echo   [OK] Docker 包: %%~fA
    echo   [OK] 大小: %%~zA bytes
)
echo.
echo   部署步骤:
echo     1. scp %RELEASES_DIR%\%DOCKER_PKG%.tar.gz user@server:~/
echo     2. tar -xzf %DOCKER_PKG%.tar.gz ^&^& cd %DOCKER_PKG%
echo     3. bash build_docker.sh
echo.
goto :EOF

REM ==========================================
REM  Windows IIS 打包
REM ==========================================
:BUILD_IIS
echo.
echo ------------------------------------------
echo   打包 Windows IIS 部署包...
echo ------------------------------------------

set IIS_PKG=DataMerge-IIS-v%VERSION%
set IIS_DIR=%RELEASES_DIR%\%IIS_PKG%

if exist "%IIS_DIR%" rd /s /q "%IIS_DIR%"
mkdir "%IIS_DIR%"
mkdir "%IIS_DIR%\logs"
mkdir "%IIS_DIR%\tenants"
mkdir "%IIS_DIR%\data"
mkdir "%IIS_DIR%\output"
mkdir "%IIS_DIR%\temp"

echo [1/6] 复制后端代码...
xcopy /E /I /Q /Y backend "%IIS_DIR%\backend" >nul
for /d /r "%IIS_DIR%\backend" %%d in (__pycache__) do @if exist "%%d" rd /s /q "%%d"
del /q "%IIS_DIR%\backend\test_*.py" 2>nul

echo [2/6] 复制前端和静态资源...
xcopy /E /I /Q /Y frontend "%IIS_DIR%\frontend" >nul
if exist global_assets xcopy /E /I /Q /Y global_assets "%IIS_DIR%\global_assets" >nul

echo [3/6] 复制 libs (Aspose.Cells .NET DLL)...
xcopy /E /I /Q /Y libs "%IIS_DIR%\libs" >nul

echo [4/6] 复制核心文件...
for %%f in (excel_parser.py aspose_init.py run.py requirements.txt) do (
    copy /Y "%%f" "%IIS_DIR%\" >nul
)
if exist web.config copy /Y web.config "%IIS_DIR%\" >nul
if exist .env.example copy /Y .env.example "%IIS_DIR%\.env" >nul

echo [5/6] 生成安装脚本...
(
echo @echo off
echo chcp 65001 ^>nul 2^>nul
echo echo ========================================
echo echo   DataMerge 安装脚本
echo echo ========================================
echo echo.
echo.
echo python --version ^>nul 2^>^&1
echo if errorlevel 1 ^(
echo     echo [错误] 未检测到 Python，请先安装 Python 3.11+
echo     pause
echo     exit /b 1
echo ^)
echo.
echo echo [1/2] 安装依赖...
echo pip install --upgrade pip
echo pip install -r requirements.txt
echo.
echo echo [2/2] 初始化数据库...
echo python -m backend.database.init_db
echo.
echo echo ========================================
echo echo   安装完成!
echo echo ========================================
echo echo.
echo echo 下一步:
echo echo   1. 编辑 .env 文件，配置 AI API Key
echo echo   2. 运行 start.bat 启动服务
echo echo   3. 访问 http://localhost:8000
echo echo.
echo pause
) > "%IIS_DIR%\install.bat"

(
echo @echo off
echo chcp 65001 ^>nul 2^>nul
echo echo 启动 DataMerge 服务...
echo python -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
) > "%IIS_DIR%\start.bat"

echo [6/6] 打包 zip...
if not exist "%RELEASES_DIR%" mkdir "%RELEASES_DIR%"
powershell -command "Compress-Archive -Path '%IIS_DIR%\*' -DestinationPath '%RELEASES_DIR%\%IIS_PKG%.zip' -Force"
rd /s /q "%IIS_DIR%"

for %%A in (%RELEASES_DIR%\%IIS_PKG%.zip) do (
    echo.
    echo   [OK] IIS 包: %%~fA
    echo   [OK] 大小: %%~zA bytes
)
echo.
echo   部署步骤:
echo     1. 解压到 C:\inetpub\wwwroot\DataMerge
echo     2. 运行 install.bat
echo     3. 编辑 .env 配置 API Key
echo     4. 在 IIS 创建网站指向该目录
echo.
goto :EOF

:DONE
echo.
echo ==========================================
echo   全部打包完成!
echo ==========================================
echo.
dir /B "%RELEASES_DIR%\datamerge-*.tar.gz" 2>nul
dir /B "%RELEASES_DIR%\DataMerge-IIS-*.zip" 2>nul
echo.
pause
