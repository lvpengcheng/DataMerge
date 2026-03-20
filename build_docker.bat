@echo off
REM DataMerge Docker镜像打包脚本 (Windows)

echo ==========================================
echo DataMerge Docker镜像打包脚本
echo ==========================================

REM 设置版本号
set VERSION=1.0.0
set IMAGE_NAME=datamerge
set IMAGE_TAG=%IMAGE_NAME%:%VERSION%
set IMAGE_LATEST=%IMAGE_NAME%:latest

echo.
echo 镜像名称: %IMAGE_TAG%
echo.

REM 1. 构建Docker镜像
echo [1/4] 构建Docker镜像...
docker build -t %IMAGE_TAG% -t %IMAGE_LATEST% .

if errorlevel 1 (
    echo 错误: Docker镜像构建失败
    pause
    exit /b 1
)

echo 镜像构建成功
echo.

REM 2. 测试镜像
echo [2/4] 测试Docker镜像...
docker run --rm %IMAGE_TAG% python -c "from backend.app.main import app; print('应用导入成功')"

if errorlevel 1 (
    echo 错误: 镜像测试失败
    pause
    exit /b 1
)

echo.

REM 3. 保存镜像为tar文件
echo [3/4] 导出Docker镜像...
if not exist releases mkdir releases
docker save %IMAGE_TAG% -o releases\%IMAGE_NAME%-%VERSION%.tar

if errorlevel 1 (
    echo 错误: 镜像导出失败
    pause
    exit /b 1
)

echo 镜像已导出到: releases\%IMAGE_NAME%-%VERSION%.tar
echo.

REM 4. 创建部署说明
echo [4/4] 创建部署说明...
echo # DataMerge Docker部署指南 > releases\DOCKER_DEPLOY.md
echo. >> releases\DOCKER_DEPLOY.md
echo ## 镜像信息 >> releases\DOCKER_DEPLOY.md
echo. >> releases\DOCKER_DEPLOY.md
echo - 镜像名称: %IMAGE_TAG% >> releases\DOCKER_DEPLOY.md
echo - 镜像文件: %IMAGE_NAME%-%VERSION%.tar >> releases\DOCKER_DEPLOY.md
echo. >> releases\DOCKER_DEPLOY.md
echo ## 部署步骤 >> releases\DOCKER_DEPLOY.md
echo. >> releases\DOCKER_DEPLOY.md
echo ### 1. 加载镜像 >> releases\DOCKER_DEPLOY.md
echo ```bash >> releases\DOCKER_DEPLOY.md
echo docker load -i %IMAGE_NAME%-%VERSION%.tar >> releases\DOCKER_DEPLOY.md
echo ``` >> releases\DOCKER_DEPLOY.md
echo. >> releases\DOCKER_DEPLOY.md
echo ### 2. 运行容器 >> releases\DOCKER_DEPLOY.md
echo ```bash >> releases\DOCKER_DEPLOY.md
echo docker run -d ^\ >> releases\DOCKER_DEPLOY.md
echo   --name datamerge ^\ >> releases\DOCKER_DEPLOY.md
echo   -p 8000:8000 ^\ >> releases\DOCKER_DEPLOY.md
echo   -v ${PWD}/tenants:/app/tenants ^\ >> releases\DOCKER_DEPLOY.md
echo   -e OPENAI_API_KEY=your_key ^\ >> releases\DOCKER_DEPLOY.md
echo   %IMAGE_TAG% >> releases\DOCKER_DEPLOY.md
echo ``` >> releases\DOCKER_DEPLOY.md
echo. >> releases\DOCKER_DEPLOY.md
echo ### 3. 访问应用 >> releases\DOCKER_DEPLOY.md
echo http://localhost:8000 >> releases\DOCKER_DEPLOY.md
echo. >> releases\DOCKER_DEPLOY.md
echo ## 常用命令 >> releases\DOCKER_DEPLOY.md
echo. >> releases\DOCKER_DEPLOY.md
echo ```bash >> releases\DOCKER_DEPLOY.md
echo # 查看日志 >> releases\DOCKER_DEPLOY.md
echo docker logs -f datamerge >> releases\DOCKER_DEPLOY.md
echo. >> releases\DOCKER_DEPLOY.md
echo # 停止容器 >> releases\DOCKER_DEPLOY.md
echo docker stop datamerge >> releases\DOCKER_DEPLOY.md
echo. >> releases\DOCKER_DEPLOY.md
echo # 启动容器 >> releases\DOCKER_DEPLOY.md
echo docker start datamerge >> releases\DOCKER_DEPLOY.md
echo ``` >> releases\DOCKER_DEPLOY.md

echo 部署说明已创建: releases\DOCKER_DEPLOY.md
echo.

REM 显示镜像信息
echo ==========================================
echo Docker镜像打包完成！
echo ==========================================
echo.
echo 镜像信息:
docker images | findstr %IMAGE_NAME%
echo.
echo 导出文件:
dir releases\%IMAGE_NAME%-%VERSION%.tar | findstr ".tar"
echo.
echo 部署说明: releases\DOCKER_DEPLOY.md
echo.
echo 快速部署:
echo   docker load -i releases\%IMAGE_NAME%-%VERSION%.tar
echo   docker run -d -p 8000:8000 %IMAGE_TAG%
echo.
echo ==========================================
pause
