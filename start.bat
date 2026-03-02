@echo off
chcp 65001 >nul

REM 激活虚拟环境
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
    echo [信息] 虚拟环境已激活
) else (
    echo [警告] 未找到虚拟环境，使用系统Python
)

echo.
echo ========================================
echo   AI驱动的Excel数据整合SaaS系统
echo ========================================
echo.

REM 启动选项
echo.
echo ========================================
echo           启动选项
echo ========================================
echo 1. 启动FastAPI服务 (默认)
echo 2. 运行测试
echo 3. 运行规则解析器演示
echo 4. 创建示例文件
echo 5. 退出
echo.

set /p CHOICE="请选择 (1-5，默认1): "
if "%CHOICE%"=="" set CHOICE=1

if "%CHOICE%"=="1" (
    echo.
    echo ========================================
    echo         启动FastAPI服务
    echo ========================================
    echo 服务地址: http://localhost:8000
    echo API文档: http://localhost:8000/docs
    echo 按 Ctrl+C 停止服务
    echo.

    REM 使用Python run.py启动，已配置排除tenants文件夹
    python run.py --start

) else if "%CHOICE%"=="2" (
    echo.
    echo ========================================
    echo           运行测试
    echo ========================================
    echo 运行测试套件...
    pytest tests\ -v

) else if "%CHOICE%"=="3" (
    echo.
    echo ========================================
    echo       运行规则解析器演示
    echo ========================================
    python examples\rule_parser_demo.py

) else if "%CHOICE%"=="4" (
    echo.
    echo ========================================
    echo         创建示例文件
    echo ========================================
    python -c "import sys; sys.path.insert(0, '.'); from run import create_example_files; create_example_files(); print('示例文件已创建到 examples\\ 目录')"

) else if "%CHOICE%"=="5" (
    echo 退出
) else (
    echo 无效选择
)

echo.
pause