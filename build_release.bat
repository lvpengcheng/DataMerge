@echo off
REM DataMerge 打包发布脚本 (Windows)

echo ==========================================
echo DataMerge 打包发布脚本
echo ==========================================

REM 1. 清理旧的构建文件
echo.
echo [1/6] 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
for /d /r %%d in (__pycache__) do @if exist "%%d" rmdir /s /q "%%d"
del /s /q *.pyc 2>nul

REM 2. 检查依赖
echo.
echo [2/6] 检查打包工具...
python -m pip install --upgrade pip setuptools wheel build twine

REM 3. 运行测试（可选）
echo.
echo [3/6] 运行测试...
if exist tests (
    python -m pytest tests/ -v || echo 警告: 测试失败，但继续打包
) else (
    echo 跳过测试（未找到tests目录）
)

REM 4. 构建源码包和wheel包
echo.
echo [4/6] 构建发布包...
python -m build

REM 5. 检查包
echo.
echo [5/6] 检查包完整性...
python -m twine check dist/*

REM 6. 创建发布压缩包
echo.
echo [6/6] 创建完整发布包...

REM 获取版本号
for /f "tokens=*" %%i in ('python -c "import re; content=open('setup.py').read(); print(re.search(r'version=\"([^\"]+)\"', content).group(1))"') do set VERSION=%%i
set RELEASE_NAME=DataMerge-v%VERSION%
set RELEASE_DIR=releases\%RELEASE_NAME%

if not exist releases mkdir releases
if exist "%RELEASE_DIR%" rmdir /s /q "%RELEASE_DIR%"
mkdir "%RELEASE_DIR%"

REM 复制必要文件
xcopy /s /y dist\* "%RELEASE_DIR%\"
copy README.md "%RELEASE_DIR%\"
copy requirements.txt "%RELEASE_DIR%\"
xcopy /s /e /y frontend "%RELEASE_DIR%\frontend\"
copy run.py "%RELEASE_DIR%\"

REM 创建安装脚本
echo @echo off > "%RELEASE_DIR%\install.bat"
echo echo 安装 DataMerge... >> "%RELEASE_DIR%\install.bat"
echo pip install -r requirements.txt >> "%RELEASE_DIR%\install.bat"
echo for %%%%f in (*.whl) do pip install %%%%f >> "%RELEASE_DIR%\install.bat"
echo echo 安装完成！ >> "%RELEASE_DIR%\install.bat"
echo echo 运行: python run.py --start >> "%RELEASE_DIR%\install.bat"
echo pause >> "%RELEASE_DIR%\install.bat"

REM 打包为zip
cd releases
powershell Compress-Archive -Path "%RELEASE_NAME%" -DestinationPath "%RELEASE_NAME%.zip" -Force
cd ..

echo.
echo ==========================================
echo 打包完成！
echo ==========================================
echo 发布包位置:
echo   - releases\%RELEASE_NAME%.zip
echo.
echo 包含内容:
echo   - Python wheel包 (可pip安装)
echo   - 源码包
echo   - 前端文件
echo   - 安装脚本
echo.
echo 发布到PyPI (可选):
echo   python -m twine upload dist/*
echo ==========================================
pause
