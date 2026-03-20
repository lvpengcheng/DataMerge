@echo off
REM DataMerge IIS部署包构建脚本
REM 用途：创建可在IIS上直接部署的完整包

echo ========================================
echo DataMerge IIS部署包构建工具
echo ========================================
echo.

REM 设置变量
set VERSION=1.0.0
set BUILD_DIR=build_iis
set RELEASE_DIR=releases
set PACKAGE_NAME=DataMerge-IIS-v%VERSION%

REM 清理旧的构建目录
if exist %BUILD_DIR% (
    echo 清理旧的构建目录...
    rmdir /s /q %BUILD_DIR%
)

REM 创建构建目录
echo 创建构建目录...
mkdir %BUILD_DIR%
mkdir %BUILD_DIR%\logs
mkdir %BUILD_DIR%\tenants
mkdir %BUILD_DIR%\temp

REM 复制核心文件
echo 复制核心文件...
xcopy /E /I /Y backend %BUILD_DIR%\backend
xcopy /E /I /Y frontend %BUILD_DIR%\frontend
copy /Y excel_parser.py %BUILD_DIR%\

REM 复制配置文件
echo 复制配置文件...
copy /Y web.config %BUILD_DIR%\
copy /Y requirements.txt %BUILD_DIR%\
copy /Y .env.example %BUILD_DIR%\.env
copy /Y README.md %BUILD_DIR%\

REM 创建IIS部署说明
echo 创建部署说明...
(
echo # DataMerge IIS部署指南
echo.
echo ## 系统要求
echo - Windows Server 2016 或更高版本
echo - IIS 10.0 或更高版本
echo - Python 3.11 或更高版本
echo - HttpPlatformHandler 模块
echo.
echo ## 部署步骤
echo.
echo ### 1. 安装前置组件
echo.
echo 1.1 安装Python 3.11
echo - 下载: https://www.python.org/downloads/
echo - 安装时勾选 "Add Python to PATH"
echo.
echo 1.2 安装IIS HttpPlatformHandler
echo - 下载: https://www.iis.net/downloads/microsoft/httpplatformhandler
echo - 或使用Web Platform Installer安装
echo.
echo ### 2. 部署应用
echo.
echo 2.1 解压部署包到IIS目录
echo ```
echo C:\inetpub\wwwroot\DataMerge\
echo ```
echo.
echo 2.2 创建Python虚拟环境
echo ```
echo cd C:\inetpub\wwwroot\DataMerge
echo python -m venv venv
echo venv\Scripts\activate
echo pip install -r requirements.txt
echo ```
echo.
echo 2.3 配置环境变量
echo 编辑 .env 文件，设置AI API密钥：
echo ```
echo OPENAI_API_KEY=your_openai_key
echo ANTHROPIC_API_KEY=your_claude_key
echo DEEPSEEK_API_KEY=your_deepseek_key
echo ```
echo.
echo 2.4 设置文件夹权限
echo - 右键点击 DataMerge 文件夹 ^> 属性 ^> 安全
echo - 添加 IIS_IUSRS 用户组
echo - 授予"修改"权限（用于写入日志和临时文件）
echo.
echo ### 3. 配置IIS
echo.
echo 3.1 创建应用程序池
echo - 打开IIS管理器
echo - 右键"应用程序池" ^> "添加应用程序池"
echo - 名称: DataMerge
echo - .NET CLR版本: 无托管代码
echo - 托管管道模式: 集成
echo - 点击"确定"
echo.
echo 3.2 配置应用程序池高级设置
echo - 右键 DataMerge 应用程序池 ^> 高级设置
echo - 常规 ^> 启用32位应用程序: False
echo - 进程模型 ^> 标识: ApplicationPoolIdentity
echo - 进程模型 ^> 空闲超时: 0 ^(禁用超时^)
echo - 回收 ^> 固定时间间隔: 0 ^(禁用定期回收^)
echo.
echo 3.3 创建网站
echo - 右键"网站" ^> "添加网站"
echo - 网站名称: DataMerge
echo - 应用程序池: DataMerge
echo - 物理路径: C:\inetpub\wwwroot\DataMerge
echo - 绑定类型: http
echo - 端口: 8000 ^(或其他可用端口^)
echo - 点击"确定"
echo.
echo ### 4. 验证部署
echo.
echo 4.1 启动网站
echo - 在IIS管理器中右键网站 ^> 管理网站 ^> 启动
echo.
echo 4.2 访问应用
echo - 打开浏览器访问: http://localhost:8000
echo - 应该看到DataMerge训练页面
echo.
echo 4.3 查看日志
echo - 应用日志: C:\inetpub\wwwroot\DataMerge\logs\stdout.log
echo - IIS日志: C:\inetpub\logs\LogFiles\
echo.
echo ### 5. 故障排查
echo.
echo 5.1 503 Service Unavailable
echo - 检查Python是否正确安装
echo - 检查虚拟环境是否创建成功
echo - 检查web.config中的路径是否正确
echo - 查看stdout.log中的错误信息
echo.
echo 5.2 500 Internal Server Error
echo - 检查.env文件是否配置正确
echo - 检查文件夹权限
echo - 查看应用日志
echo.
echo 5.3 应用启动慢
echo - 首次启动需要加载依赖，可能需要1-2分钟
echo - 增加web.config中的startupTimeLimit值
echo.
echo ### 6. 生产环境优化
echo.
echo 6.1 启用HTTPS
echo - 在IIS中配置SSL证书
echo - 修改绑定为https，端口443
echo.
echo 6.2 配置应用程序池回收
echo - 设置固定时间回收^(如每天凌晨3点^)
echo - 避免在业务高峰期回收
echo.
echo 6.3 监控和日志
echo - 定期检查logs目录大小
echo - 配置日志轮转
echo - 监控应用程序池状态
echo.
echo 6.4 备份策略
echo - 定期备份tenants目录^(租户数据^)
echo - 备份.env配置文件
echo.
echo ## 技术支持
echo.
echo 如有问题，请查看：
echo - 应用日志: logs\stdout.log
echo - IIS日志: C:\inetpub\logs\LogFiles\
echo - Python错误: 在命令行手动运行应用查看详细错误
echo.
echo 手动测试命令：
echo ```
echo cd C:\inetpub\wwwroot\DataMerge
echo venv\Scripts\activate
echo python -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
echo ```
) > %BUILD_DIR%\IIS_DEPLOYMENT.md

REM 创建安装脚本
echo 创建安装脚本...
(
echo @echo off
echo echo ========================================
echo echo DataMerge IIS 自动安装脚本
echo echo ========================================
echo echo.
echo.
echo REM 检查Python
echo python --version ^>nul 2^>^&1
echo if errorlevel 1 ^(
echo     echo [错误] 未检测到Python，请先安装Python 3.11或更高版本
echo     echo 下载地址: https://www.python.org/downloads/
echo     pause
echo     exit /b 1
echo ^)
echo.
echo echo [1/4] 创建Python虚拟环境...
echo python -m venv venv
echo if errorlevel 1 ^(
echo     echo [错误] 虚拟环境创建失败
echo     pause
echo     exit /b 1
echo ^)
echo.
echo echo [2/4] 激活虚拟环境...
echo call venv\Scripts\activate.bat
echo.
echo echo [3/4] 安装依赖包...
echo pip install --upgrade pip
echo pip install -r requirements.txt
echo if errorlevel 1 ^(
echo     echo [错误] 依赖安装失败
echo     pause
echo     exit /b 1
echo ^)
echo.
echo echo [4/4] 配置完成
echo echo.
echo echo ========================================
echo echo 安装完成！
echo echo ========================================
echo echo.
echo echo 下一步：
echo echo 1. 编辑 .env 文件，配置AI API密钥
echo echo 2. 在IIS中创建网站，指向此目录
echo echo 3. 参考 IIS_DEPLOYMENT.md 完成IIS配置
echo echo.
echo pause
) > %BUILD_DIR%\install.bat

REM 创建测试脚本
echo 创建测试脚本...
(
echo @echo off
echo echo 启动DataMerge测试服务器...
echo call venv\Scripts\activate.bat
echo python -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000 --reload
) > %BUILD_DIR%\test_server.bat

REM 创建停止脚本
echo 创建停止脚本...
(
echo @echo off
echo echo 停止DataMerge服务...
echo for /f "tokens=5" %%%%a in ^('netstat -ano ^| findstr :8000'^) do ^(
echo     taskkill /F /PID %%%%a
echo ^)
echo echo 服务已停止
echo pause
) > %BUILD_DIR%\stop_server.bat

REM 创建.env示例文件
echo 创建环境变量示例...
(
echo # AI Provider Configuration
echo AI_PROVIDER=openai
echo.
echo # OpenAI Configuration
echo OPENAI_API_KEY=your_openai_api_key_here
echo OPENAI_BASE_URL=https://api.openai.com/v1
echo OPENAI_MODEL=gpt-4
echo.
echo # Claude Configuration
echo ANTHROPIC_API_KEY=your_anthropic_api_key_here
echo ANTHROPIC_MODEL=claude-3-5-sonnet-20241022
echo.
echo # DeepSeek Configuration
echo DEEPSEEK_API_KEY=your_deepseek_api_key_here
echo DEEPSEEK_BASE_URL=https://api.deepseek.com/v1
echo DEEPSEEK_MODEL=deepseek-chat
echo.
echo # Training Configuration
echo MAX_TRAINING_ITERATIONS=10
echo BEST_CODE_THRESHOLD=0.95
echo.
echo # Server Configuration
echo HOST=0.0.0.0
echo PORT=8000
) > %BUILD_DIR%\.env.example

REM 打包
echo 打包发布文件...
if not exist %RELEASE_DIR% mkdir %RELEASE_DIR%

REM 使用PowerShell压缩
powershell -command "Compress-Archive -Path '%BUILD_DIR%\*' -DestinationPath '%RELEASE_DIR%\%PACKAGE_NAME%.zip' -Force"

echo.
echo ========================================
echo 构建完成！
echo ========================================
echo.
echo 发布包位置: %RELEASE_DIR%\%PACKAGE_NAME%.zip
echo 构建目录: %BUILD_DIR%\
echo.
echo 部署说明：
echo 1. 解压 %PACKAGE_NAME%.zip 到 C:\inetpub\wwwroot\DataMerge
echo 2. 运行 install.bat 安装依赖
echo 3. 参考 IIS_DEPLOYMENT.md 配置IIS
echo.
pause
