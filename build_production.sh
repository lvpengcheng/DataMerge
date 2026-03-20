#!/bin/bash
# DataMerge 生产环境打包脚本 (Linux/Mac)
# 生成可直接部署的完整包

set -e

echo "=========================================="
echo "DataMerge 生产环境打包脚本"
echo "=========================================="

# 设置版本号
VERSION="1.0.0"
RELEASE_NAME="DataMerge-v${VERSION}-production"

echo ""
echo "版本: ${VERSION}"
echo "发布名称: ${RELEASE_NAME}"
echo ""

# 1. 清理旧文件
echo "[1/7] 清理旧的构建文件..."
rm -rf build/ dist/ releases/${RELEASE_NAME}
find . -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
find . -type f -name "*.pyc" -delete 2>/dev/null || true

# 2. 创建发布目录
echo ""
echo "[2/7] 创建发布目录..."
mkdir -p releases/${RELEASE_NAME}

# 3. 复制应用文件
echo ""
echo "[3/7] 复制应用文件..."
cp -r backend releases/${RELEASE_NAME}/
cp -r frontend releases/${RELEASE_NAME}/
cp excel_parser.py releases/${RELEASE_NAME}/
cp run.py releases/${RELEASE_NAME}/
cp requirements.txt releases/${RELEASE_NAME}/
cp README.md releases/${RELEASE_NAME}/
[ -f .env.example ] && cp .env.example releases/${RELEASE_NAME}/
[ -f DEPLOYMENT.md ] && cp DEPLOYMENT.md releases/${RELEASE_NAME}/
[ -f Dockerfile ] && cp Dockerfile releases/${RELEASE_NAME}/
[ -f docker-compose.yml ] && cp docker-compose.yml releases/${RELEASE_NAME}/

# 4. 清理不需要的文件
echo ""
echo "[4/7] 清理不需要的文件..."
find releases/${RELEASE_NAME} -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
find releases/${RELEASE_NAME} -type f -name "*.pyc" -delete 2>/dev/null || true
find releases/${RELEASE_NAME} -type f -name "*.log" -delete 2>/dev/null || true

# 5. 创建启动脚本
echo ""
echo "[5/7] 创建启动脚本..."

# Linux启动脚本
cat > releases/${RELEASE_NAME}/start.sh << 'EOF'
#!/bin/bash
echo "=========================================="
echo "DataMerge 启动脚本"
echo "=========================================="

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到Python，请先安装Python 3.8+"
    exit 1
fi

# 检查虚拟环境
if [ ! -d "venv" ]; then
    echo "首次运行，创建虚拟环境..."
    python3 -m venv venv
    source venv/bin/activate
    echo "安装依赖..."
    pip install -r requirements.txt
else
    source venv/bin/activate
fi

echo "启动DataMerge服务..."
python run.py --start
EOF

chmod +x releases/${RELEASE_NAME}/start.sh

# Windows启动脚本
cat > releases/${RELEASE_NAME}/start.bat << 'EOF'
@echo off
echo ==========================================
echo DataMerge 启动脚本
echo ==========================================

REM 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.8+
    pause
    exit /b 1
)

REM 检查虚拟环境
if not exist venv (
    echo 首次运行，创建虚拟环境...
    python -m venv venv
    call venv\Scripts\activate
    echo 安装依赖...
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate
)

echo 启动DataMerge服务...
python run.py --start
EOF

# 6. 创建README
echo ""
echo "[6/7] 创建部署说明..."
cat > releases/${RELEASE_NAME}/INSTALL.md << EOF
# DataMerge v${VERSION} 生产环境部署包

## 快速开始

### Windows
1. 解压此文件夹到目标位置
2. 双击运行 \`start.bat\`
3. 访问 http://localhost:8000

### Linux/Mac
\`\`\`bash
chmod +x start.sh
./start.sh
\`\`\`

## 配置

复制 \`.env.example\` 为 \`.env\` 并配置:
- AI_PROVIDER: AI提供商 (openai/claude/deepseek)
- OPENAI_API_KEY: OpenAI API密钥

## Docker部署

\`\`\`bash
docker-compose up -d
\`\`\`

## 系统要求

- Python 3.8+
- 2GB+ RAM
- 1GB+ 磁盘空间

## 端口

- 8000: Web服务端口

## 目录结构

- backend/: 后端代码
- frontend/: 前端文件
- tenants/: 租户数据（自动创建）
- logs/: 日志文件（自动创建）

## 故障排查

### 端口被占用
\`\`\`bash
# Windows
netstat -ano | findstr :8000

# Linux/Mac
lsof -i :8000
\`\`\`

### 依赖安装失败
\`\`\`bash
# 使用国内镜像
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
\`\`\`

## 技术支持

- 文档: README.md, DEPLOYMENT.md
- 问题反馈: GitHub Issues
EOF

# 7. 打包
echo ""
echo "[7/7] 打包为压缩文件..."
cd releases
tar -czf ${RELEASE_NAME}.tar.gz ${RELEASE_NAME}
zip -r ${RELEASE_NAME}.zip ${RELEASE_NAME} >/dev/null 2>&1 || echo "zip命令不可用，跳过zip打包"
cd ..

echo ""
echo "=========================================="
echo "打包完成！"
echo "=========================================="
echo ""
echo "发布包位置:"
ls -lh releases/${RELEASE_NAME}.tar.gz 2>/dev/null || true
ls -lh releases/${RELEASE_NAME}.zip 2>/dev/null || true
echo ""
echo "部署方式:"
echo "1. 解压文件"
echo "2. 运行 start.sh (Linux/Mac) 或 start.bat (Windows)"
echo "3. 访问 http://localhost:8000"
echo ""
echo "=========================================="
