#!/bin/bash
# DataMerge 打包发布脚本

set -e

echo "=========================================="
echo "DataMerge 打包发布脚本"
echo "=========================================="

# 1. 清理旧的构建文件
echo ""
echo "[1/6] 清理旧的构建文件..."
rm -rf build/ dist/ *.egg-info
find . -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
find . -type f -name "*.pyc" -delete 2>/dev/null || true

# 2. 检查依赖
echo ""
echo "[2/6] 检查打包工具..."
python -m pip install --upgrade pip setuptools wheel build twine

# 3. 运行测试（可选）
echo ""
echo "[3/6] 运行测试..."
if [ -d "tests" ]; then
    python -m pytest tests/ -v || echo "警告: 测试失败，但继续打包"
else
    echo "跳过测试（未找到tests目录）"
fi

# 4. 构建源码包和wheel包
echo ""
echo "[4/6] 构建发布包..."
python -m build

# 5. 检查包
echo ""
echo "[5/6] 检查包完整性..."
python -m twine check dist/*

# 6. 创建发布压缩包
echo ""
echo "[6/6] 创建完整发布包..."
VERSION=$(python -c "import re; content=open('setup.py').read(); print(re.search(r'version=\"([^\"]+)\"', content).group(1))")
RELEASE_NAME="DataMerge-v${VERSION}"
RELEASE_DIR="releases/${RELEASE_NAME}"

mkdir -p "${RELEASE_DIR}"

# 复制必要文件
cp -r dist/* "${RELEASE_DIR}/"
cp README.md "${RELEASE_DIR}/"
cp requirements.txt "${RELEASE_DIR}/"
cp -r frontend "${RELEASE_DIR}/"
cp run.py "${RELEASE_DIR}/"

# 创建安装脚本
cat > "${RELEASE_DIR}/install.sh" << 'EOF'
#!/bin/bash
echo "安装 DataMerge..."
pip install -r requirements.txt
pip install *.whl
echo "安装完成！"
echo "运行: python run.py --start"
EOF

cat > "${RELEASE_DIR}/install.bat" << 'EOF'
@echo off
echo 安装 DataMerge...
pip install -r requirements.txt
for %%f in (*.whl) do pip install %%f
echo 安装完成！
echo 运行: python run.py --start
pause
EOF

chmod +x "${RELEASE_DIR}/install.sh"

# 打包
cd releases
tar -czf "${RELEASE_NAME}.tar.gz" "${RELEASE_NAME}"
zip -r "${RELEASE_NAME}.zip" "${RELEASE_NAME}"
cd ..

echo ""
echo "=========================================="
echo "打包完成！"
echo "=========================================="
echo "发布包位置:"
echo "  - releases/${RELEASE_NAME}.tar.gz"
echo "  - releases/${RELEASE_NAME}.zip"
echo ""
echo "包含内容:"
echo "  - Python wheel包 (可pip安装)"
echo "  - 源码包"
echo "  - 前端文件"
echo "  - 安装脚本"
echo ""
echo "发布到PyPI (可选):"
echo "  python -m twine upload dist/*"
echo "=========================================="
