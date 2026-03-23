#!/bin/bash
# ==========================================
#  DataMerge Docker 部署脚本（Ubuntu 服务器上运行）
# ==========================================
set -e

VERSION="1.0.0"
IMAGE_NAME="datamerge"
IMAGE_TAG="${IMAGE_NAME}:${VERSION}"

echo "=========================================="
echo "  DataMerge Docker 部署"
echo "=========================================="
echo ""

# 检查 Docker
if ! command -v docker &> /dev/null; then
    echo "[!] Docker 未安装，开始自动安装..."
    curl -fsSL https://get.docker.com | sh
    sudo usermod -aG docker $USER
    echo "[OK] Docker 已安装，请重新登录后再运行此脚本"
    exit 0
fi

# 检查 docker-compose
if ! command -v docker-compose &> /dev/null && ! docker compose version &> /dev/null; then
    echo "[!] docker-compose 未安装，安装中..."
    sudo apt-get update && sudo apt-get install -y docker-compose-plugin
fi

# 检查 .env
if [ ! -f .env ]; then
    if [ -f .env.example ]; then
        cp .env.example .env
        echo "[!] 已创建 .env 文件，请先编辑配置 API Key："
        echo "    vi .env"
        echo ""
        echo "  必填项："
        echo "    AI_PROVIDER=deepseek          # 或 openai / claude"
        echo "    DEEPSEEK_API_KEY=sk-xxx       # 对应的 API Key"
        echo ""
        read -p "是否现在编辑 .env? [Y/n] " answer
        if [ "$answer" != "n" ] && [ "$answer" != "N" ]; then
            ${EDITOR:-vi} .env
        fi
    else
        echo "[错误] 缺少 .env 和 .env.example 文件"
        exit 1
    fi
fi

echo ""
echo "[1/3] 构建 Docker 镜像（首次约 5-10 分钟）..."
docker build -t ${IMAGE_TAG} -t ${IMAGE_NAME}:latest .

echo ""
echo "[2/3] 创建数据目录..."
mkdir -p tenants data logs output

echo ""
echo "[3/3] 启动服务..."

# 优先使用 docker compose (v2)，回退到 docker-compose (v1)
if docker compose version &> /dev/null 2>&1; then
    docker compose up -d
else
    docker-compose up -d
fi

echo ""
echo "=========================================="
echo "  部署完成!"
echo "=========================================="
echo ""
echo "  访问地址: http://$(hostname -I | awk '{print $1}'):8000"
echo ""
echo "  常用命令:"
echo "    docker compose logs -f app     # 查看日志"
echo "    docker compose restart app     # 重启服务"
echo "    docker compose down            # 停止所有"
echo "    docker compose up -d --build   # 重新构建并启动"
echo ""
echo "=========================================="
