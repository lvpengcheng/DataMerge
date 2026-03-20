#!/bin/bash
# DataMerge Docker镜像打包脚本

set -e

echo "=========================================="
echo "DataMerge Docker镜像打包脚本"
echo "=========================================="

# 设置版本号
VERSION="1.0.0"
IMAGE_NAME="datamerge"
IMAGE_TAG="${IMAGE_NAME}:${VERSION}"
IMAGE_LATEST="${IMAGE_NAME}:latest"

echo ""
echo "镜像名称: ${IMAGE_TAG}"
echo ""

# 1. 构建Docker镜像
echo "[1/4] 构建Docker镜像..."
docker build -t ${IMAGE_TAG} -t ${IMAGE_LATEST} .

if [ $? -ne 0 ]; then
    echo "错误: Docker镜像构建失败"
    exit 1
fi

echo "✓ 镜像构建成功"

# 2. 测试镜像
echo ""
echo "[2/4] 测试Docker镜像..."
docker run --rm ${IMAGE_TAG} python -c "from backend.app.main import app; print('✓ 应用导入成功')"

if [ $? -ne 0 ]; then
    echo "错误: 镜像测试失败"
    exit 1
fi

# 3. 保存镜像为tar文件
echo ""
echo "[3/4] 导出Docker镜像..."
mkdir -p releases
docker save ${IMAGE_TAG} -o releases/${IMAGE_NAME}-${VERSION}.tar

if [ $? -eq 0 ]; then
    echo "✓ 镜像已导出到: releases/${IMAGE_NAME}-${VERSION}.tar"

    # 压缩tar文件
    echo "  压缩镜像文件..."
    gzip -f releases/${IMAGE_NAME}-${VERSION}.tar
    echo "✓ 压缩完成: releases/${IMAGE_NAME}-${VERSION}.tar.gz"
else
    echo "错误: 镜像导出失败"
    exit 1
fi

# 4. 创建部署说明
echo ""
echo "[4/4] 创建部署说明..."
cat > releases/DOCKER_DEPLOY.md << EOF
# DataMerge Docker部署指南

## 镜像信息

- 镜像名称: ${IMAGE_TAG}
- 镜像文件: ${IMAGE_NAME}-${VERSION}.tar.gz
- 构建时间: $(date)

## 部署步骤

### 方式1: 使用导出的镜像文件

1. 解压并加载镜像
\`\`\`bash
gunzip ${IMAGE_NAME}-${VERSION}.tar.gz
docker load -i ${IMAGE_NAME}-${VERSION}.tar
\`\`\`

2. 运行容器
\`\`\`bash
docker run -d \\
  --name datamerge \\
  -p 8000:8000 \\
  -v \$(pwd)/tenants:/app/tenants \\
  -v \$(pwd)/logs:/app/logs \\
  -e AI_PROVIDER=openai \\
  -e OPENAI_API_KEY=your_api_key_here \\
  ${IMAGE_TAG}
\`\`\`

3. 访问应用
\`\`\`
http://localhost:8000
\`\`\`

### 方式2: 使用docker-compose

1. 创建 docker-compose.yml
\`\`\`yaml
version: '3.8'

services:
  datamerge:
    image: ${IMAGE_TAG}
    container_name: datamerge
    ports:
      - "8000:8000"
    volumes:
      - ./tenants:/app/tenants
      - ./logs:/app/logs
    environment:
      - AI_PROVIDER=openai
      - OPENAI_API_KEY=your_api_key_here
    restart: unless-stopped
\`\`\`

2. 启动服务
\`\`\`bash
docker-compose up -d
\`\`\`

## 环境变量

| 变量名 | 说明 | 默认值 |
|--------|------|--------|
| AI_PROVIDER | AI提供商 | openai |
| OPENAI_API_KEY | OpenAI API密钥 | - |
| ANTHROPIC_API_KEY | Claude API密钥 | - |
| DEEPSEEK_API_KEY | DeepSeek API密钥 | - |
| MAX_TRAINING_ITERATIONS | 最大训练次数 | 10 |

## 常用命令

\`\`\`bash
# 查看日志
docker logs -f datamerge

# 停止容器
docker stop datamerge

# 启动容器
docker start datamerge

# 重启容器
docker restart datamerge

# 进入容器
docker exec -it datamerge bash

# 删除容器
docker rm -f datamerge

# 查看镜像
docker images | grep datamerge
\`\`\`

## 数据持久化

重要目录需要挂载到宿主机:
- \`/app/tenants\`: 租户数据
- \`/app/logs\`: 日志文件

## 健康检查

容器内置健康检查，可通过以下命令查看:
\`\`\`bash
docker inspect --format='{{.State.Health.Status}}' datamerge
\`\`\`

## 故障排查

### 容器无法启动
\`\`\`bash
# 查看详细日志
docker logs datamerge

# 检查容器状态
docker ps -a | grep datamerge
\`\`\`

### 端口冲突
\`\`\`bash
# 修改映射端口
docker run -p 8080:8000 ...
\`\`\`

### 权限问题
\`\`\`bash
# 确保挂载目录有正确权限
chmod -R 755 tenants logs
\`\`\`

## 生产环境建议

1. 使用环境变量文件
\`\`\`bash
docker run --env-file .env ...
\`\`\`

2. 限制资源使用
\`\`\`bash
docker run --memory="2g" --cpus="2" ...
\`\`\`

3. 配置日志驱动
\`\`\`bash
docker run --log-driver json-file --log-opt max-size=10m ...
\`\`\`

4. 使用反向代理（Nginx）
\`\`\`nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://localhost:8000;
        proxy_set_header Host \$host;
        proxy_buffering off;
    }
}
\`\`\`
EOF

echo "✓ 部署说明已创建: releases/DOCKER_DEPLOY.md"

# 显示镜像信息
echo ""
echo "=========================================="
echo "Docker镜像打包完成！"
echo "=========================================="
echo ""
echo "镜像信息:"
docker images | grep ${IMAGE_NAME} | head -2
echo ""
echo "导出文件:"
ls -lh releases/${IMAGE_NAME}-${VERSION}.tar.gz
echo ""
echo "部署说明: releases/DOCKER_DEPLOY.md"
echo ""
echo "快速部署:"
echo "  docker load -i releases/${IMAGE_NAME}-${VERSION}.tar.gz"
echo "  docker run -d -p 8000:8000 ${IMAGE_TAG}"
echo ""
echo "=========================================="
