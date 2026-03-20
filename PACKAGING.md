# DataMerge 打包发布指南

## 📦 打包方式

### 方式1: 快速打包（推荐）

**Windows:**
```bash
build_release.bat
```

**Linux/Mac:**
```bash
chmod +x build_release.sh
./build_release.sh
```

生成的发布包位于 `releases/` 目录，包含：
- Python wheel包（可pip安装）
- 源码包
- 前端文件
- 安装脚本

---

### 方式2: Docker镜像

```bash
# 构建镜像
docker build -t datamerge:1.0.0 .

# 保存镜像为tar文件
docker save datamerge:1.0.0 -o datamerge-1.0.0.tar

# 或使用docker-compose
docker-compose build
```

---

### 方式3: 手动打包

```bash
# 1. 安装打包工具
pip install build twine

# 2. 构建包
python -m build

# 3. 检查包
python -m twine check dist/*

# 4. 创建发布压缩包
zip -r DataMerge-v1.0.0.zip \
  dist/ \
  frontend/ \
  backend/ \
  run.py \
  requirements.txt \
  README.md \
  DEPLOYMENT.md
```

---

## 🚀 部署方式

### 1. 标准部署（推荐给客户）

客户收到发布包后：

**Windows:**
```bash
# 解压
unzip DataMerge-v1.0.0.zip
cd DataMerge-v1.0.0

# 安装
install.bat

# 配置环境变量（编辑.env文件）
# 启动
python run.py --start
```

**Linux/Mac:**
```bash
# 解压
tar -xzf DataMerge-v1.0.0.tar.gz
cd DataMerge-v1.0.0

# 安装
chmod +x install.sh
./install.sh

# 配置环境变量（编辑.env文件）
# 启动
python run.py --start
```

---

### 2. Docker部署

```bash
# 使用docker-compose（最简单）
docker-compose up -d

# 或手动运行
docker run -d \
  --name datamerge \
  -p 8000:8000 \
  -v $(pwd)/tenants:/app/tenants \
  -e OPENAI_API_KEY=your_key \
  datamerge:1.0.0
```

---

### 3. 生产环境部署

#### 使用Gunicorn（多进程）

```bash
pip install gunicorn

gunicorn backend.app.main:app \
  --workers 4 \
  --worker-class uvicorn.workers.UvicornWorker \
  --bind 0.0.0.0:8000 \
  --access-logfile logs/access.log \
  --error-logfile logs/error.log
```

#### 使用Systemd服务

创建 `/etc/systemd/system/datamerge.service`:

```ini
[Unit]
Description=DataMerge Service
After=network.target

[Service]
Type=simple
User=www-data
WorkingDirectory=/opt/datamerge
Environment="PATH=/opt/datamerge/venv/bin"
ExecStart=/opt/datamerge/venv/bin/python run.py --start
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

启动服务:
```bash
sudo systemctl daemon-reload
sudo systemctl enable datamerge
sudo systemctl start datamerge
sudo systemctl status datamerge
```

#### 使用Nginx反向代理

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;

        # SSE支持（重要！）
        proxy_buffering off;
        proxy_cache off;
        proxy_read_timeout 86400s;
    }

    location /static {
        alias /opt/datamerge/frontend/static;
    }
}
```

---

## 📋 发布清单

打包前检查：

- [ ] 更新版本号（setup.py）
- [ ] 更新 README.md
- [ ] 更新 CHANGELOG.md
- [ ] 运行测试：`pytest tests/`
- [ ] 检查依赖：`pip list --outdated`
- [ ] 清理临时文件：`__pycache__`, `*.pyc`, `*.log`
- [ ] 检查 .gitignore 和 MANIFEST.in
- [ ] 测试打包：`python -m build`
- [ ] 测试安装：`pip install dist/*.whl`

---

## 🔧 环境变量配置

创建 `.env` 文件：

```bash
# AI提供商配置
AI_PROVIDER=openai  # openai/claude/deepseek/local

# API密钥（根据AI_PROVIDER选择配置）
OPENAI_API_KEY=sk-xxx
# ANTHROPIC_API_KEY=sk-ant-xxx
# DEEPSEEK_API_KEY=sk-xxx

# API基础URL（可选）
OPENAI_BASE_URL=https://api.openai.com/v1

# 训练参数
MAX_TRAINING_ITERATIONS=10
BEST_CODE_THRESHOLD=0.95

# 服务配置
HOST=0.0.0.0
PORT=8000
```

---

## 📊 版本管理

### 版本号规则

遵循语义化版本：`主版本.次版本.修订号`

- **主版本**：不兼容的API修改
- **次版本**：向下兼容的功能新增
- **修订号**：向下兼容的问题修正

### 发布流程

1. 开发分支完成功能
2. 合并到 main 分支
3. 更新版本号和文档
4. 打包测试
5. 创建 Git tag：`git tag v1.0.0`
6. 推送 tag：`git push origin v1.0.0`
7. 创建 GitHub Release
8. 上传发布包

---

## 🐛 常见问题

### Q: 打包后运行报错找不到模块？
A: 检查 MANIFEST.in 是否包含所有必要文件，确保 `include_package_data=True`

### Q: 前端静态文件没有打包进去？
A: 在 setup.py 中添加 `package_data` 或使用 MANIFEST.in

### Q: Docker镜像太大？
A: 使用 `.dockerignore` 排除不必要的文件（venv, tests, .git等）

### Q: 如何更新已部署的版本？
A:
```bash
# 停止服务
sudo systemctl stop datamerge

# 备份数据
cp -r tenants tenants.backup

# 安装新版本
pip install --upgrade datamerge-1.0.1.whl

# 重启服务
sudo systemctl start datamerge
```

---

## 📞 技术支持

- 详细部署文档：查看 `DEPLOYMENT.md`
- 问题反馈：GitHub Issues
- 邮箱：support@datamerge.com

---

## ✅ 快速命令参考

```bash
# 打包
./build_release.sh  # 或 build_release.bat

# 本地测试
python run.py --start

# Docker部署
docker-compose up -d

# 查看日志
docker-compose logs -f

# 停止服务
docker-compose down

# 生产部署
sudo systemctl start datamerge
sudo systemctl status datamerge
```
