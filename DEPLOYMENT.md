# DataMerge 部署指南

## 快速开始

### 方式1: 使用发布包（推荐）

1. 下载发布包
```bash
# 解压发布包
unzip DataMerge-v1.0.0.zip
cd DataMerge-v1.0.0

# Windows
install.bat

# Linux/Mac
chmod +x install.sh
./install.sh
```

2. 配置环境变量
```bash
# 复制环境变量模板
cp .env.example .env

# 编辑.env文件，配置AI API密钥
# OPENAI_API_KEY=your_key_here
# 或 ANTHROPIC_API_KEY=your_key_here
# 或 DEEPSEEK_API_KEY=your_key_here
```

3. 启动服务
```bash
python run.py --start
```

访问: http://localhost:8000

---

### 方式2: 从源码安装

1. 克隆代码
```bash
git clone <repository_url>
cd DataMerge
```

2. 创建虚拟环境
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

3. 安装依赖
```bash
pip install -r requirements.txt
```

4. 配置环境变量（同上）

5. 启动服务
```bash
python run.py --start
```

---

## 生产环境部署

### 使用 Docker 部署

1. 构建镜像
```bash
docker build -t datamerge:1.0.0 .
```

2. 运行容器
```bash
docker run -d \
  --name datamerge \
  -p 8000:8000 \
  -v $(pwd)/tenants:/app/tenants \
  -e OPENAI_API_KEY=your_key \
  datamerge:1.0.0
```

### 使用 Nginx 反向代理

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;

        # SSE 支持
        proxy_buffering off;
        proxy_cache off;
        proxy_read_timeout 86400s;
    }
}
```

### 使用 Systemd 服务

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

[Install]
WantedBy=multi-user.target
```

启动服务:
```bash
sudo systemctl daemon-reload
sudo systemctl enable datamerge
sudo systemctl start datamerge
```

---

## 环境变量配置

| 变量名 | 说明 | 默认值 |
|--------|------|--------|
| `AI_PROVIDER` | AI提供商 (openai/claude/deepseek/local) | openai |
| `OPENAI_API_KEY` | OpenAI API密钥 | - |
| `ANTHROPIC_API_KEY` | Claude API密钥 | - |
| `DEEPSEEK_API_KEY` | DeepSeek API密钥 | - |
| `OPENAI_BASE_URL` | OpenAI API基础URL | https://api.openai.com/v1 |
| `MAX_TRAINING_ITERATIONS` | 最大训练迭代次数 | 10 |
| `BEST_CODE_THRESHOLD` | 最佳代码合格率阈值 | 0.95 |

---

## 打包发布

### 创建发布包

Windows:
```bash
build_release.bat
```

Linux/Mac:
```bash
chmod +x build_release.sh
./build_release.sh
```

生成的发布包位于 `releases/` 目录。

### 发布到 PyPI（可选）

```bash
# 测试环境
python -m twine upload --repository testpypi dist/*

# 正式环境
python -m twine upload dist/*
```

---

## 故障排查

### 端口被占用
```bash
# Windows
netstat -ano | findstr :8000
taskkill /PID <pid> /F

# Linux/Mac
lsof -i :8000
kill -9 <pid>
```

### 依赖安装失败
```bash
# 使用国内镜像
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 日志查看
```bash
# 查看应用日志
tail -f logs/datamerge.log

# 查看系统服务日志
sudo journalctl -u datamerge -f
```

---

## 性能优化

1. **使用 Gunicorn 多进程**
```bash
gunicorn backend.app.main:app \
  --workers 4 \
  --worker-class uvicorn.workers.UvicornWorker \
  --bind 0.0.0.0:8000
```

2. **启用 Redis 缓存**
```python
# 在 .env 中配置
REDIS_URL=redis://localhost:6379/0
```

3. **数据库优化**
- 定期清理过期的训练数据
- 使用索引加速查询

---

## 安全建议

1. 使用 HTTPS
2. 设置强密码策略
3. 限制 API 访问频率
4. 定期备份租户数据
5. 不要在代码中硬编码 API 密钥

---

## 技术支持

- 文档: https://docs.datamerge.com
- 问题反馈: https://github.com/yourusername/datamerge/issues
- 邮箱: support@datamerge.com
