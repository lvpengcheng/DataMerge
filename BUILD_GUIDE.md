# DataMerge 完整打包部署指南

## 📦 三种打包方式

### 1. 生产环境打包（推荐给客户）

**特点**: 包含完整源码，自动安装依赖，一键启动

**Windows:**
```bash
build_production.bat
```

**Linux/Mac:**
```bash
chmod +x build_production.sh
./build_production.sh
```

**生成文件:**
- `releases/DataMerge-v1.0.0-production.zip`
- `releases/DataMerge-v1.0.0-production.tar.gz`

**客户部署:**
```bash
# 解压
unzip DataMerge-v1.0.0-production.zip
cd DataMerge-v1.0.0-production

# 启动（自动创建虚拟环境和安装依赖）
start.bat  # Windows
./start.sh # Linux/Mac
```

---

### 2. Docker镜像打包

**特点**: 容器化部署，环境隔离，易于管理

**Windows:**
```bash
build_docker.bat
```

**Linux/Mac:**
```bash
chmod +x build_docker.sh
./build_docker.sh
```

**生成文件:**
- `releases/datamerge-1.0.0.tar.gz` (Docker镜像)
- `releases/DOCKER_DEPLOY.md` (部署说明)

**客户部署:**
```bash
# 加载镜像
docker load -i datamerge-1.0.0.tar.gz

# 运行
docker run -d \
  --name datamerge \
  -p 8000:8000 \
  -v $(pwd)/tenants:/app/tenants \
  -e OPENAI_API_KEY=your_key \
  datamerge:1.0.0
```

---

### 3. Python包打包（开发者）

**特点**: 标准Python包，可发布到PyPI

**命令:**
```bash
# 清理旧文件
rm -rf build/ dist/

# 构建
python -m build

# 检查
python -m twine check dist/*
```

**生成文件:**
- `dist/datamerge-1.0.0-py3-none-any.whl`
- `dist/datamerge-1.0.0.tar.gz`

**安装:**
```bash
pip install datamerge-1.0.0-py3-none-any.whl
```

---

## 🎯 选择哪种打包方式？

| 场景 | 推荐方式 | 原因 |
|------|---------|------|
| 交付给客户 | 生产环境打包 | 简单易用，一键启动 |
| 服务器部署 | Docker镜像 | 环境隔离，易于管理 |
| 开发者分发 | Python包 | 标准化，可发布到PyPI |
| 内网部署 | 生产环境打包 | 无需Docker，兼容性好 |

---

## 📋 打包前检查清单

- [ ] 更新版本号（setup.py, build_production.bat/sh, build_docker.bat/sh）
- [ ] 测试所有功能正常
- [ ] 清理临时文件（`__pycache__`, `*.pyc`, `*.log`）
- [ ] 检查 MANIFEST.in 包含所有必要文件
- [ ] 更新 README.md 和文档
- [ ] 提交所有代码到Git

---

## 🚀 快速打包命令

### 打包所有格式（推荐）

**Windows:**
```bash
# 1. 生产环境包
build_production.bat

# 2. Docker镜像
build_docker.bat

# 3. Python包
python -m build
```

**Linux/Mac:**
```bash
# 1. 生产环境包
./build_production.sh

# 2. Docker镜像
./build_docker.sh

# 3. Python包
python -m build
```

---

## 📁 打包后的目录结构

```
releases/
├── DataMerge-v1.0.0-production.zip      # 生产环境包（Windows）
├── DataMerge-v1.0.0-production.tar.gz   # 生产环境包（Linux）
├── datamerge-1.0.0.tar.gz               # Docker镜像
└── DOCKER_DEPLOY.md                     # Docker部署说明

dist/
├── datamerge-1.0.0-py3-none-any.whl     # Python wheel包
└── datamerge-1.0.0.tar.gz               # Python源码包
```

---

## 🔧 各打包方式对比

| 特性 | 生产环境包 | Docker镜像 | Python包 |
|------|-----------|-----------|---------|
| 大小 | ~50MB | ~500MB | ~300KB |
| 安装难度 | ⭐ 简单 | ⭐⭐ 中等 | ⭐⭐⭐ 复杂 |
| 环境隔离 | ❌ 无 | ✅ 完全隔离 | ❌ 无 |
| 依赖管理 | 自动安装 | 内置 | 手动安装 |
| 适用场景 | 客户交付 | 服务器部署 | 开发分发 |
| 跨平台 | ✅ 是 | ✅ 是 | ✅ 是 |

---

## 📝 客户部署文档

### 生产环境包部署

1. **解压文件**
   ```bash
   unzip DataMerge-v1.0.0-production.zip
   cd DataMerge-v1.0.0-production
   ```

2. **配置环境变量**
   ```bash
   # 复制配置模板
   cp .env.example .env

   # 编辑.env文件，设置API密钥
   # AI_PROVIDER=openai
   # OPENAI_API_KEY=sk-xxx
   ```

3. **启动服务**
   ```bash
   # Windows
   start.bat

   # Linux/Mac
   chmod +x start.sh
   ./start.sh
   ```

4. **访问应用**
   ```
   http://localhost:8000
   ```

### Docker部署

1. **加载镜像**
   ```bash
   docker load -i datamerge-1.0.0.tar.gz
   ```

2. **创建配置文件**
   ```bash
   # 创建.env文件
   echo "OPENAI_API_KEY=your_key_here" > .env
   ```

3. **运行容器**
   ```bash
   docker run -d \
     --name datamerge \
     -p 8000:8000 \
     -v $(pwd)/tenants:/app/tenants \
     --env-file .env \
     datamerge:1.0.0
   ```

4. **查看日志**
   ```bash
   docker logs -f datamerge
   ```

---

## 🐛 常见问题

### Q1: 生产环境包启动失败？
**A:** 检查Python版本（需要3.8+）
```bash
python --version
```

### Q2: Docker镜像太大？
**A:** 使用 `.dockerignore` 排除不必要文件，已优化到最小

### Q3: 如何更新版本？
**A:** 修改以下文件中的版本号：
- `setup.py`
- `build_production.bat` / `build_production.sh`
- `build_docker.bat` / `build_docker.sh`

### Q4: 打包后缺少文件？
**A:** 检查 `MANIFEST.in` 是否包含该文件

### Q5: 如何验证打包是否成功？
**A:** 运行验证脚本
```bash
python verify_package.py
```

---

## 📞 技术支持

- 详细文档: `DEPLOYMENT.md`
- Docker部署: `releases/DOCKER_DEPLOY.md`
- 问题反馈: GitHub Issues

---

## ✅ 打包成功标志

打包成功后，你应该看到：

1. **生产环境包**
   - ✅ `releases/DataMerge-v1.0.0-production.zip` 存在
   - ✅ 包含 `start.bat` 和 `start.sh`
   - ✅ 包含 `INSTALL.md`

2. **Docker镜像**
   - ✅ `releases/datamerge-1.0.0.tar.gz` 存在
   - ✅ `docker images` 能看到镜像
   - ✅ 包含 `DOCKER_DEPLOY.md`

3. **Python包**
   - ✅ `dist/*.whl` 和 `dist/*.tar.gz` 存在
   - ✅ `twine check` 通过

---

## 🎉 快速开始

**最简单的打包方式（推荐给客户）:**

```bash
# Windows
build_production.bat

# Linux/Mac
./build_production.sh
```

生成的 ZIP 文件可以直接交付给客户，客户只需解压并运行 `start.bat` 即可！
