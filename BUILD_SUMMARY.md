# DataMerge IIS 部署包构建完成

## ✅ 构建成功

**部署包位置**：`releases/DataMerge-IIS-v1.0.0.zip`
**包大小**：805 KB
**构建时间**：2026-03-18

## 📦 包内容清单

### 核心代码
- ✅ backend/ - 后端Python代码（AI引擎、API、沙箱等）
- ✅ frontend/ - 前端页面（训练页面、智算页面）
- ✅ excel_parser.py - Excel智能解析器

### 配置文件
- ✅ web.config - IIS配置文件（HttpPlatformHandler）
- ✅ requirements.txt - Python依赖列表
- ✅ .env - 环境变量配置模板
- ✅ .env.example - 环境变量示例

### 部署脚本
- ✅ install.bat - 自动安装Python依赖
- ✅ test_server.bat - 本地测试服务器
- ✅ stop_server.bat - 停止服务器

### 文档
- ✅ PACKAGE_README.md - 部署包说明（完整）
- ✅ IIS_QUICK_START.md - 快速部署指南（5分钟）
- ✅ IIS_DEPLOYMENT.md - 详细部署文档
- ✅ README.md - 项目说明

### 目录结构
- ✅ logs/ - 日志目录
- ✅ tenants/ - 租户数据目录
- ✅ temp/ - 临时文件目录

## 🚀 部署步骤（简要）

### 1. 解压部署包
```
解压到：C:\inetpub\wwwroot\DataMerge\
```

### 2. 安装依赖
```cmd
双击运行 install.bat
```

### 3. 配置API密钥
```
编辑 .env 文件，填入你的API密钥
```

### 4. 安装HttpPlatformHandler
```
下载：https://www.iis.net/downloads/microsoft/httpplatformhandler
```

### 5. 配置IIS
```
创建应用程序池 → 创建网站 → 设置权限 → 启动
```

详细步骤请参考包内的 `IIS_QUICK_START.md`

## 🔧 系统要求

| 组件 | 版本要求 |
|------|---------|
| Windows Server | 2016+ |
| IIS | 10.0+ |
| Python | 3.11+ |
| HttpPlatformHandler | 1.2+ |
| 内存 | 4GB+ |
| 磁盘 | 10GB+ |

## 📝 重要配置说明

### web.config 关键配置
```xml
<httpPlatform
  processPath="C:\inetpub\wwwroot\DataMerge\venv\Scripts\python.exe"
  arguments="-m uvicorn backend.app.main:app --host 0.0.0.0 --port %HTTP_PLATFORM_PORT%"
  startupTimeLimit="60"
  startupRetryCount="3">
```

**注意**：
- `processPath` 需要指向虚拟环境中的Python
- 首次启动建议将 `startupTimeLimit` 设为 120

### .env 必填项
```env
# 至少配置一个AI提供商的API密钥
OPENAI_API_KEY=sk-xxx        # OpenAI
ANTHROPIC_API_KEY=sk-ant-xxx # Claude
DEEPSEEK_API_KEY=sk-xxx      # DeepSeek
```

## ✨ 新功能特性

### 前端优化
- ✅ 浅蓝色主题设计
- ✅ 左侧导航菜单
- ✅ 移除租户ID显示
- ✅ 统一的视觉风格

### 流式日志增强
- ✅ 实时显示后端日志
- ✅ 智能进度条（根据实际步骤递进）
- ✅ 详细的预加载信息
- ✅ 文件映射过程可视化

### 性能优化
- ✅ 预加载源数据（跳过重复解析）
- ✅ 智能表头匹配
- ✅ 快速文件映射

## 🧪 测试建议

### 部署前测试
1. 解压到测试目录
2. 运行 `install.bat`
3. 运行 `test_server.bat`
4. 访问 http://localhost:8000
5. 上传测试文件验证功能

### 部署后验证
1. 检查IIS网站状态
2. 访问应用首页
3. 查看日志文件
4. 测试训练功能
5. 测试智算功能

## 🐛 常见问题速查

| 问题 | 原因 | 解决方案 |
|------|------|---------|
| 503错误 | 应用启动失败 | 查看logs\stdout.log |
| 500错误 | 配置错误 | 检查.env和权限 |
| 启动慢 | 首次加载依赖 | 增加startupTimeLimit |
| 端口占用 | 8000端口被占用 | 修改IIS绑定端口 |

## 📊 性能基准

### 启动时间
- 首次启动：30-60秒
- 后续启动：10-20秒

### 训练性能
- 简单规则：2-5分钟
- 复杂规则：5-15分钟
- 最大迭代：10次（可配置）

### 计算性能
- 小文件（<1000行）：5-10秒
- 中文件（1000-10000行）：10-30秒
- 大文件（>10000行）：30-60秒

## 🔒 安全检查清单

部署到生产环境前：

- [ ] 启用HTTPS
- [ ] 配置防火墙规则
- [ ] 设置IP白名单
- [ ] 启用身份验证
- [ ] 定期备份tenants目录
- [ ] 定期轮换API密钥
- [ ] 监控日志文件大小
- [ ] 配置日志轮转
- [ ] 设置应用程序池回收策略
- [ ] 测试灾难恢复流程

## 📞 技术支持

### 日志位置
```
应用日志：C:\inetpub\wwwroot\DataMerge\logs\stdout.log
IIS日志：C:\inetpub\logs\LogFiles\
```

### 手动测试命令
```cmd
cd C:\inetpub\wwwroot\DataMerge
venv\Scripts\activate
python -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
```

### 获取帮助
如遇问题，请提供：
1. 错误截图
2. logs\stdout.log 完整内容
3. IIS日志相关部分
4. 系统版本（Windows版本、IIS版本、Python版本）
5. 部署步骤说明

## 🎉 部署成功标志

当你看到以下内容时，说明部署成功：

1. ✅ 浏览器访问 http://localhost:8000 显示训练页面
2. ✅ 页面显示浅蓝色主题
3. ✅ 左侧导航菜单正常显示
4. ✅ 可以上传文件并开始训练
5. ✅ 实时日志正常显示
6. ✅ logs\stdout.log 无错误信息

---

**构建工具**：build_iis.bat
**版本**：v1.0.0
**构建日期**：2026-03-18
**支持平台**：Windows Server 2016+, IIS 10.0+, Python 3.11+

**祝部署顺利！** 🚀
