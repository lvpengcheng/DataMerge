# 🎉 DataMerge IIS 部署包已就绪

## 📦 部署包信息

**文件名**：`DataMerge-IIS-v1.0.0.zip`
**位置**：`releases/DataMerge-IIS-v1.0.0.zip`
**大小**：806 KB
**构建日期**：2026-03-18
**版本**：v1.0.0

---

## 🚀 快速开始（3步）

### 第1步：解压部署包
```
解压到：C:\inetpub\wwwroot\DataMerge\
```

### 第2步：验证完整性
```cmd
cd C:\inetpub\wwwroot\DataMerge
verify_package.bat
```

### 第3步：按照文档部署
打开 `IIS_QUICK_START.md` 按步骤操作（5分钟完成）

---

## 📚 文档导航

| 文档 | 用途 | 适合人群 |
|------|------|---------|
| **IIS_QUICK_START.md** | 快速部署指南（5分钟） | 熟悉IIS的运维人员 |
| **PACKAGE_README.md** | 完整部署说明 | 所有部署人员 |
| **IIS_DEPLOYMENT.md** | 详细技术文档 | 需要深入了解的技术人员 |
| **README.md** | 项目介绍 | 了解项目功能 |

---

## ✅ 部署包内容

### 核心文件 ✓
- `backend/` - 后端Python代码
- `frontend/` - 前端页面（浅蓝色主题）
- `excel_parser.py` - Excel智能解析器
- `web.config` - IIS配置文件

### 配置文件 ✓
- `.env` - 环境变量配置
- `.env.example` - 配置示例
- `requirements.txt` - Python依赖

### 部署工具 ✓
- `install.bat` - 自动安装依赖
- `test_server.bat` - 本地测试
- `stop_server.bat` - 停止服务
- `verify_package.bat` - 验证完整性 ⭐新增

### 文档 ✓
- `PACKAGE_README.md` - 部署包说明
- `IIS_QUICK_START.md` - 快速指南
- `IIS_DEPLOYMENT.md` - 详细文档

---

## 🔧 系统要求

| 组件 | 最低版本 | 推荐版本 |
|------|---------|---------|
| Windows Server | 2016 | 2019/2022 |
| IIS | 10.0 | 10.0+ |
| Python | 3.11 | 3.11/3.12 |
| HttpPlatformHandler | 1.2 | 最新版 |
| 内存 | 4GB | 8GB+ |
| 磁盘 | 10GB | 20GB+ |

---

## 📋 部署检查清单

部署前请确认：

**环境准备**
- [ ] Windows Server 2016+ 或 Windows 10+
- [ ] IIS 已安装并启用
- [ ] Python 3.11+ 已安装
- [ ] HttpPlatformHandler 已安装

**文件准备**
- [ ] 部署包已解压到 `C:\inetpub\wwwroot\DataMerge\`
- [ ] 运行 `verify_package.bat` 验证完整性
- [ ] 运行 `install.bat` 安装依赖

**配置准备**
- [ ] 编辑 `.env` 文件，配置API密钥
- [ ] 检查 `web.config` 中的路径是否正确

**IIS配置**
- [ ] 创建应用程序池（无托管代码）
- [ ] 创建网站并绑定端口
- [ ] 设置 IIS_IUSRS 用户组权限（修改权限）
- [ ] 启动网站

**验证部署**
- [ ] 访问 http://localhost:8000
- [ ] 查看 `logs\stdout.log` 无错误
- [ ] 测试上传文件功能
- [ ] 测试训练功能

---

## 🎯 部署流程图

```
1. 解压部署包
   ↓
2. 运行 verify_package.bat（验证完整性）
   ↓
3. 运行 install.bat（安装Python依赖）
   ↓
4. 编辑 .env（配置API密钥）
   ↓
5. 安装 HttpPlatformHandler
   ↓
6. 配置IIS（应用程序池 + 网站）
   ↓
7. 设置文件夹权限（IIS_IUSRS）
   ↓
8. 启动网站
   ↓
9. 访问测试（http://localhost:8000）
   ↓
10. 部署完成 ✓
```

---

## 🧪 测试步骤

### 本地测试（推荐先测试）
```cmd
cd C:\inetpub\wwwroot\DataMerge
test_server.bat
```
访问：http://localhost:8000

### IIS测试
1. 在IIS管理器中启动网站
2. 访问：http://localhost:8000
3. 检查页面是否正常显示
4. 上传测试文件进行训练
5. 查看实时日志是否正常

---

## 🐛 常见问题

### Q1: 503 Service Unavailable
**A**: 应用启动失败，查看 `logs\stdout.log` 获取详细错误信息

### Q2: 应用启动很慢
**A**: 首次启动需要加载依赖，等待1-2分钟。可增加 `web.config` 中的 `startupTimeLimit` 到 120

### Q3: 权限错误
**A**: 确保 IIS_IUSRS 用户组对 DataMerge 文件夹有"修改"权限

### Q4: Python未找到
**A**: 检查 `web.config` 中的 `processPath` 是否指向正确的Python路径

### Q5: API密钥无效
**A**: 检查 `.env` 文件中的API密钥是否正确配置

---

## 📊 性能参考

### 启动时间
- 首次启动：30-60秒
- 后续启动：10-20秒

### 训练时间
- 简单规则（1-2个源文件）：2-5分钟
- 中等规则（3-5个源文件）：5-10分钟
- 复杂规则（5+个源文件）：10-15分钟

### 计算时间
- 小数据（<1000行）：5-10秒
- 中数据（1000-10000行）：10-30秒
- 大数据（>10000行）：30-60秒

---

## 🔒 安全建议

### 生产环境必做
1. ✅ 启用HTTPS（配置SSL证书）
2. ✅ 配置防火墙规则
3. ✅ 设置强密码策略
4. ✅ 定期备份 `tenants` 目录
5. ✅ 定期轮换API密钥

### 可选优化
- 配置IP白名单
- 启用身份验证
- 配置日志轮转
- 监控应用程序池状态

---

## 📞 获取帮助

### 查看日志
```
应用日志：C:\inetpub\wwwroot\DataMerge\logs\stdout.log
IIS日志：C:\inetpub\logs\LogFiles\
```

### 手动测试
```cmd
cd C:\inetpub\wwwroot\DataMerge
venv\Scripts\activate
python -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
```

### 报告问题时请提供
1. 错误截图
2. `logs\stdout.log` 完整内容
3. IIS日志相关部分
4. 系统信息（Windows版本、IIS版本、Python版本）
5. 部署步骤说明

---

## ✨ 新功能亮点

### 前端优化
- 🎨 浅蓝色主题设计
- 📱 左侧导航菜单
- 🔍 移除租户ID显示
- 🎯 统一的视觉风格

### 流式日志增强
- 📡 实时显示后端日志
- 📊 智能进度条（根据实际步骤）
- 📝 详细的预加载信息
- 🔄 文件映射过程可视化

### 性能优化
- ⚡ 预加载源数据（跳过重复解析）
- 🎯 智能表头匹配
- 🚀 快速文件映射

---

## 🎉 部署成功标志

当你看到以下内容时，说明部署成功：

1. ✅ 浏览器访问 http://localhost:8000 显示训练页面
2. ✅ 页面显示浅蓝色主题
3. ✅ 左侧导航菜单正常显示（训练、智算）
4. ✅ 可以上传文件并开始训练
5. ✅ 实时日志正常显示
6. ✅ `logs\stdout.log` 无错误信息

---

**祝部署顺利！** 🚀

如有问题，请参考包内的详细文档或查看日志文件。

---

**构建工具**：`build_iis.bat`
**版本**：v1.0.0
**构建日期**：2026-03-18
**支持平台**：Windows Server 2016+, IIS 10.0+, Python 3.11+
