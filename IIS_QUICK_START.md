# DataMerge IIS 快速部署指南

## 一、准备工作

### 1.1 系统要求
- Windows Server 2016+ 或 Windows 10+
- IIS 10.0+
- Python 3.11+
- 至少 4GB 可用磁盘空间

### 1.2 下载安装包
解压 `DataMerge-IIS-v1.0.0.zip` 到：
```
C:\inetpub\wwwroot\DataMerge\
```

## 二、快速安装（5分钟）

### 步骤1：安装Python依赖
双击运行 `install.bat`，等待安装完成。

### 步骤2：配置API密钥
编辑 `.env` 文件，填入你的AI API密钥：
```env
OPENAI_API_KEY=sk-xxxxx
```

### 步骤3：安装HttpPlatformHandler
下载并安装：https://www.iis.net/downloads/microsoft/httpplatformhandler

### 步骤4：配置IIS

#### 4.1 创建应用程序池
1. 打开 IIS 管理器
2. 右键"应用程序池" → "添加应用程序池"
3. 设置：
   - 名称：`DataMerge`
   - .NET CLR版本：`无托管代码`
   - 托管管道模式：`集成`

#### 4.2 配置应用程序池
右键 `DataMerge` 应用程序池 → 高级设置：
- 空闲超时：`0`（禁用）
- 固定时间间隔：`0`（禁用自动回收）

#### 4.3 创建网站
1. 右键"网站" → "添加网站"
2. 设置：
   - 网站名称：`DataMerge`
   - 应用程序池：`DataMerge`
   - 物理路径：`C:\inetpub\wwwroot\DataMerge`
   - 端口：`8000`

#### 4.4 设置权限
右键 `C:\inetpub\wwwroot\DataMerge` → 属性 → 安全：
- 添加 `IIS_IUSRS` 用户组
- 授予"修改"权限

### 步骤5：启动服务
在IIS管理器中，右键网站 → 管理网站 → 启动

### 步骤6：验证
浏览器访问：http://localhost:8000

## 三、故障排查

### 问题1：503 Service Unavailable
**原因**：应用启动失败

**解决**：
1. 查看日志：`C:\inetpub\wwwroot\DataMerge\logs\stdout.log`
2. 检查Python路径是否正确
3. 手动测试：
```cmd
cd C:\inetpub\wwwroot\DataMerge
venv\Scripts\activate
python -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
```

### 问题2：500 Internal Server Error
**原因**：配置错误或权限不足

**解决**：
1. 检查 `.env` 文件是否配置正确
2. 确认 IIS_IUSRS 有修改权限
3. 查看应用日志

### 问题3：应用启动慢
**原因**：首次启动需要加载依赖

**解决**：
- 首次启动等待1-2分钟
- 增加 `web.config` 中的 `startupTimeLimit` 到 120

## 四、生产环境配置

### 4.1 启用HTTPS
1. 在IIS中导入SSL证书
2. 修改网站绑定为 https:443

### 4.2 配置域名
1. 在DNS中添加A记录指向服务器IP
2. 在IIS网站绑定中添加域名

### 4.3 性能优化
- 定期清理 `logs` 目录
- 定期备份 `tenants` 目录
- 监控应用程序池内存使用

## 五、日常维护

### 查看日志
```
应用日志：C:\inetpub\wwwroot\DataMerge\logs\stdout.log
IIS日志：C:\inetpub\logs\LogFiles\
```

### 重启服务
在IIS管理器中：
1. 右键网站 → 管理网站 → 停止
2. 右键网站 → 管理网站 → 启动

### 更新应用
1. 停止IIS网站
2. 备份 `tenants` 目录和 `.env` 文件
3. 解压新版本覆盖
4. 恢复 `tenants` 和 `.env`
5. 运行 `install.bat` 更新依赖
6. 启动IIS网站

## 六、技术支持

如遇问题，请提供：
1. `logs\stdout.log` 日志文件
2. IIS日志文件
3. 错误截图
4. 操作系统版本和IIS版本

---

**提示**：首次部署建议先使用 `test_server.bat` 测试应用是否正常运行，确认无误后再配置IIS。
