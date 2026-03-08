# 邮件自动处理API文档

## 功能概述

邮件自动处理功能可以自动从邮箱接收邮件，解析附件，匹配训练好的脚本，自动执行计算并发送结果邮件。

## 工作流程

1. **接收邮件**：从配置的邮箱账户中获取新邮件（使用POP3协议）
2. **解析主题**：检查邮件主题是否符合格式 `{租户名称}_{薪资年}_{薪资月}`（如：`达美乐_2025_12`）
3. **检查租户**：验证租户是否有训练成功的脚本
4. **保存附件**：将邮件中的Excel附件保存到临时目录
5. **匹配文件**：根据训练时的文件结构，匹配并重命名附件
6. **检查完整性**：确认文件数量与训练时一致
7. **获取工时**：调用AI获取该年月的法定工作小时
8. **执行计算**：调用计算接口处理数据
9. **发送结果**：将计算结果通过邮件发送给指定收件人（使用SMTP协议）

## API接口

### 1. 添加邮件账户

**接口**: `POST /api/email/add-account`

**参数**:
- `email_address` (string, 必填): 邮箱地址
- `pop3_server` (string, 必填): POP3服务器地址
  - 示例: `pop.qq.com`, `pop.163.com`, `pop.gmail.com`
- `pop3_port` (int, 必填): POP3端口号
  - POP3 SSL: 995
  - POP3非SSL: 110
- `pop3_ssl` (bool, 默认true): POP3是否使用SSL加密
- `pop3_password` (string, 必填): POP3密码或授权码
- `smtp_server` (string, 必填): SMTP服务器地址
  - 示例: `smtp.qq.com`, `smtp.163.com`, `smtp.gmail.com`
- `smtp_port` (int, 必填): SMTP端口号
  - SMTP SSL: 465
  - SMTP非SSL: 25
- `smtp_ssl` (bool, 默认true): SMTP是否使用SSL加密
- `smtp_password` (string, 必填): SMTP密码或授权码
- `recipients` (string, 可选): 收件人列表，多个邮箱用逗号分隔

**示例请求**:
```bash
curl -X POST "http://localhost:8000/api/email/add-account" \
  -F "email_address=lupech@163.com" \
  -F "pop3_server=pop.163.com" \
  -F "pop3_port=995" \
  -F "pop3_ssl=true" \
  -F "pop3_password=YRbEUpDrNDTpCtHw" \
  -F "smtp_server=smtp.163.com" \
  -F "smtp_port=465" \
  -F "smtp_ssl=true" \
  -F "smtp_password=YRbEUpDrNDTpCtHw" \
  -F "recipients=lupech@163.com,manager@company.com"
```

**响应**:
```json
{
  "success": true,
  "message": "邮箱账户已添加"
}
```

### 2. 检查邮件并处理

**接口**: `POST /api/email/check`

**说明**: 检查所有配置的邮箱账户，处理新邮件

**示例请求**:
```bash
curl -X POST "http://localhost:8000/api/email/check"
```

**响应**:
```json
{
  "success": true,
  "message": "检查完成，处理了 2 封邮件",
  "results": [
    {
      "success": true,
      "tenant_name": "达美乐",
      "salary_year": 2025,
      "salary_month": 12,
      "matched_files": 14,
      "monthly_standard_hours": 174.0,
      "message": "文件已准备好，可以开始计算",
      "calculation_result": {
        "success": true,
        "download_url": "/api/download/result.xlsx?tenant_id=达美乐&batch_id=202512"
      }
    }
  ]
}
```

## 邮件主题格式

邮件主题必须严格遵循以下格式：

```
{租户名称}_{年份}_{月份}
```

**示例**:
- `达美乐_2025_12` ✅
- `达美乐_2025_1` ✅ (月份可以是1位数)
- `达美乐2025_12` ❌ (缺少下划线)
- `达美乐_25_12` ❌ (年份必须是4位数)

## 附件要求

1. **文件格式**: 只支持 `.xlsx` 和 `.xls` 格式
2. **文件数量**: 必须与训练时的源文件数量一致
3. **文件结构**: 表头结构需要与训练时的文件匹配（匹配度>80%）

## 配置文件

邮件账户配置保存在 `email_config.json` 文件中：

```json
{
  "last_check_time": "2026-03-03T10:00:00",
  "email_accounts": [
    {
      "email_address": "lupech@163.com",
      "pop3_server": "pop.163.com",
      "pop3_port": 995,
      "pop3_ssl": true,
      "pop3_password": "YRbEUpDrNDTpCtHw",
      "smtp_server": "smtp.163.com",
      "smtp_port": 465,
      "smtp_ssl": true,
      "smtp_password": "YRbEUpDrNDTpCtHw",
      "recipients": ["lupech@163.com", "manager@company.com"],
      "last_check_time": "2026-03-03T10:00:00"
    }
  ]
}
```

## 常见邮箱配置

### QQ邮箱
- POP3服务器: `pop.qq.com`
- POP3端口: 995 (SSL)
- SMTP服务器: `smtp.qq.com`
- SMTP端口: 465 (SSL)
- 需要开启POP3/SMTP服务并获取授权码

### 163邮箱
- POP3服务器: `pop.163.com`
- POP3端口: 995 (SSL)
- SMTP服务器: `smtp.163.com`
- SMTP端口: 465 (SSL)
- 需要开启POP3/SMTP服务并获取授权码

### Gmail
- POP3服务器: `pop.gmail.com`
- POP3端口: 995 (SSL)
- SMTP服务器: `smtp.gmail.com`
- SMTP端口: 465 (SSL)
- 需要开启"允许不够安全的应用"或使用应用专用密码

### 企业邮箱
- 请咨询邮箱服务提供商获取POP3/SMTP服务器地址和端口

## 定时任务配置

建议使用系统定时任务（如cron）定期调用检查接口：

### Linux/Mac (crontab)
```bash
# 每小时检查一次
0 * * * * curl -X POST http://localhost:8000/api/email/check

# 每30分钟检查一次
*/30 * * * * curl -X POST http://localhost:8000/api/email/check
```

### Windows (任务计划程序)
1. 打开"任务计划程序"
2. 创建基本任务
3. 触发器：每小时/每30分钟
4. 操作：启动程序
   - 程序：`curl.exe`
   - 参数：`-X POST http://localhost:8000/api/email/check`

## 文件存储结构

```
tenants/
  └── {租户名}/
      └── calculations/
          └── {年份}{月份}/
              ├── temp/              # 临时目录（处理后删除）
              │   └── *.xlsx         # 原始附件
              ├── *.xlsx             # 重命名后的文件
              └── output/
                  └── result.xlsx    # 计算结果
```

## 错误处理

### 常见错误

1. **邮件主题格式不匹配**
   - 错误: `邮件主题格式不匹配`
   - 解决: 确保主题格式为 `租户名_年份_月份`

2. **租户没有训练脚本**
   - 错误: `租户 xxx 没有训练成功的脚本`
   - 解决: 先完成训练，确保有活跃脚本

3. **文件数量不一致**
   - 错误: `文件数量不一致: 实际 10, 预期 14`
   - 解决: 确保邮件附件包含所有必需的文件

4. **文件匹配失败**
   - 错误: 某些文件无法匹配
   - 解决: 检查文件表头结构是否与训练时一致

5. **邮箱连接失败**
   - 错误: `获取邮件失败`
   - 解决: 检查服务器地址、端口、密码是否正确

## 安全建议

1. **使用授权码**: 不要使用邮箱登录密码，使用邮箱服务商提供的授权码
2. **启用SSL**: 始终使用SSL加密连接
3. **限制收件人**: 只添加可信的收件人邮箱
4. **定期更新密码**: 定期更换授权码
5. **监控日志**: 定期检查处理日志，发现异常及时处理

## 日志查看

所有邮件处理日志会记录在应用日志中：

```bash
# 查看最近的邮件处理日志
tail -f logs/app.log | grep "email"
```

## 测试建议

1. **先测试连接**: 使用邮箱客户端（如Outlook、Thunderbird）测试POP3/SMTP连接
2. **发送测试邮件**: 发送一封符合格式的测试邮件
3. **手动触发检查**: 调用 `/api/email/check` 接口
4. **查看日志**: 检查处理日志，确认流程正常
5. **验证结果**: 检查计算结果和结果邮件

## 性能优化

1. **批量处理**: 一次检查可以处理多封邮件
2. **增量检查**: 只获取上次检查后的新邮件
3. **异步处理**: 邮件处理和计算都是异步执行
4. **自动清理**: 临时文件处理后自动删除

## 扩展功能

未来可以扩展的功能：

1. **邮件模板**: 自定义结果邮件模板
2. **错误通知**: 处理失败时发送通知邮件
3. **多租户支持**: 一封邮件包含多个租户的数据
4. **附件压缩**: 支持压缩包附件
5. **邮件归档**: 自动归档已处理的邮件
