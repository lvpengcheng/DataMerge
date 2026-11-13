# IIS 方案运行修复

此次更新支持由 Linux Docker 导出的普通 Excel 计算方案在 Windows IIS 上运行。

- Python 完整版本、实现、CPU 架构、.NET Core 版本及 12 个计算依赖仍需匹配。
- AMD64 / x86_64 视为同一架构；操作系统、Python 二进制哈希、Web/开发/GPU 依赖不再要求完全一致。
- `calculation-constraints.txt` 与仓库的 Docker 锁定文件保持一致，`install.bat` 会按此安装。已有依赖版本满足时不重复安装。
- 使用本机受信任的启动层加载经完整性校验的冻结代码、模板和程序集。Linux 包缺少 Windows Skia 原生库时，仅在临时副本中补齐，并校验配套 SkiaSharp 程序集一致。
- 已有方案无需重新训练或改数据库；升级启动层后也能获得多文件临时目录清理修复。方案存储目录不会被修改。
- 清理被占用的临时工作簿会先回收对象并重试，持续占用则记录日志、留待任务目录清理，不覆盖真实计算异常，也不把已经交付成功的计算标成失败。
- 日志中的 `SOLUTION_RUNTIME` 记录原环境标识、实际环境标识及实际依赖。跨平台兼容不等于结果逐字节一致；上线后应使用已知输入核对金额、日期、公式及字体显示。脚本若自行使用 Linux 命令、绝对路径或额外第三方依赖，仍需做相应适配。

## 正式环境更新

停止正在运行的 DataMerge 后端，更新发布包中的应用文件，保留现有 `.env`、`tenants`、`data`、`solution_store` 和授权文件。不要用开发库覆盖正式数据库。

在站点目录执行 `install.bat`，然后重启 DataMerge 后端。若 IIS 使用 ARR 反向代理，还需确认监听 8000 的 Python 进程确实已重启。若使用 HttpPlatformHandler，检查 `web.config` 的站点路径。

如果使用旧离线 wheelhouse，需要补充 `python-dateutil==2.9.0.post0`、`pytz==2026.5`、`tzdata==2026.5` 等锁定版本，否则离线安装会明确失败。
