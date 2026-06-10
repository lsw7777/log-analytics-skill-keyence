# 批量日志分析器 - 定时执行功能

## 更新内容

本项目为 10 个日志分析表提供定时执行工作流。

关键文件：
- `scripts/install-scheduler.ps1` - 安装每日定时任务并写入调度配置
- `scripts/scheduled-run.ps1` - 按配置的表集合执行分析，不打开浏览器
- `scripts/tray.ps1` - 在 Windows 通知区域提供托盘等待进程
- `scripts/run-all.ps1` - 支持 `-NoOpen` 参数用于无人值守执行

默认行为：
- 每天 `01:00` 执行
- 处理全部 10 个表
- 生成 HTML 报告，定时模式下不强制弹出浏览器

## 文件位置

- `scripts/install-scheduler.ps1`
- `scripts/scheduled-run.ps1`
- `scripts/tray.ps1`
- `scripts/run-all.ps1`

## 默认安装

在项目根目录运行：

```powershell
.\scripts\install-scheduler.ps1
```

这将：
- 创建包含默认值的配置文件
- 注册名为 `BulkLogAnalyzerDaily` 的每日定时任务
- 安排在 `01:00` 执行
- 使用全部 10 个表

## 自定义安装

你可以选择不同的运行时间和表子集：

```powershell
.\scripts\install-scheduler.ps1 -RunAt "03:30" -Tables @("AuditGeneralDCR_CL", "SigninLogs")
.\scripts\install-scheduler.ps1 -RunAt "17:20" -Tables @("AuditLogs", "MessageTraceDataDCR_CL")
```

如需同时在启动文件夹中创建托盘快捷方式：

```powershell
.\scripts\install-scheduler.ps1 -RunAt "03:30" -Tables @("AuditGeneralDCR_CL", "SigninLogs") -CreateTrayShortcut
.\scripts\install-scheduler.ps1 -RunAt "17:20" -Tables @("AuditLogs", "MessageTraceDataDCR_CL") -CreateTrayShortcut
```

规则：
- `-RunAt` 必须采用 `HH:mm`  格式
- `-Tables` 可包含任意受支持的表名
- 如果省略 `-Tables`，则使用全部 10 个表


## 工作原理

1. `install-scheduler.ps1` 写入 `schedule-config.json` 配置文件
2. 它注册一个 Windows 定时任务
3. 在预定时间，Windows 启动 scheduled-run.ps1
4. `scheduled-run.ps1` 对每个表依次调用 `run-all.ps1 -NoOpen`
5. `run-all.ps1` 获取日志、进行分析并生成 HTML 报告
6. 如果启用了托盘模式，`tray.ps1` 会在通知区域保持一个可见的托盘进程，等待下一次执行


## 手动使用脚本

### 安装或更新定时任务

```powershell
.\scripts\install-scheduler.ps1
```

### 直接运行计划批处理

```powershell
.\scripts\scheduled-run.ps1 -ConfigPath .\scripts\schedule-config.json
```

### 启动系统托盘进程

```powershell
.\scripts\tray.ps1 -ConfigPath .\scripts\schedule-config.json
```

### 不打开浏览器运行单个表

```powershell
.\scripts\run-all.ps1 -TableName "AuditGeneralDCR_CL" -AnalysisDate "2026-05-26" -NoOpen
```

## 计划任务名称

- Task name: `BulkLogAnalyzerDaily`

## 输出

Generated HTML reports are written to the repository root as `final_report_*.html`.
生成的 HTML 报告将写入仓库根目录，文件名为 `final_report_*.html`。

## Notes

- 计划执行时使用 `-NoOpen` 参数，以确保在无人值守运行时不会打开浏览器。
- 如果某个表的查询返回零行，则不会持久化对应的缓存。
- 空的缓存数据将被视为无效，并会被自动删除。
