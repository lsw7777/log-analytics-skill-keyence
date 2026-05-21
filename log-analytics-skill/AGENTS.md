# AGENTS.md — log-skill

## 项目概述

单文件仓库 (`azure_log_query.ps1`)，用于查询 Azure Log Analytics (中国云) 中的 Office365 审计日志。
目标是构建 OpenCode skills 来分析这些日志。

## 关键信息

### Azure 连接详情
- **云环境**: `AzureChinaCloud`
- **Workspace ID**: `703a5771-97fc-4bf3-a585-f607d18c4479`
- **Tenant ID**: `420c4dab-8603-402f-afe0-75bc28c51c13`
- **认证**: 通过 `Connect-AzAccount` 交互式浏览器登录，支持 `-ForceLogin`
- **模块**: `Az.OperationalInsights` — 通过 `Invoke-AzOperationalInsightsQuery` 执行查询

### 支持的日志表 (DCR_CL)
- `AuditGeneralDCR_CL` — Office 365 通用审计日志 (默认)
- `SharePointAuditDCR_CL` — SharePoint 审计日志
- `MessageTraceDataDCR_CL` — 邮件追踪数据
- `AssignedLicensesDCR_CL` — 已分配许可证信息
- `AzureADUsersDCR_CL` — Azure AD 用户信息
- `MailboxStatisticsDCR_CL` — 邮箱统计信息
- `WQCLogDCR_CL` — WQC 日志

**编写分析代码前，务必先运行快速架构探测查询来发现当前字段** — 字段会随着数据管道的更新而发生变化。

### 运行脚本
```powershell
# 快速探测 — 发现当前字段 (默认表)
.\azure_log_query.ps1 -Query "AuditGeneralDCR_CL | take 1"

# 使用 -TableName 参数指定表
.\azure_log_query.ps1 -TableName "SharePointAuditDCR_CL" -Query "SharePointAuditDCR_CL | take 1"

# 按时间范围查询
.\azure_log_query.ps1 -Query "<KQL>" -Hours 24
.\azure_log_query.ps1 -ForceLogin        # 401/403 时重新认证
.\azure_log_query.ps1 -UseDeviceCode     # 设备代码流

# 导出到 CSV
.\azure_log_query.ps1 -TableName "AuditGeneralDCR_CL" -Hours 24 -ExportCsv ".\General_20260521.csv"
```

### 待构建的 Skills
1. **批量日志分析器** — 处理大量数据集，计算统计数据，生成独立的 HTML 报告
2. **交互式 HTML 查询工具** — 按用户/条件过滤，数据嵌入 HTML，提供客户端筛选 UI

## 开发指南

- PowerShell 脚本 (`azure_log_query.ps1`) 是唯一的数据源
- 每次分析日志前，先用 `| take 1` 探测字段结构
- Office365 审计日志对象包含标准字段如 `UserId`, `Operation`, `Workload`, `CreationDate`, `ClientIP` 等
- 生成的 HTML 文件应该是自包含的 (CSS/JS 内联)，无需外部依赖
- KQL 查询使用 `AuditGeneralDCR_CL` 表名 (注意 `_CL` 后缀)
