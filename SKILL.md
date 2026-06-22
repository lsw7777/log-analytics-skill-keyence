---
name: log-analytics-skill
description: "Azure Log Analytics 安全日志分析 Skill。自动查询 Azure Log Analytics 中的安全相关日志表，生成 HTML 风险报告，并使用 Azure MCP 深入分析高危、中危、低危项，输出 Markdown 格式的详细调查报告。适用于排查登录失败、可疑成功登录、Service Principal 对象和权限变动、许可证使用量、邮箱容量、DCR 采集错误、Intune 审计记录等安全场景。"
license: 专有
version: 2.0.0
author: Keyence IT
tags:
  - azure
  - log-analytics
  - security
  - kql
  - mcp
  - report
---

# Log Analytics 安全分析 Skill

## 概述

本 Skill 提供完整的 Azure Log Analytics 安全日志分析能力，包含两个阶段：

1. **HTML 报告生成**：运行 PowerShell 脚本查询指定时间范围内的日志数据，生成合并的 HTML 风险报告
2. **MCP 深度分析**：使用 Azure MCP 查看对应时间段的日志表数据，分析高危/中危/低危项，使用 KQL 深入调查，输出 Markdown 格式调查报告

## 触发条件

当用户提出以下类型的请求时，应调用本 Skill：

### 时间范围触发
- "查询最近N天的微软日志"
- "生成最近N天的 Log Analytics 风险报告"
- "查最近N小时的登录风险"
- "分析上周的日志安全"
- "last N days log analysis"

### 场景触发
- "检查登录失败情况"
- "查看可疑登录"
- "分析 Service Principal 权限变动"
- "检查许可证使用情况"
- "查看 DCR 采集错误"
- "Intune 审计记录分析"
- "邮箱容量风险检查"

## 执行流程

### 阶段一：生成 HTML 报告

#### 1. 解析用户输入的时间范围

从用户自然语言中提取时间范围，支持以下格式：

| 用户输入格式 | 解析结果 |
|-------------|---------|
| "最近N天" / "近N天" | 最近 N 天 |
| "last N days" | 最近 N 天 |
| "最近N小时" / "近N小时" | 最近 N 小时 |
| "last N hours" | 最近 N 小时 |
| "上周" | 最近 7 天 |
| "昨天" | 最近 1 天 |
| "今天" | 最近 24 小时 |

#### 2. 执行 main.ps1 脚本

在 Skill 根目录执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\main.ps1 -Prompt "<用户原始输入>" -SkipTotalCount -NoOpen
```

常用参数组合：

| 场景 | 命令 |
|------|------|
| 自然语言查询 | `.\scripts\main.ps1 -Prompt "查询最近15天的微软日志" -SkipTotalCount -NoOpen` |
| 指定表查询 | `.\scripts\main.ps1 -Prompt "查询最近1天的登录日志" -TableName "SigninLogs" -SkipTotalCount -NoOpen` |
| 自定义时间 | `.\scripts\main.ps1 -CustomStart "2026-06-10T00:00:00" -CustomEnd "2026-06-10T03:00:00" -SkipTotalCount -NoOpen` |
| 强制刷新 | `.\scripts\main.ps1 -Prompt "查询最近7天" -ForceRefresh -SkipTotalCount -NoOpen` |

#### 3. 获取报告路径

脚本执行成功后会输出：

```
HTML: <生成的 HTML 报告完整路径>
URL: file:///...
```

记录 HTML 报告路径，用于后续阶段。

### 阶段二：MCP 深度分析

#### 1. 使用 Azure MCP 查询日志表

使用 `use_mcp_tool` 调用 Azure MCP，对每个日志表执行 KQL 查询。

**查询的日志表**：
- `SigninLogs` - 用户登录日志
- `AADServicePrincipalSignInLogs` - 服务主体登录日志
- `AADManagedIdentitySignInLogs` - 托管身份登录日志
- `AuditLogs` - 审计日志
- `AssignedLicensesDCR_CL` - 许可证使用
- `DCRLogErrors` - 采集错误
- `MailboxStatisticsDCR_CL` - 邮箱统计
- `IntuneAuditLogsDCR_CL` - Intune 审计日志

**MCP 工具调用示例**：

```json
{
  "server_name": "Azure MCP Server",
  "tool_name": "log-analytics-execute-query",
  "arguments": {
    "workspace": "EntraID-workspace",
    "query": "<KQL 查询语句>",
    "timespan": "P7D"
  }
}
```

#### 2. 风险等级分类标准

对每个表的记录进行分析，按以下标准分类：

##### 高危 (High Risk)
| 类别 | 条件 | 表 |
|------|------|-----|
| 认证失败风暴 | 单 SP 失败率 > 90% 或失败次数 > 1000 | AADServicePrincipalSignInLogs |
| 密码暴力尝试 | 单 IP 对多用户尝试登录失败 (ResultType=50126) | SigninLogs |
| CA 策略阻断 | 被 Conditional Access 策略阻止 (ResultType=53003) | SigninLogs |
| SP 对象删除 | Service Principal 被删除或硬删除 | AuditLogs |
| 权限变动 | App Role Assignment 被添加或移除 | AuditLogs |
| 密钥/证书变动 | 应用证书或密钥被添加/更新/删除 | AuditLogs |

##### 中危 (Medium Risk)
| 类别 | 条件 | 表 |
|------|------|-----|
| 设备认证失败 | 设备认证失败 (ResultType=50155) | SigninLogs |
| 设备未注册 | 设备未在租户注册 (ResultType=700003) | SigninLogs |
| 设备被禁用 | 认证时设备被禁用 (ResultType=135011) | SigninLogs |
| 许可证异常 | 许可证字段缺失或总量为 0 | AssignedLicensesDCR_CL |
| 采集错误 | DCR 采集失败 | DCRLogErrors |
| 邮箱容量风险 | 可用空间 < 配额 5% | MailboxStatisticsDCR_CL |
| Shared Mailbox | 共享邮箱配置 | MailboxStatisticsDCR_CL |

##### 低危 (Low Risk)
| 类别 | 条件 | 表 |
|------|------|-----|
| KMSI 中断 | "保持登录"提示中断 (ResultType=50140) | SigninLogs |
| 消息提示中断 | 登录时需额外信息 (ResultType=50201) | SigninLogs |
| 流令牌过期 | 认证流程超时 (ResultType=50089) | SigninLogs |
| 托管身份登录 | 正常的托管身份登录记录 | AADManagedIdentitySignInLogs |

#### 3. 深入调查

对每个识别出的风险项，使用 KQL 进行深入调查：

**示例 KQL 查询**：

```kusto
-- 密码暴力尝试调查
SigninLogs
| where TimeGenerated > ago(7d) and ResultType == '50126'
| summarize AttemptCount=count(), FirstAttempt=min(TimeGenerated), LastAttempt=max(TimeGenerated)
  by UserPrincipalName, IPAddress
| where AttemptCount >= 3
| top 20 by AttemptCount desc

-- SP 认证失败调查
AADServicePrincipalSignInLogs
| where TimeGenerated > ago(7d)
  and ServicePrincipalName == 'AIP-DelegatedUser'
| extend __isFailed = tolower(ResultType) in ('false','fail','failed','1')
    or (tolower(ResultType) matches regex @"^\d+$" and toint(ResultType) != 0)
| where __isFailed
| summarize FailCount=count(), FirstFail=min(TimeGenerated), LastFail=max(TimeGenerated),
  SampleResult=take_any(ResultDescription)
  by IPAddress

-- CA 策略阻断调查
SigninLogs
| where TimeGenerated > ago(7d) and ResultType == '53003'
| summarize BlockCount=count(), LastBlock=max(TimeGenerated)
  by UserPrincipalName, AppDisplayName, IPAddress
| top 20 by BlockCount desc
```

#### 4. 生成 Markdown 调查报告

将调查结果写入 Markdown 文档，保存到 `mcp分析结果/` 目录。

**文件命名规则**：
```
Log-Analytics-<时间范围描述>-风险分析报告.md
```

**示例文件名**：
- `Log-Analytics-上周日志风险分析报告.md`
- `Log-Analytics-最近7天风险分析报告.md`
- `Log-Analytics-20260615-20260622风险分析报告.md`

**报告模板结构**：

```markdown
# Log Analytics <时间范围>风险分析报告

> **工作区**: <workspace名称> (<workspace_id>)
> **订阅**: <subscription_name> (<subscription_id>)
> **时间范围**: <开始日期> ~ <结束日期> (N天)
> **生成时间**: <生成时间>

---

## 总览

| 表名 | 记录数 | 高危 | 中危 | 低危 |
|------|--------|------|------|------|
| SigninLogs | ... | ... | ... | ... |
| AADServicePrincipalSignInLogs | ... | ... | ... | ... |
| ... | ... | ... | ... | ... |

---

## 1. SigninLogs (N 条)

**总量**: N | **失败**: N (X%) | **成功**: N

### 每日趋势

| 日期 | 总登录 | 失败数 |
|------|--------|--------|
| ... | ... | ... |

### 高危项

#### 1.1 <风险描述>

<详细说明>

| 用户 | 应用 | IP | 次数 | 最后时间 |
|------|------|-----|------|----------|
| ... | ... | ... | ... | ... |

**深入调查 KQL**:
```kusto
<KQL 查询语句>
```

### 中危项

...

### 低危项

...

---

## 风险汇总与处置建议

### 高危 (需立即处理)

| # | 问题 | 表 | 建议 |
|---|------|-----|------|
| 1 | ... | ... | ... |

### 中危 (需关注)

| # | 问题 | 表 | 建议 |
|---|------|-----|------|
| 1 | ... | ... | ... |

### 低危 (可观察)

| # | 问题 | 表 | 建议 |
|---|------|-----|------|
| 1 | ... | ... | ... |
```

## 脚本说明

### 核心脚本

| 脚本 | 用途 |
|------|------|
| `scripts/main.ps1` | 主入口脚本，负责时间范围解析、缓存管理、查询调度和报告生成 |
| `scripts/query-log-analytics.ps1` | 执行 Azure Log Analytics KQL 查询 |
| `scripts/log-analyzer-shared.ps1` | 共享函数库：表清单、时间范围、缓存、可信 IP、KQL 生成、字段解析 |
| `scripts/generate-html-report.ps1` | 读取 CSV 数据，按风险规则聚合，生成合并 HTML 报告 |
| `scripts/run-skill.ps1` | Skill 包装脚本，简化 Agent 调用 |

### 可信 IP 配置

可疑 IP 排除规则：
- `scripts/config/TrustedLocation_KJ.txt` - 可信位置 IP
- `scripts/config/TrustedLocation_IDC_Ali.txt` - IDC 阿里云 IP
- Microsoft Service Tags - Azure AD、Power BI、Azure Front Door、Microsoft Defender 等相关公网段

## 当前默认处理的表

| 表名 | 说明 | 关注点 |
|------|------|--------|
| AADManagedIdentitySignInLogs | 托管身份登录日志 | 登录失败、可信位置外成功登录 |
| AADServicePrincipalSignInLogs | 服务主体登录日志 | 登录失败(>10次聚合)、可信位置外成功登录 |
| AssignedLicensesDCR_CL | 许可证使用 | 4类许可证统计、Graph补齐总量 |
| AuditLogs | 审计日志 | SP对象变动、App Role权限变动、排除PIM噪声 |
| DCRLogErrors | 采集错误 | 最近30天去重统计 |
| IntuneAuditLogsDCR_CL | Intune审计日志 | 按Actor/Operation/Target提取 |
| MailboxStatisticsDCR_CL | 邮箱统计 | 可用空间<5%配额、SharedMailbox |
| SigninLogs | 用户登录日志 | 失败登录、可疑IP成功登录 |

## 维护规则

1. 新增表时，需同时更新：
   - `scripts/log-analyzer-shared.ps1` 的 `$SupportedLogTables`
   - 风险 KQL 生成函数
   - 字段解析规则
   - 报告聚合逻辑
   - 本 `SKILL.md` 中的表说明

2. 修改查询范围、风险条件、字段合并或 HTML 展示规则后，需运行测试并优先用最近 3 小时范围验证输出

3. 定期更新 `scripts/config/` 下的可信 IP 配置

## 输出产物

| 产物 | 位置 | 说明 |
|------|------|------|
| HTML 报告 | `html报告结果/` | 合并的风险 HTML 报告 |
| HTML 报告（桌面副本） | `~/Desktop/` | 自动复制到用户桌面 |
| Markdown 报告 | `mcp分析结果/` | MCP 深度分析的 Markdown 调查报告 |
| Markdown 报告（桌面副本） | `~/Desktop/` | 自动复制到用户桌面 |
| CSV 缓存 | `scripts/cache/` | 查询结果缓存 |

## 自动复制到桌面

每次生成报告后，Skill 会自动将文件复制到用户桌面，方便用户快速访问：

### 执行步骤

在 HTML 报告和 Markdown 报告生成完成后，执行以下操作：

#### 1. 复制 HTML 报告到桌面

```powershell
# 获取用户桌面路径
$desktopPath = [Environment]::GetFolderPath("Desktop")

# 复制 HTML 报告到桌面
Copy-Item -Path "<html报告路径>" -Destination $desktopPath -Force
```

#### 2. 复制 Markdown 报告到桌面

```powershell
# 复制 Markdown 报告到桌面
Copy-Item -Path "<md报告路径>" -Destination $desktopPath -Force
```

### 文件命名

复制到桌面的文件保持原文件名：
- HTML 报告：`final_report_merged_<时间范围>.html`
- Markdown 报告：`Log-Analytics-<时间范围>-风险分析报告.md`

### 注意事项

- 如果桌面已存在同名文件，会自动覆盖
- 原始文件仍保留在 `html报告结果/` 和 `mcp分析结果/` 目录中
- 桌面路径通过 `[Environment]::GetFolderPath("Desktop")` 自动获取，兼容不同系统配置
