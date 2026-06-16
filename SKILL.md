---
name: log-analytics-skill
description: "用于查询 Azure Log Analytics 中的安全相关日志表，生成合并后的风险 HTML 报告。适用于排查登录失败、可疑成功登录、Service Principal 对象和权限变动、许可证使用量、邮箱容量、DCR 采集错误、Intune 审计记录。"
license: 专有
---

# Log Analytics 安全风险报告

## 用途

本 skill 从 Azure Log Analytics 查询指定时间范围内的日志数据，并生成一个合并 HTML 风险报告。报告只关注失败、异常、可疑、容量不足、删除、禁用、Service Principal 对象或权限变动、DCR 采集错误、Intune 审计记录等需要处理的信息。

脚本目录为 `scripts`，HTML 报告模板已内嵌在 `scripts/generate-html-report.ps1`，可信 IP 配置在 `scripts/config`。根目录只保留这一个 `SKILL.md`。

## 运行方式

### Agent / OpenCode 自然语言启动

当用户在 OpenCode 或其他支持 Skill 的 Agent 中提出类似下面的自然语言请求时，应直接调用本 skill 的包装脚本，不要再要求用户手动选择时间：

```text
查询最近15天的微软日志
生成最近7天的 Log Analytics 风险报告
查最近3小时的登录风险
```

Agent 应在 skill 根目录执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\run-skill.ps1 "查询最近15天的微软日志"
```

`run-skill.ps1` 会把自然语言传给 `main.ps1 -Prompt`，自动解析 `最近 n 天`、`近 n 天`、`last n days`、`最近 n 小时`、`last n hours`，并默认使用 `-SkipTotalCount -NoOpen`，避免在 Agent 环境中卡在总数预检查或弹出浏览器。脚本结束后会在输出中打印：

```text
HTML: <生成的 HTML 报告完整路径>
URL: file:///...
```

Agent 需要把 `HTML:` 后面的路径明确告诉用户，例如“报告已生成在：...”。

也可以直接调用 `main.ps1`：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\main.ps1 -Prompt "查询最近15天的微软日志" -SkipTotalCount -NoOpen
```

如果用户明确指定表，可以加 `-TableName`，例如：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\run-skill.ps1 "查询最近1天的登录日志" -TableName "SigninLogs"
```

### 手工启动

在 skill 根目录运行：

```powershell
.\scripts\main.ps1
```

脚本会提示输入时间范围：

```text
0 = 最近 3 小时，用于测试
1 = 最近 1 天
n = 最近 n 天，最大 90 天
```

常用参数：

```powershell
.\scripts\main.ps1 -TableName "SigninLogs" -CustomStart "2026-06-10T00:00:00" -CustomEnd "2026-06-10T03:00:00"
.\scripts\main.ps1 -StartDate "2026-06-01" -EndDate "2026-06-03"
.\scripts\main.ps1 -ForceRefresh
.\scripts\main.ps1 -NoRiskFilter
```

## 当前默认处理的表

```text
AADManagedIdentitySignInLogs
AADServicePrincipalSignInLogs
AssignedLicensesDCR_CL
AuditLogs
DCRLogErrors
IntuneAuditLogsDCR_CL
MailboxStatisticsDCR_CL
SigninLogs
```

不再默认查询或报告：

```text
AuditGeneralDCR_CL
AzureADUsersDCR_CL
MessageTraceDataDCR_CL
SharePointAuditDCR_CL
WQCLogDCR_CL
```

## 报告关注点

`AADManagedIdentitySignInLogs`：托管身份登录失败，以及可信位置外的成功登录。

`AADServicePrincipalSignInLogs`：服务主体登录失败只展示超过 10 次的聚合记录，并明确展示 `ServicePrincipalName`；同时展示可信位置外的成功登录。

`AssignedLicensesDCR_CL`：统计 4 类许可证名称、已使用数量、总量、剩余数量；日志缺少总量时通过 Microsoft Graph `subscribedSkus` 补齐。

`AuditLogs`：只展示操作者是用户的条目，排除 PIM 噪声；重点关注 Service Principal 对象变动和 app role assignment 权限变动。

`DCRLogErrors`：按最近 30 天的 `InputStreamId`、`OperationName`、`Message` 去重统计采集错误。

`IntuneAuditLogsDCR_CL`：展示所选时间范围内的 Intune 审计记录，兼容自定义日志常见的 `_s` 后缀字段，并按 `Actor`、`Operation`、`Target` 提取。

`MailboxStatisticsDCR_CL`：在查询端保留可用空间低于配额 5% 的异常邮箱，以及字段标识为 SharedMailbox 的邮箱；报告中分别展示邮箱容量风险和 SharedMailbox。

`SigninLogs`：关注失败登录，以及 IP 不在可信位置内且登录应用不属于 `Windows Sign In`、`Microsoft Edge`、`Sangfor SASE VPN`、`Microsoft Office` 的成功登录。可疑 IP 栏只从 `SigninLogs` 产生。

## 可信 IP 规则

可疑 IP 会排除：

```text
scripts/config/TrustedLocation_KJ.txt
scripts/config/TrustedLocation_IDC_Ali.txt
Microsoft Service Tags 中与 Azure AD、Power BI、Azure Front Door、Microsoft Defender、Microsoft Cloud App Security 等相关的公网段
```

脚本会对日志中的 IP 做标准化，排除私网、回环、链路本地、未指定地址和广播地址。客户端 IP 排行不会包含已经出现在“可疑 IP”栏中的 IP。

## 脚本链路

`scripts/main.ps1` 负责选择时间范围、选择表、管理缓存、调用查询脚本并触发报告生成。

`scripts/query-log-analytics.ps1` 负责登录 Azure China Cloud，并使用 `Invoke-AzOperationalInsightsQuery` 查询 Log Analytics。

`scripts/log-analyzer-shared.ps1` 保存表清单、时间范围、缓存、可信 IP、KQL 生成、字段解析和报告路径等公共逻辑。

`scripts/generate-html-report.ps1` 读取 CSV，按风险规则聚合，生成合并 HTML 报告。

## KQL 模板

下面的 KQL 用于说明当前风险预过滤逻辑。实际执行时，`{StartUtc}`、`{EndUtc}` 和 `{TrustedIpCidrs}` 会由脚本动态替换。

### AADManagedIdentitySignInLogs

```kusto
let __base =
AADManagedIdentitySignInLogs
| where TimeGenerated >= datetime({StartUtc}) and TimeGenerated < datetime({EndUtc})
| extend ServicePrincipalName = tostring(coalesce(column_ifexists("ServicePrincipalName", ""), column_ifexists("ManagedIdentityName", ""), column_ifexists("Identity", ""), column_ifexists("AppDisplayName", ""), column_ifexists("ServicePrincipalId", ""), column_ifexists("AppId", ""), "Unknown"))
| extend UserPrincipalName = ServicePrincipalName
| extend ResourceDisplayName = tostring(coalesce(column_ifexists("ResourceDisplayName", ""), column_ifexists("ResourceIdentity", ""), column_ifexists("ResourceServicePrincipalId", ""), ""))
| extend IPAddress = tostring(coalesce(column_ifexists("IPAddress", ""), column_ifexists("IpAddress", ""), column_ifexists("ClientIP", ""), column_ifexists("ClientIpAddress", ""), ""))
| extend ResultType = tostring(coalesce(column_ifexists("ResultType", ""), column_ifexists("Status", ""), column_ifexists("ResultDescription", ""), ""))
| extend ResultDescription = tostring(coalesce(column_ifexists("ResultDescription", ""), column_ifexists("FailureReason", ""), column_ifexists("Status", ""), column_ifexists("ConditionalAccessStatus", ""), ""))
| extend __status = tolower(ResultType)
| extend __ip = extract(@"(?<!\d)(\d{1,3}(?:\.\d{1,3}){3})(?!\d)", 1, IPAddress)
| extend __isFailed = (__status in ("false","fail","failed","failure","denied","error","timeout","1") or (__status matches regex @"^\d+$" and toint(__status) != 0) or (tolower(ResultDescription) matches regex @"\b(fail|failed|failure|denied|error|timeout)\b"))
| extend __isSuccess = (__status in ("true","success","succeeded","completed","complete","ok","pass","passed","0"))
| extend __isPublicUntrustedIp = isnotempty(__ip) and not(ipv4_is_in_any_range(__ip, dynamic({TrustedIpCidrs})));
let __failed = __base | where __isFailed | summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by UserPrincipalName, ServicePrincipalName, ResourceDisplayName, IPAddress, ResultType, ResultDescription;
let __suspicious = __base | where __isSuccess and __isPublicUntrustedIp | summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count(), ResultType=take_any(ResultType), ResultDescription=take_any(ResultDescription) by UserPrincipalName, ServicePrincipalName, ResourceDisplayName, IPAddress;
__failed | union isfuzzy=true __suspicious
```

### AADServicePrincipalSignInLogs

```kusto
let __base =
AADServicePrincipalSignInLogs
| where TimeGenerated >= datetime({StartUtc}) and TimeGenerated < datetime({EndUtc})
| extend ServicePrincipalName = tostring(coalesce(column_ifexists("ServicePrincipalName", ""), column_ifexists("AppDisplayName", ""), column_ifexists("ServicePrincipalId", ""), column_ifexists("AppId", ""), "Unknown"))
| extend UserPrincipalName = ServicePrincipalName
| extend ResourceDisplayName = tostring(coalesce(column_ifexists("ResourceDisplayName", ""), column_ifexists("ResourceServicePrincipalId", ""), ""))
| extend IPAddress = tostring(coalesce(column_ifexists("IPAddress", ""), column_ifexists("IpAddress", ""), column_ifexists("ClientIP", ""), column_ifexists("ClientIpAddress", ""), ""))
| extend ResultType = tostring(coalesce(column_ifexists("ResultType", ""), column_ifexists("Status", ""), column_ifexists("ResultDescription", ""), ""))
| extend ResultDescription = tostring(coalesce(column_ifexists("ResultDescription", ""), column_ifexists("FailureReason", ""), column_ifexists("Status", ""), column_ifexists("ConditionalAccessStatus", ""), ""))
| extend __status = tolower(ResultType)
| extend __ip = extract(@"(?<!\d)(\d{1,3}(?:\.\d{1,3}){3})(?!\d)", 1, IPAddress)
| extend __isFailed = (__status in ("false","fail","failed","failure","denied","error","timeout","1") or (__status matches regex @"^\d+$" and toint(__status) != 0) or (tolower(ResultDescription) matches regex @"\b(fail|failed|failure|denied|error|timeout)\b"))
| extend __isSuccess = (__status in ("true","success","succeeded","completed","complete","ok","pass","passed","0"))
| extend __isPublicUntrustedIp = isnotempty(__ip) and not(ipv4_is_in_any_range(__ip, dynamic({TrustedIpCidrs})));
let __failed = __base | where __isFailed | summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by UserPrincipalName, ServicePrincipalName, ResourceDisplayName, IPAddress, ResultType, ResultDescription | where EventCount > 10;
let __suspicious = __base | where __isSuccess and __isPublicUntrustedIp | summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count(), ResultType=take_any(ResultType), ResultDescription=take_any(ResultDescription) by UserPrincipalName, ServicePrincipalName, ResourceDisplayName, IPAddress;
__failed | union isfuzzy=true __suspicious
```

### AssignedLicensesDCR_CL

```kusto
let __base =
AssignedLicensesDCR_CL
| where TimeGenerated >= datetime({StartUtc}) and TimeGenerated < datetime({EndUtc})
| extend UserPrincipalName = tostring(coalesce(column_ifexists("UserPrincipalName", ""), column_ifexists("UserUPN", ""), column_ifexists("UPN", ""), column_ifexists("Mail", ""), column_ifexists("EmailAddress", ""), column_ifexists("DisplayName", ""), column_ifexists("UserId", "")))
| extend SkuPartNumber = tostring(coalesce(column_ifexists("SkuPartNumber", ""), column_ifexists("LicenseName", ""), column_ifexists("SkuDisplayName", ""), column_ifexists("ServicePlanName", ""), column_ifexists("AssignedLicenses", ""), column_ifexists("Licenses", ""), "Unknown License"))
| extend TotalLicenses = todouble(tostring(coalesce(column_ifexists("TotalLicenses", ""), column_ifexists("TotalUnits", ""), column_ifexists("PrepaidUnitsEnabled", ""), column_ifexists("SkuPrepaidUnitsEnabled", ""), column_ifexists("EnabledUnits", ""), column_ifexists("Enabled", ""))));
__base
| summarize TimeGenerated=max(TimeGenerated), UsedUsers=dcount(UserPrincipalName), TotalLicenses=max(TotalLicenses) by SkuPartNumber
```

### AuditLogs

```kusto
AuditLogs
| where TimeGenerated >= datetime({StartUtc}) and TimeGenerated < datetime({EndUtc})
| extend __initiated = tostring(column_ifexists("InitiatedBy", ""))
| extend __actorUpn = tostring(coalesce(column_ifexists("InitiatedByUserPrincipalName", ""), column_ifexists("ActorUserPrincipalName", ""), column_ifexists("UserPrincipalName", ""), extract(@"""userPrincipalName""\s*:\s*""([^""]+)""", 1, __initiated), ""))
| extend __actorName = tostring(coalesce(column_ifexists("Actor", ""), column_ifexists("Identity", ""), extract(@"""displayName""\s*:\s*""([^""]+)""", 1, __initiated), ""))
| extend Actor = iff(isnotempty(__actorName) and isnotempty(__actorUpn) and not(__actorName contains __actorUpn), strcat(__actorName, " / ", __actorUpn), iff(isnotempty(__actorUpn), __actorUpn, __actorName))
| extend __actorIsUser = isnotempty(__actorUpn) or Actor contains "@" or __initiated contains @"""user"""
| where __actorIsUser and isnotempty(Actor)
| extend OperationName = tostring(coalesce(column_ifexists("OperationName", ""), column_ifexists("ActivityDisplayName", ""), column_ifexists("Activity", ""), column_ifexists("Operation", ""), "Audit Log Event"))
| extend Result = tostring(coalesce(column_ifexists("Result", ""), column_ifexists("ResultType", ""), column_ifexists("Status", ""), column_ifexists("ActivityStatus", ""), ""))
| extend __pimText = strcat(OperationName, " ", tostring(column_ifexists("ResultReason", "")), " ", tostring(column_ifexists("ResultDescription", "")), " ", tostring(column_ifexists("TargetResources", "")), " ", tostring(column_ifexists("ModifiedProperties", "")))
| where __pimText !matches regex @"(?i)\bPIM\b|PIM activation expired"
| extend __isSuccess = tolower(Result) in ("true","success","succeeded","completed","complete","ok","pass","passed","0")
| extend __isSpObjectChange = __isSuccess and OperationName in~ ("Add service principal", "Remove service principal", "Hard delete service principal")
| extend __isSpAppRoleChange = __isSuccess and OperationName in~ ("Add app role assignment to service principal", "Remove app role assignment from service principal")
| where __isSpObjectChange or __isSpAppRoleChange
| summarize TimeGenerated=max(TimeGenerated), EventCount=count() by Actor, OperationName, Target, PermissionName
```

### DCRLogErrors

```kusto
DCRLogErrors
| where TimeGenerated > ago(30d)
| distinct InputStreamId, OperationName, Message
```

### IntuneAuditLogsDCR_CL

```kusto
IntuneAuditLogsDCR_CL
| where TimeGenerated >= datetime({StartUtc}) and TimeGenerated < datetime({EndUtc})
| extend Actor = tostring(coalesce(column_ifexists("ActorUPN", ""), column_ifexists("ActorUserPrincipalName", ""), column_ifexists("Actor", ""), column_ifexists("UserPrincipalName", ""), "Unknown"))
| extend OperationName = tostring(coalesce(column_ifexists("OperationName", ""), column_ifexists("ActivityDisplayName", ""), column_ifexists("Activity", ""), column_ifexists("Operation", ""), column_ifexists("Action", ""), "Intune Audit Event"))
| extend TargetDisplayName = tostring(coalesce(column_ifexists("TargetDisplayName", ""), column_ifexists("Target", ""), column_ifexists("ObjectId", ""), column_ifexists("DeviceName", ""), ""))
| extend Result = tostring(coalesce(column_ifexists("Result", ""), column_ifexists("ResultStatus", ""), column_ifexists("Status", ""), column_ifexists("ActivityResult", ""), ""))
| extend ResultDescription = tostring(coalesce(column_ifexists("ResultDescription", ""), column_ifexists("FailureReason", ""), column_ifexists("Message", ""), column_ifexists("ErrorMessage", ""), ""))
| extend __status = tolower(Result)
| extend __isFailed = (__status in ("false","fail","failed","failure","denied","error","timeout","1") or (tolower(ResultDescription) matches regex @"\b(fail|failed|failure|denied|error|timeout)\b"))
| extend __isDeleteDisable = tolower(OperationName) matches regex @"(^|[^a-z])(delete|deleted|remove|removed|disable|disabled|deactivate|deactivated)([^a-z]|$)"
| extend __RecordKind = case(__isDeleteDisable, "AggregatedDeleteDisable", __isFailed, "AggregatedIntuneAuditRisk", "AggregatedIntuneAuditRecord")
| summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by Actor, OperationName, TargetDisplayName, Result, ResultDescription, __RecordKind
```

### MailboxStatisticsDCR_CL

```kusto
MailboxStatisticsDCR_CL
| where TimeGenerated >= datetime({StartUtc}) and TimeGenerated < datetime({EndUtc})
| extend DisplayName = tostring(coalesce(column_ifexists("DisplayName", ""), column_ifexists("MailboxDisplayName", ""), column_ifexists("Name", ""), ""))
| extend AvailableSpaceGB = todouble(extract(@"-?\d+(\.\d+)?", 0, tostring(coalesce(column_ifexists("AvailableSpaceGB", ""), column_ifexists("AvailableSpaceInGB", ""), column_ifexists("AvailableSpace", "")))))
| extend QuotaLimitGB = todouble(extract(@"-?\d+(\.\d+)?", 0, tostring(coalesce(column_ifexists("QuotaLimitGB", ""), column_ifexists("QuotaGB", ""), column_ifexists("StorageQuotaGB", ""), column_ifexists("ProhibitSendReceiveQuotaGB", "")))))
| extend IsSharedMailbox = tostring(coalesce(column_ifexists("IsSharedMailbox", ""), column_ifexists("IsSharedMailBox", ""), column_ifexists("IsShared", ""), column_ifexists("SharedMailbox", ""), column_ifexists("SharedMailBox", ""), ""))
| extend RecipientTypeDetails = tostring(coalesce(column_ifexists("RecipientTypeDetails", ""), column_ifexists("RecipientTypeDetail", ""), column_ifexists("RecipientTypeDetails_s", ""), column_ifexists("MailboxRecipientType", ""), column_ifexists("MailboxType", ""), column_ifexists("RecipientType", ""), ""))
| where (QuotaLimitGB > 0 and AvailableSpaceGB < QuotaLimitGB * 0.05) or RecipientTypeDetails contains "Shared" or IsSharedMailbox in~ ("true", "1", "yes", "y")
| extend UsagePercent = round((1 - AvailableSpaceGB / QuotaLimitGB) * 100, 2)
| project TimeGenerated, DisplayName, RecipientTypeDetails, IsSharedMailbox, AvailableSpaceGB, QuotaLimitGB, UsagePercent
```

### SigninLogs

```kusto
let __base =
SigninLogs
| where TimeGenerated >= datetime({StartUtc}) and TimeGenerated < datetime({EndUtc})
| extend UserPrincipalName = tostring(coalesce(column_ifexists("UserPrincipalName", ""), column_ifexists("UserDisplayName", ""), column_ifexists("Identity", ""), column_ifexists("UserId", ""), column_ifexists("User", ""), "Unknown"))
| extend AppDisplayName = tostring(coalesce(column_ifexists("AppDisplayName", ""), column_ifexists("Application", ""), column_ifexists("ApplicationDisplayName", ""), column_ifexists("ClientAppUsed", ""), "Unknown"))
| extend IPAddress = tostring(coalesce(column_ifexists("IPAddress", ""), column_ifexists("IpAddress", ""), column_ifexists("ClientIP", ""), column_ifexists("ClientIpAddress", ""), ""))
| extend ResultType = tostring(coalesce(column_ifexists("ResultType", ""), column_ifexists("Status", ""), column_ifexists("ResultDescription", ""), ""))
| extend ResultDescription = tostring(coalesce(column_ifexists("ResultDescription", ""), column_ifexists("FailureReason", ""), column_ifexists("Status", ""), ""))
| extend __status = tolower(ResultType)
| extend __ip = extract(@"(?<!\d)(\d{1,3}(?:\.\d{1,3}){3})(?!\d)", 1, IPAddress)
| extend __isFailed = (__status in ("false","fail","failed","failure","denied","error","timeout","1") or (__status matches regex @"^\d+$" and toint(__status) != 0) or (tolower(ResultDescription) matches regex @"\b(fail|failed|failure|denied|error|timeout)\b"))
| extend __isSuccess = (__status in ("true","success","succeeded","completed","complete","ok","pass","passed","0"))
| extend __isPublicUntrustedIp = isnotempty(__ip) and not(ipv4_is_in_any_range(__ip, dynamic({TrustedIpCidrs})))
| extend __isSigninSuspiciousSuccess = (__isSuccess and __isPublicUntrustedIp and not(AppDisplayName in~ ("Windows Sign In", "Microsoft Edge", "Sangfor SASE VPN", "Microsoft Office")));
let __failed = __base | where __isFailed | summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by UserPrincipalName, AppDisplayName, IPAddress, ResultType, ResultDescription;
let __suspicious = __base | where __isSigninSuspiciousSuccess | summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count(), ResultType=take_any(ResultType), ResultDescription=take_any(ResultDescription) by UserPrincipalName, AppDisplayName, IPAddress;
__failed | union isfuzzy=true __suspicious
```

## 维护规则

新增表时，需要同时更新 `scripts/log-analyzer-shared.ps1` 的 `$SupportedLogTables`、风险 KQL 生成函数、字段解析规则、报告聚合逻辑，以及本 `SKILL.md` 中的表说明和 KQL 模板。

修改查询范围、风险条件、字段合并或 HTML 展示规则后，需要运行 `scripts/tests` 下的测试，并优先用最近 3 小时范围验证输出。
