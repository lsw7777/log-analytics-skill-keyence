# Bulk Log Analyzer 架构与数据处理说明

本文档说明 `bulk-log-analyzer` 的程序架构、数据来源与去向、核心数据处理方式，以及目前采用的数据结构和缓存优化。文中的代码引用均指向仓库内当前实现，方便维护人员对照源码。

## 1. 程序定位

`bulk-log-analyzer` 是一个 PowerShell 日志分析工具，用于从 Azure Log Analytics 查询指定表的数据，导出 CSV，再生成自包含的 HTML 风险报告。报告重点关注失败、异常、可疑、删除、禁用、权限变动、许可证和容量风险。

当前支持 10 张表，定义在 `scripts/log-analyzer-core.ps1:1`：

```powershell
$SupportedLogTables = @(
    [PSCustomObject]@{ Name = 'AADManagedIdentitySignInLogs'; Description = 'AAD Managed Identity Sign-in Logs' },
    [PSCustomObject]@{ Name = 'AADServicePrincipalSignInLogs'; Description = 'AAD Service Principal Sign-in Logs' },
    [PSCustomObject]@{ Name = 'AssignedLicensesDCR_CL'; Description = 'Assigned Licenses' },
    [PSCustomObject]@{ Name = 'AuditGeneralDCR_CL'; Description = 'Audit General' },
    [PSCustomObject]@{ Name = 'AuditLogs'; Description = 'Microsoft Entra Audit Logs' },
    [PSCustomObject]@{ Name = 'AzureADUsersDCR_CL'; Description = 'Azure AD Users' },
    [PSCustomObject]@{ Name = 'MailboxStatisticsDCR_CL'; Description = 'Mailbox Statistics' },
    [PSCustomObject]@{ Name = 'MessageTraceDataDCR_CL'; Description = 'Message Trace Data' },
    [PSCustomObject]@{ Name = 'SharePointAuditDCR_CL'; Description = 'SharePoint Audit' },
    [PSCustomObject]@{ Name = 'SigninLogs'; Description = 'Microsoft Entra Sign-in Logs' }
)
```

## 2. 总体架构

程序由 3 层脚本组成：

| 层级 | 文件 | 职责 |
| --- | --- | --- |
| 编排层 | `scripts/run-all.ps1` | 接收用户参数、选择表和时间范围、检查缓存、调用查询脚本、调用报告生成脚本、可选打开浏览器。 |
| 查询层 | `scripts/azure_log_query.ps1` | 登录 Azure China Cloud，执行 KQL 查询，将 Log Analytics 结果导出为 CSV。 |
| 分析层 | `scripts/analyze.ps1` | 读取 CSV，按表画像归一化字段，统计指标，生成 HTML 报告。 |
| 公共核心 | `scripts/log-analyzer-core.ps1` | 支持表列表、时间范围、路径生成、表级字段画像、缓存校验、KQL 构造、调度配置等公共函数。 |

核心执行链路如下：

```text
用户运行 scripts/run-all.ps1
  -> 解析表名和时间范围
  -> 生成 CSV / HTML / Cache 路径
  -> 检查缓存是否可用
  -> 缓存命中：复制缓存 CSV 到工作 CSV
  -> 缓存未命中：调用 azure_log_query.ps1 查询 Log Analytics 并导出 CSV
  -> 调用 analyze.ps1 读取 CSV 并生成 HTML
  -> 可选用浏览器打开 HTML
```

对应的编排入口在 `run-all.ps1:220` 到 `run-all.ps1:331`：

```powershell
Write-Host "[1/4] Checking data source..." -ForegroundColor Yellow

if ($cacheResult -and $cacheResult.Hit) {
    Copy-Item -Path $CacheCsv -Destination $CsvFile -Force
} else {
    & "$ScriptDir\azure_log_query.ps1" @QueryParams
    if ($UseCache) {
        Save-Cache -SourceCsv $CsvFile -CacheCsv $CacheCsv -CacheMeta $CacheMeta -TableName $TableName -CacheTTL $CacheTTL -StartTime $StartTime -EndTime $EndTime
    }
}

& "$ScriptDir\analyze.ps1" -CsvPath $CsvFile -OutputPath $HtmlFilePath -AnalysisDate $AnalysisDateDisplay -TableName $TableName

if (-not $NoOpen) {
    Start-Process $HtmlFilePath
}
```

## 3. 数据来源

### 3.1 Azure Log Analytics

主要数据来源是 Azure Log Analytics 工作区。查询层使用 `Az.Accounts` 和 `Az.OperationalInsights` 模块，并固定面向 Azure China Cloud。配置在 `azure_log_query.ps1:32` 和 `azure_log_query.ps1:53`：

```powershell
[string]$WorkspaceId = "703a5771-97fc-4bf3-a585-f607d18c4479"
[string]$TenantId = "420c4dab-8603-402f-afe0-75bc28c51c13"
$AzureEnvironment = "AzureChinaCloud"
```

认证逻辑优先复用现有 Azure session。如果租户和环境匹配，则直接使用；否则打开浏览器或设备码登录。实现见 `azure_log_query.ps1:82`。

```powershell
$context = Get-AzContext -ErrorAction SilentlyContinue
if ($context -and $context.Account) {
    if ($currentTenant -and $currentTenant -eq $TenantId) {
        if ($context.Environment.Name -like "*China*") {
            return
        }
    }
}

Connect-AzAccount @connectParams | Out-Null
```

### 3.2 KQL 查询

程序不会直接拼接用户输入的任意查询来生成报告，而是基于受支持表名和时间范围生成 KQL。实现见 `log-analyzer-core.ps1:521`：

```powershell
function New-LogTableQuery {
    $start = $StartTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $end = $EndTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    return "$TableName | where TimeGenerated >= datetime($start) and TimeGenerated < datetime($end) | sort by TimeGenerated desc"
}
```

明确日期范围时，查询条件直接进入 KQL 的 `where TimeGenerated >= ... and < ...`，并且不会再额外传 `Timespan`。是否使用 `Timespan` 由 `Get-LogQueryExecutionMode` 决定，见 `log-analyzer-core.ps1:381`：

```powershell
if ($QueryStartTime -and $QueryEndTime) {
    return [PSCustomObject]@{
        UseTimespan = $false
        Timespan = $null
    }
}

return [PSCustomObject]@{
    UseTimespan = $true
    Timespan = New-TimeSpan -Hours $Hours
}
```

查询执行使用官方 Az cmdlet，见 `azure_log_query.ps1:177` 到 `azure_log_query.ps1:189`：

```powershell
$queryMode = Get-LogQueryExecutionMode -QueryStartTime $QueryStartTime -QueryEndTime $QueryEndTime -Hours $Hours

$queryParams = @{
    WorkspaceId = $WorkspaceId
    Query = $Query
    ErrorAction = 'Stop'
}
if ($queryMode.UseTimespan) {
    $queryParams['Timespan'] = $queryMode.Timespan
}

$response = Invoke-AzOperationalInsightsQuery @queryParams
```

## 4. 数据去向

### 4.1 工作 CSV

Log Analytics 查询结果会先导出到本机临时目录：

```text
C:\Users\<User>\AppData\Local\Temp\opencode\<TableName>_<Date>.csv
```

路径生成在 `log-analyzer-core.ps1:492`：

```powershell
return [PSCustomObject]@{
    CsvFile = Join-Path $TempDir "$($TableName)_$AnalysisDateStr.csv"
    HtmlFilePath = $htmlTarget
    CacheCsv = Join-Path (Join-Path $TempDir 'cache') "$($TableName)_$AnalysisDateStr.csv"
    CacheMeta = Join-Path (Join-Path $TempDir 'cache') "$($TableName)_$AnalysisDateStr.meta.json"
}
```

CSV 导出发生在 `azure_log_query.ps1:220`：

```powershell
if ($ExportCsv) {
    $response.Results | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
}
```

### 4.2 缓存 CSV 与元数据

缓存文件位于：

```text
C:\Users\<User>\AppData\Local\Temp\opencode\cache\<TableName>_<Date>.csv
C:\Users\<User>\AppData\Local\Temp\opencode\cache\<TableName>_<Date>.meta.json
```

缓存元数据保存表名、缓存时间、TTL、记录数、UTC 起止时间。实现见 `run-all.ps1:183`：

```powershell
$meta = @{
    TableName = $TableName
    CacheTime = (Get-Date).ToString("o")
    CacheTTL = $CacheTTL
    RecordCount = $recordCount
    StartTimeUtc = Get-LogCacheTimeKey -Time $StartTime
    EndTimeUtc = Get-LogCacheTimeKey -Time $EndTime
}
$meta | ConvertTo-Json | Out-File -FilePath $CacheMeta -Encoding UTF8 -Force
```

### 4.3 HTML 报告

HTML 报告生成到仓库根目录，命名格式如下：

```text
final_report_<TableName>_<Date>_<HHmm>.html
```

HTML 绝对路径和展示用相对路径都在 `Get-LogArtifactPaths` 中生成，见 `log-analyzer-core.ps1:506`：

```powershell
$repoRoot = Get-LogAnalyzerRepositoryRoot
$htmlTarget = Join-Path $repoRoot "final_report_$($TableName)_$($AnalysisDateStr)_$timestamp.html"
$htmlRelative = ConvertTo-RelativeLogPath -BasePath $repoRoot -TargetPath $htmlTarget
```

报告由 `analyze.ps1` 一次性生成自包含 HTML，不依赖外部前端服务。

## 5. 数据处理方式

### 5.1 表级字段画像

不同 Log Analytics 表字段不一致，例如用户字段可能叫 `UserUPN`、`UserId`、`UserPrincipalName`、`SenderAddress`；操作字段可能叫 `Activity`、`Operation`、`OperationType`、`Status`。程序通过表级画像统一抽象出：

- 用户字段：`UserFields`
- 操作字段：`OperationFields`
- 工作负载字段：`WorkloadFields`
- 客户端 IP 字段：`ClientIpFields`
- 成功/失败字段：`SuccessFields`
- 默认值和组合分组规则

公共画像定义在 `log-analyzer-core.ps1:155`：

```powershell
$common = [PSCustomObject]@{
    UserFields = @('UserUPN', 'UserId', 'UserPrincipalName', ...)
    OperationFields = @('Operation', 'Activity', 'Action', ...)
    WorkloadFields = @('Workload', 'Service', 'SourceSystem', ...)
    ClientIpFields = @('ClientIP', 'ClientIp', 'ClientIPAddress', ...)
    SuccessFields = @('IsSuccess', 'ResultStatus', 'Status', 'Result', ...)
    DefaultSuccess = 'unknown'
}
```

表级覆盖示例见 `log-analyzer-core.ps1:176`。例如 `AssignedLicensesDCR_CL` 使用 `ProvisioningStatus` 和 `ServicePlanName`，`SharePointAuditDCR_CL` 和 `AuditGeneralDCR_CL` 使用组合分组：

```powershell
'AssignedLicensesDCR_CL' {
    $common.OperationFields = @('ProvisioningStatus', 'ServicePlanName', ...)
    $common.SuccessFields = @('ProvisioningStatus') + $common.SuccessFields
    $common.DefaultSuccess = 'unknown'
}

'AuditGeneralDCR_CL' {
    $common.GroupFields = @('Activity', 'Operation', 'Workload')
    $common.UseCompositeOperationGroup = $true
}
```

### 5.2 字段取值和归一化

字段读取使用优先级列表：依次检查候选字段，返回第一个非空值，否则使用默认值。实现见 `analyze.ps1:74`：

```powershell
function Get-FieldValue {
    foreach ($name in $Names) {
        if ($Row.PSObject.Properties.Name -contains $name) {
            $value = [string]$Row.$name
            if ($value -and $value.Trim() -ne '') {
                return $value
            }
        }
    }

    return $Default
}
```

用户、操作、工作负载、IP、成功状态分别通过统一函数读取，见 `analyze.ps1:93` 到 `analyze.ps1:167`。

### 5.3 操作类型处理

普通表直接取画像中的操作字段。组合审计表按 `Activity + Operation + Workload` 生成组合操作名。实现见 `analyze.ps1:100`：

```powershell
if (-not $profile.UseCompositeOperationGroup) {
    return Get-FieldValue -Row $Row -Names $profile.OperationFields -Default $profile.DefaultOperation
}

$parts = @()
foreach ($field in $profile.GroupFields) {
    $value = Get-FieldValue -Row $Row -Names @($field) -Default ''
    if ($value) { $parts += $value }
}

return ($parts -join ' | ')
```

`AssignedLicensesDCR_CL` 是许可证快照表，操作类型专门构造为 `ProvisioningStatus | ServicePlanName`。实现见 `analyze.ps1:115`：

```powershell
if ($TableName -eq 'AssignedLicensesDCR_CL') {
    $status = Get-FieldValue -Row $Row -Names @('ProvisioningStatus') -Default ''
    $servicePlan = Get-FieldValue -Row $Row -Names @('ServicePlanName', 'SkuPartNumber', 'LicenseName') -Default ''
    if ($status -and $servicePlan) { return "$status | $servicePlan" }
    if ($status) { return $status }
    if ($servicePlan) { return $servicePlan }
}
```

### 5.4 成功/失败判定

通用表使用 `IsSuccess`、`ResultStatus`、`Status`、`Result` 等字段，将常见成功/失败词归一化为 `true`、`false`、`unknown`。实现见 `analyze.ps1:153`：

```powershell
if ($value -match '^(true|success|succeeded|delivered|expanded|completed|complete|ok|pass|passed|0)$') { return 'true' }
if ($value -match '^(false|fail|failed|failure|undelivered|blocked|rejected|denied|error|timeout|quarantined|1)$') { return 'false' }
return 'unknown'
```

`AssignedLicensesDCR_CL` 单独按 `ProvisioningStatus` 判断：

```powershell
if ($TableName -eq 'AssignedLicensesDCR_CL') {
    if ([string]::IsNullOrWhiteSpace($value) -or $value -eq 'unknown') { return 'unknown' }
    if ($value -eq 'success') { return 'true' }
    return 'false'
}
```

### 5.5 用户姓名映射

报告展示用户时优先显示姓名，再显示邮箱，例如：

```text
Shoyo Gao (C250105@china.keyence.com.cn)
```

姓名映射来自当前数据或同目录/缓存目录下最近的 `AzureADUsersDCR_CL*.csv`。实现见 `analyze.ps1:201` 到 `analyze.ps1:258`：

```powershell
$candidateFiles += @(Get-ChildItem -Path $csvDir -Filter 'AzureADUsersDCR_CL*.csv' -File ...)
$candidateFiles += @(Get-ChildItem -Path $cacheDir -Filter 'AzureADUsersDCR_CL*.csv' -File ...)

foreach ($file in $candidateFiles) {
    foreach ($row in @(Import-Csv -Path $file.FullName -Encoding UTF8)) {
        Add-UserDisplayNameMapping -Map $map -Row $row
    }
}
```

格式化逻辑见 `analyze.ps1:260`：

```powershell
return [regex]::Replace($User, $emailPattern, {
    $email = $match.Value
    $key = $email.ToLowerInvariant()
    if ($script:userDisplayNameMap.ContainsKey($key)) {
        $name = $script:userDisplayNameMap[$key]
        return "$name ($email)"
    }
    return $email
}, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
```

### 5.6 风险分析

风险分析目前统计以下维度：

- 失败操作
- 可疑 IP
- 非工作时间活动
- 高权限操作
- 敏感数据事件
- IP 多用户关联

风险指标数量由非空风险维度累加，见 `analyze.ps1:654`：

```powershell
$riskCount = 0
if ($failCount -gt 0) { $riskCount++ }
if ((Get-SafeCount $suspiciousIPs) -gt 0) { $riskCount++ }
if ((Get-SafeCount $offHoursEvents) -gt 0) { $riskCount++ }
if ((Get-SafeCount $highPrivEvents) -gt 0) { $riskCount++ }
if ((Get-SafeCount $sensitiveEvents) -gt 0) { $riskCount++ }
if ((Get-SafeCount $ipVelocity) -gt 0) { $riskCount++ }
```

对于快照表 `AssignedLicensesDCR_CL`，程序跳过 IP、非工作时间、高权限、敏感数据等审计语义扫描，只保留 `ProvisioningStatus` 失败风险。这避免把许可证快照误判为登录或访问行为，见 `analyze.ps1:317`、`analyze.ps1:358`、`analyze.ps1:388`。

## 6. HTML 生成与安全编码

报告中的图表数据由 PowerShell 统计后转为 JSON，再嵌入 HTML 中由浏览器端 JavaScript 渲染。为了避免 SharePoint 路径中单引号等字符打断 JS 字符串，程序不会直接写 `JSON.parse('...')`，而是先 URL encode，再在浏览器端 `decodeURIComponent`。

编码函数见 `analyze.ps1:489`：

```powershell
function ConvertTo-JsJsonLiteral {
    param([string]$Json)
    return [System.Uri]::EscapeDataString($Json)
}
```

图表数据生成见 `analyze.ps1:577`：

```powershell
$topUsersJson = ToSortedJsonArray $topUsers
$topUsersJsonJs = ConvertTo-JsJsonLiteral -Json $topUsersJson
$topOpsJson = ToSortedJsonArray $topOps
$topOpsJsonJs = ConvertTo-JsJsonLiteral -Json $topOpsJson
```

浏览器端解析见 `analyze.ps1:1275` 附近：

```javascript
renderBarChart('users-chart', JSON.parse(decodeURIComponent('$topUsersJsonJs')), 15);
renderBarChart('ops-chart', JSON.parse(decodeURIComponent('$topOpsJsonJs')), 15);
renderBarChart('ips-chart', JSON.parse(decodeURIComponent('$topIPsJsonJs')), 10, '$clientIpEmptyKey');
```

## 7. 数据结构优化

### 7.1 Hashtable 用于计数

统计排行类数据大量使用 PowerShell hashtable，将字符串键映射到计数值。这样可以在单次扫描中完成聚合，避免重复 `Group-Object` 带来的额外对象分配。

示例见 `analyze.ps1:293` 和 `analyze.ps1:307`：

```powershell
$workloadMap = @{}
foreach ($row in $data) {
    $wl = Get-WorkloadValue -Row $row
    $workloadMap[$wl] = ($workloadMap[$wl] + 1)
}

$opMap = @{}
foreach ($o in $allOps) {
    $opName = if ($o) { $o } else { 'Unknown' }
    $opMap[$opName] = ($opMap[$opName] + 1)
}
```

IP 多用户关联也使用嵌套 hashtable，外层 key 是 IP，内层 key 是用户。实现见 `analyze.ps1:416`：

```powershell
$ipUsers = @{}
foreach ($row in $data) {
    $ip = Get-ClientIpValue -Row $row
    if (-not (Test-UsableIpValue -IP $ip)) { continue }
    $u = Format-UserDisplayValue -User (Get-UserValue -Row $row)
    if (-not $ipUsers.ContainsKey($ip)) { $ipUsers[$ip] = @{} }
    $ipUsers[$ip][$u] = 1
}
```

### 7.2 Generic List 替代数组追加

PowerShell 数组 `+=` 会不断创建新数组，在大 CSV 上成本较高。热点路径中已改用 `.NET Generic List` 收集数据，再一次性转数组。

示例见 `analyze.ps1:280`：

```powershell
$allUsersList = [System.Collections.Generic.List[string]]::new()
foreach ($row in $data) {
    $u = Format-UserDisplayValue -User (Get-UserValue -Row $row)
    $allUsersList.Add($u) | Out-Null
}
$allUsers = $allUsersList.ToArray()
```

非工作时间事件同样使用 `Generic List`，见 `analyze.ps1:356`：

```powershell
$offHoursEventsList = [System.Collections.Generic.List[object]]::new()
...
$offHoursEventsList.Add($row) | Out-Null
...
$offHoursEvents = $offHoursEventsList.ToArray()
```

### 7.3 表画像缓存

`analyze.ps1` 在开始统计时只调用一次 `Get-TableAnalysisProfile`，并把结果保存在 `$analysisProfile`。后续字段读取函数复用该对象，避免每行每字段重复构造画像。

实现见 `analyze.ps1:31` 和 `analyze.ps1:93`：

```powershell
$analysisProfile = Get-TableAnalysisProfile -TableName $TableName

function Get-UserValue {
    $profile = $analysisProfile
    return Get-FieldValue -Row $Row -Names $profile.UserFields -Default $profile.DefaultUser
}
```

### 7.4 快照表跳过无意义扫描

`AssignedLicensesDCR_CL`、`AzureADUsersDCR_CL`、`MailboxStatisticsDCR_CL` 这类表本质是快照，不是用户行为日志。程序为许可证表跳过 IP、非工作时间、高权限、敏感数据扫描，避免误判并减少无效计算。

示例见 `analyze.ps1:317`、`analyze.ps1:358`、`analyze.ps1:388`：

```powershell
if ($TableName -ne 'AssignedLicensesDCR_CL') {
    foreach ($row in $data) {
        $ip = Get-ClientIpValue -Row $row
        if (-not (Test-UsableIpValue -IP $ip)) { continue }
        $ipMap[$ip] = ($ipMap[$ip] + 1)
    }
}

$highPrivEvents = if ($TableName -eq 'AssignedLicensesDCR_CL') { @() } else { @($data | Where-Object { Test-OperationContains -Row $_ -Operations $highPrivOps }) }
```

## 8. 缓存优化

### 8.1 缓存命中条件

缓存命中必须同时满足：

- 缓存 CSV 存在
- 缓存 meta JSON 存在
- meta 中的 `TableName` 与当前表一致
- meta 中的 `StartTimeUtc` 与当前查询 UTC 起始时间一致
- meta 中的 `EndTimeUtc` 与当前查询 UTC 结束时间一致
- 缓存未过 TTL
- 缓存 CSV 有效且记录数大于 0

缓存检查入口在 `run-all.ps1:133`：

```powershell
function Test-Cache {
    if (-not (Test-Path $CacheCsv)) { return @{ Hit = $false; Reason = "Cache file not found" } }
    if (-not (Test-Path $CacheMeta)) { return @{ Hit = $false; Reason = "Cache metadata not found" } }

    $meta = Get-Content $CacheMeta -Raw | ConvertFrom-Json
    if (-not (Test-LogCacheMetadataMatches -Meta $meta -TableName $TableName -StartTime $StartTime -EndTime $EndTime)) {
        return @{ Hit = $false; Reason = 'Cache metadata does not match table and time range' }
    }
}
```

严格匹配逻辑在 `log-analyzer-core.ps1:323`：

```powershell
if ($Meta.TableName -ne $TableName) { return $false }
if ($Meta.StartTimeUtc -ne (Get-LogCacheTimeKey -Time $StartTime)) { return $false }
if ($Meta.EndTimeUtc -ne (Get-LogCacheTimeKey -Time $EndTime)) { return $false }
```

### 8.2 空缓存保护

查询结果为 0 行时不会保存缓存，避免后续误用空 CSV。实现见 `run-all.ps1:193`：

```powershell
$recordCount = Get-LogCsvRecordCount -CsvPath $SourceCsv
if ($recordCount -le 0) {
    Remove-Item -Path $CacheCsv -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $CacheMeta -Force -ErrorAction SilentlyContinue
    Write-Host "Cache skipped: query returned 0 records" -ForegroundColor Yellow
    return
}
```

缓存 payload 校验见 `log-analyzer-core.ps1:352`：

```powershell
if (-not (Test-Path $CacheCsv)) { return $false }
if ((Get-Item -Path $CacheCsv).Length -eq 0) { return $false }
if ($null -eq $RecordCount) { return $false }
if ([int]$RecordCount -le 0) { return $false }
return $true
```

如果命中检查时发现 payload 无效，程序会删除对应 CSV 和 meta，并重新查询，见 `run-all.ps1:157`：

```powershell
if (-not (Test-LogCachePayloadValid -CacheCsv $CacheCsv -RecordCount $meta.RecordCount)) {
    Remove-Item -Path $CacheCsv -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $CacheMeta -Force -ErrorAction SilentlyContinue
    return @{ Hit = $false; Reason = 'Cache payload is empty or invalid' }
}
```

### 8.3 LRU 缓存数量限制

缓存目录使用 LRU 策略限制缓存组合数量，避免长期运行后缓存无限增长。调度和批量运行时默认最多保留 20 个缓存组合。相关函数在 `log-analyzer-core.ps1` 中实现，并由 `run-all.ps1` 在保存缓存后调用。

用户也可以主动清理缓存。`run-all.ps1` 支持 `-ClearCache`，并兼容拼写 `-ClearCashe`，见 `run-all.ps1:40` 和 `run-all.ps1:55`：

```powershell
[switch]$ClearCache,
[switch]$ClearCashe,

if ($ClearCache -or $ClearCashe) {
    $removedCount = Clear-LogCache -CacheDir $CacheDir
    Write-Host "Cache cleared: $CacheDir ($removedCount items removed)" -ForegroundColor Green
    exit 0
}
```

## 9. 输出报告内容

HTML 报告主要包含：

- 总事件数、唯一用户数、唯一操作数、工作负载数
- 活动时间线或快照写入时间线
- 工作负载分布
- 活跃用户排行
- 操作类型排行
- 客户端 IP 排行或表级缺失说明
- 成功/失败/未知比率
- 风险分析
- 术语表
- 明细数据预览

明细表预览前 1000 行，避免 HTML 过大。实现见 `analyze.ps1:640`：

```powershell
$previewRows = [Math]::Min(1000, $data.Count)
for ($i = 0; $i -lt $previewRows; $i++) {
    $row = $data[$i]
    $user = Format-UserDisplayValue -User (Get-UserValue -Row $row)
    $op = Get-OperationValue -Row $row
    $success = Get-SuccessValue -Row $row
    $tableRows += "<tr>...</tr>`n"
}
```

## 10. 定时执行

程序支持 Windows 任务计划程序和托盘状态程序。调度配置由公共函数生成，默认时间是 `01:00`，默认表集合是全部支持表。实现见 `log-analyzer-core.ps1:405`：

```powershell
function New-LogAnalyzerScheduleConfig {
    param(
        [string]$RunAt = '01:00',
        [string[]]$Tables = @()
    )

    if (-not $Tables -or $Tables.Count -eq 0) {
        $Tables = @($SupportedLogTables | ForEach-Object { $_.Name })
    }
}
```

批量执行命令由 `Get-LogAnalyzerBatchCommand` 生成，见 `log-analyzer-core.ps1:439`：

```powershell
$scriptPath = Join-Path $RootDir 'scheduled-run.ps1'
return "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ConfigPath `"$ConfigPath`""
```

定时执行时通常传 `-NoOpen`，避免无人值守场景自动打开浏览器。

## 11. 维护建议

新增表时，优先修改 `Get-TableAnalysisProfile`，为新表声明字段优先级和默认值，而不是在统计逻辑中散落特殊判断。

如果新表是快照表，应补充时间线说明、IP 缺失说明，并避免运行登录/访问类风险扫描。

如果报告中出现图表空白，优先检查 HTML 中嵌入 JSON 的安全编码，尤其是数据中是否存在单引号、反斜杠、换行或 URL 字符。

如果查询结果明显不匹配时间范围，优先检查 `New-LogTableQuery` 和 `Get-LogQueryExecutionMode`，确保明确日期范围通过 KQL 过滤，不额外叠加 `Timespan`。

如果性能下降，优先检查热点循环中是否重新引入了数组 `+=`、重复 `Get-TableAnalysisProfile`、重复 `Import-Csv`、或对快照表执行了无意义风险扫描。
