param(
    [Parameter(Mandatory = $true)]
    [string[]]$CsvPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [string]$AnalysisDate,

    [Parameter(Mandatory = $true)]
    [string[]]$TableName
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $ScriptDir 'log-analyzer-core.ps1')

if ($CsvPath.Count -ne $TableName.Count) {
    throw 'CsvPath and TableName must have the same number of items.'
}

function Escape-Html {
    param([object]$Text)
    if ($null -eq $Text) { return '' }
    $s = [string]$Text
    return $s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&#39;'
}

function Get-LocalTimeText {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    try { return ([DateTime]::Parse($Value)).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss') } catch { return $Value }
}

function Get-NumberValue {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if (-not $text) { return $null }
    $match = [regex]::Match($text, '-?\d+(\.\d+)?')
    if (-not $match.Success) { return $null }
    return [double]$match.Value
}

function Get-AnyFieldValue {
    param(
        [object]$Row,
        [string[]]$Names,
        [string]$Default = ''
    )

    foreach ($name in $Names) {
        if ($Row.PSObject.Properties.Name -contains $name) {
            $value = [string]$Row.$name
            if (-not [string]::IsNullOrWhiteSpace($value)) { return $value.Trim() }
        }
    }
    return $Default
}

function Add-Count {
    param(
        [hashtable]$Map,
        [string]$Key,
        [int]$By = 1
    )
    if ([string]::IsNullOrWhiteSpace($Key)) { $Key = 'Unknown' }
    $Map[$Key] = ($Map[$Key] + $By)
}

function New-EventRecord {
    param(
        [string]$Table,
        [object]$Row,
        [string]$Reason
    )

    $op = Get-OperationValue -Row $Row -TableName $Table
    $user = Format-UserForReport -User (Get-UserValue -Row $Row -TableName $Table)
    $ip = Get-NormalizedIpValue -IP (Get-ClientIpValue -Row $Row -TableName $Table)
    $status = Get-SuccessValue -Row $Row -TableName $Table
    $detail = Get-AnyFieldValue -Row $Row -Names @('ResultDescription', 'ResultReason', 'FailureReason', 'Status', 'DeliveryStatus', 'ErrorCode', 'ResultType', 'Subject', 'TargetResources', 'ModifiedProperties') -Default ''
    if ($Table -eq 'MessageTraceDataDCR_CL') {
        $traceStatus = Get-AnyFieldValue -Row $Row -Names @('Status', 'DeliveryStatus', 'EventType', 'Action', 'Result') -Default ''
        $subject = Get-AnyFieldValue -Row $Row -Names @('Subject') -Default ''
        $detail = (@($traceStatus, $subject) | Where-Object { $_ }) -join ' | '
    }
    $target = Get-AnyFieldValue -Row $Row -Names @('TargetResources', 'TargetResource', 'ObjectId', 'ResourceDisplayName', 'AppDisplayName', 'ServicePrincipalName', 'DisplayName', 'RecipientAddress', 'Subject') -Default ''

    return [PSCustomObject]@{
        Table = $Table
        Time = Get-LocalTimeText -Value ([string]$Row.TimeGenerated)
        User = $user
        Operation = $op
        IP = $ip
        Status = $status
        Target = $target
        Detail = $detail
        Reason = $Reason
    }
}

function New-TableHtml {
    param(
        [object[]]$Rows,
        [string[]]$Columns,
        [scriptblock]$CellBuilder
    )

    if (-not $Rows -or @($Rows).Count -eq 0) {
        return '<p class="empty">未发现相关风险。</p>'
    }

    $html = '<div class="table-scroll"><table><thead><tr>'
    foreach ($col in $Columns) { $html += '<th>' + (Escape-Html $col) + '</th>' }
    $html += '</tr></thead><tbody>'
    foreach ($row in $Rows) {
        $html += '<tr>'
        $cells = & $CellBuilder $row
        foreach ($cell in $cells) { $html += '<td>' + (Escape-Html $cell) + '</td>' }
        $html += '</tr>'
    }
    $html += '</tbody></table></div>'
    return $html
}

function Format-UserForReport {
    param([string]$User)

    if ([string]::IsNullOrWhiteSpace($User)) { return 'Unknown' }
    $emailPattern = '[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}'
    return [regex]::Replace($User, $emailPattern, {
        param($match)
        $email = $match.Value
        $key = $email.ToLowerInvariant()
        if ($script:DisplayNameMap.ContainsKey($key)) {
            return "$($script:DisplayNameMap[$key]) ($email)"
        }
        return $email
    }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Add-DisplayNameFromRow {
    param([object]$Row)

    $displayName = Get-AnyFieldValue -Row $Row -Names @('displayName', 'DisplayName', 'UserDisplayName') -Default ''
    if (-not $displayName) { return }
    $ids = @(
        (Get-AnyFieldValue -Row $Row -Names @('userPrincipalName', 'UserPrincipalName', 'UserUPN', 'UPN') -Default ''),
        (Get-AnyFieldValue -Row $Row -Names @('mail', 'Mail', 'EmailAddress') -Default '')
    )
    foreach ($id in $ids) {
        if ($id -and $id -match '@') {
            $script:DisplayNameMap[$id.ToLowerInvariant()] = $displayName
        }
    }
}

function Add-DisabledUserFromRow {
    param([object]$Row)

    $enabled = (Get-AnyFieldValue -Row $Row -Names @('accountEnabled', 'AccountEnabled') -Default '').ToLowerInvariant()
    if ($enabled -ne 'false') { return }
    $upn = Get-AnyFieldValue -Row $Row -Names @('userPrincipalName', 'UserPrincipalName', 'UserUPN', 'UPN', 'mail', 'Mail') -Default ''
    if (-not $upn) { return }
    $time = Get-AnyFieldValue -Row $Row -Names @('disabledDateTime', 'DisabledDateTime', 'accountDisabledDateTime', 'AccountDisabledDateTime', 'TimeGenerated') -Default ''
    $script:DisabledUserMap[$upn.ToLowerInvariant()] = [PSCustomObject]@{
        User = Format-UserForReport -User $upn
        DisabledTime = Get-LocalTimeText -Value $time
    }
}

function Test-InterestingMessageTrace {
    param([object]$Row)
    $text = @(
        (Get-AnyFieldValue -Row $Row -Names @('Status', 'DeliveryStatus', 'EventType', 'Action', 'Result') -Default ''),
        (Get-AnyFieldValue -Row $Row -Names @('Subject', 'SenderAddress', 'RecipientAddress') -Default '')
    ) -join ' '
    return ($text -match '(?i)fail|failed|failure|blocked|quarantined|rejected|undeliver|error|timeout|Power\s*BI|PBI|skyguard')
}

function Get-SigninAppName {
    param([object]$Row)
    return Get-AnyFieldValue -Row $Row -Names @('AppDisplayName', 'Application', 'ApplicationDisplayName', 'ClientAppUsed') -Default ''
}

function Test-AllowedSigninApp {
    param([string]$AppName)
    if ([string]::IsNullOrWhiteSpace($AppName)) { return $false }
    return @('Windows Sign In', 'Microsoft Edge', 'Sangfor SASE VPN', 'Microsoft Office') -contains $AppName.Trim()
}

Write-Host 'Loading CSV data...' -ForegroundColor Cyan
$datasets = [System.Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $CsvPath.Count; $i++) {
    $path = $CsvPath[$i]
    $table = $TableName[$i]
    if (-not (Test-Path $path)) {
        Write-Host "Skipping missing CSV for $table`: $path" -ForegroundColor Yellow
        continue
    }
    $rows = @(Import-Csv -Path $path -Encoding UTF8)
    $datasets.Add([PSCustomObject]@{ Table = $table; Path = $path; Rows = $rows }) | Out-Null
    Write-Host "  $table`: $($rows.Count) records" -ForegroundColor Green
}

$script:DisplayNameMap = @{}
$script:DisabledUserMap = @{}
foreach ($dataset in $datasets) {
    if ($dataset.Table -eq 'AzureADUsersDCR_CL') {
        foreach ($row in $dataset.Rows) {
            Add-DisplayNameFromRow -Row $row
        }
    }
}
foreach ($dataset in $datasets) {
    if ($dataset.Table -eq 'AzureADUsersDCR_CL') {
        foreach ($row in $dataset.Rows) {
            Add-DisabledUserFromRow -Row $row
        }
    }
}

$trustedRules = @(Get-TrustedIpRules)
$failedSignins = [System.Collections.Generic.List[object]]::new()
$failedOperations = [System.Collections.Generic.List[object]]::new()
$deleteDisableEvents = [System.Collections.Generic.List[object]]::new()
$offHoursEvents = [System.Collections.Generic.List[object]]::new()
$suspiciousSigninSuccess = [System.Collections.Generic.List[object]]::new()
$suspiciousIpReasons = @{}
$clientIpCounts = @{}
$messageTraceRisks = [System.Collections.Generic.List[object]]::new()
$identityPermissionChanges = [System.Collections.Generic.List[object]]::new()
$azureAdDisabledUsers = [System.Collections.Generic.List[object]]::new()
$azureAdMissingFields = [System.Collections.Generic.List[object]]::new()
$sourceStatusRows = [System.Collections.Generic.List[object]]::new()

foreach ($dataset in $datasets) {
    $table = $dataset.Table
    $sourceStatusRows.Add([PSCustomObject]@{ Table = $table; Records = $dataset.Rows.Count; Source = (Split-Path -Leaf $dataset.Path) }) | Out-Null

    foreach ($row in $dataset.Rows) {
        $op = Get-OperationValue -Row $row -TableName $table
        $success = Get-SuccessValue -Row $row -TableName $table
        $ip = Get-NormalizedIpValue -IP (Get-ClientIpValue -Row $row -TableName $table)
        $isUsablePublicIp = -not (Test-PrivateOrInvalidIp -IP $ip)
        $isTrustedIp = Test-IpInTrustedRules -IP $ip -Rules $trustedRules

        if ($isUsablePublicIp) {
            Add-Count -Map $clientIpCounts -Key $ip
        }

        if ($success -eq 'false') {
            $record = New-EventRecord -Table $table -Row $row -Reason '失败/异常'
            if ($table -in @('AADManagedIdentitySignInLogs', 'AADServicePrincipalSignInLogs', 'SigninLogs')) {
                $failedSignins.Add($record) | Out-Null
            } else {
                $failedOperations.Add($record) | Out-Null
            }
        }

        if (Test-DeleteOrDisableOperation -Operation $op -TableName $table) {
            $deleteDisableEvents.Add((New-EventRecord -Table $table -Row $row -Reason '删除 / Disable 操作')) | Out-Null
        }

        if ($table -notin @('AssignedLicensesDCR_CL', 'AzureADUsersDCR_CL', 'MailboxStatisticsDCR_CL') -and (Test-LogOffHours -TimeGenerated ([string]$row.TimeGenerated))) {
            $offHoursEvents.Add((New-EventRecord -Table $table -Row $row -Reason '非工作时间 21:00 - 08:00')) | Out-Null
        }

        if ($table -eq 'SigninLogs' -and $success -eq 'true' -and $isUsablePublicIp -and -not $isTrustedIp) {
            $appName = Get-SigninAppName -Row $row
            if (-not (Test-AllowedSigninApp -AppName $appName)) {
                $event = New-EventRecord -Table $table -Row $row -Reason "可信位置外成功登录，应用：$appName"
                $suspiciousSigninSuccess.Add($event) | Out-Null
                $suspiciousIpReasons[$ip] = 'SigninLogs 可信位置外成功登录'
            }
        }

        if ($table -ne 'AssignedLicensesDCR_CL' -and $isUsablePublicIp -and -not $isTrustedIp) {
            $workload = Get-WorkloadValue -Row $row -TableName $table
            if (-not $suspiciousIpReasons.ContainsKey($ip)) {
                $suspiciousIpReasons[$ip] = "公共 IP，来源表/工作负载：$table / $workload"
            }
        }

        if ($table -eq 'MessageTraceDataDCR_CL' -and (Test-InterestingMessageTrace -Row $row)) {
            $messageTraceRisks.Add((New-EventRecord -Table $table -Row $row -Reason '邮件投递异常 / PBI / SkyGuard 相关')) | Out-Null
        }

        $permissionText = "$op " + (Get-AnyFieldValue -Row $row -Names @('TargetResources', 'ModifiedProperties', 'ResultDescription', 'ActivityDisplayName') -Default '')
        if ($table -in @('AuditLogs', 'AADManagedIdentitySignInLogs', 'AADServicePrincipalSignInLogs') -and $permissionText -match '(?i)permission|consent|credential|secret|certificate|app role|approle|service principal|managed identity') {
            $identityPermissionChanges.Add((New-EventRecord -Table $table -Row $row -Reason '身份 / 应用权限变动')) | Out-Null
        }

        if ($table -eq 'AzureADUsersDCR_CL') {
            $enabled = (Get-AnyFieldValue -Row $row -Names @('accountEnabled', 'AccountEnabled') -Default '').ToLowerInvariant()
            if ($enabled -eq 'false') {
                $azureAdDisabledUsers.Add((New-EventRecord -Table $table -Row $row -Reason 'AAD 用户已禁用')) | Out-Null
            }
            $upn = Get-AnyFieldValue -Row $row -Names @('userPrincipalName', 'UserPrincipalName', 'UserUPN', 'UPN') -Default ''
            $mail = Get-AnyFieldValue -Row $row -Names @('mail', 'Mail') -Default ''
            $displayName = Get-AnyFieldValue -Row $row -Names @('displayName', 'DisplayName') -Default ''
            if (-not $upn -or -not $mail -or -not $displayName) {
                $userForMissing = if ($upn) { $upn } elseif ($mail) { $mail } else { $displayName }
                $azureAdMissingFields.Add([PSCustomObject]@{
                    User = Format-UserForReport -User $userForMissing
                    Missing = (@(
                        if (-not $upn) { 'UPN' }
                        if (-not $mail) { 'Mail' }
                        if (-not $displayName) { 'DisplayName' }
                    ) -join ', ')
                }) | Out-Null
            }
        }
    }
}

# License usage: infer the four license names from the data and show remaining count if total fields exist.
$licenseRows = @($datasets | Where-Object { $_.Table -eq 'AssignedLicensesDCR_CL' } | ForEach-Object { $_.Rows })
$licenseGroups = @{}
foreach ($row in $licenseRows) {
    $licenseName = Get-AnyFieldValue -Row $row -Names @('SkuPartNumber', 'LicenseName', 'SkuDisplayName', 'ServicePlanName', 'AssignedLicenses', 'Licenses') -Default 'Unknown License'
    if (-not $licenseGroups.ContainsKey($licenseName)) {
        $licenseGroups[$licenseName] = [PSCustomObject]@{
            Name = $licenseName
            Users = @{}
            Rows = 0
            Total = $null
            UsedOverride = $null
        }
    }
    $group = $licenseGroups[$licenseName]
    $group.Rows++
    $usedUsers = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('UsedUsers', 'Used') -Default '')
    if ($null -ne $usedUsers) {
        if ($null -eq $group.UsedOverride -or $usedUsers -gt $group.UsedOverride) {
            $group.UsedOverride = [int]$usedUsers
        }
    }
    $user = Get-UserValue -Row $row -TableName 'AssignedLicensesDCR_CL'
    if ($user) { $group.Users[$user.ToLowerInvariant()] = 1 }
    $total = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('TotalLicenses', 'TotalUnits', 'PrepaidUnitsEnabled', 'SkuPrepaidUnitsEnabled', 'EnabledUnits', 'Enabled') -Default '')
    if ($null -ne $total) {
        if ($null -eq $group.Total -or $total -gt $group.Total) { $group.Total = $total }
    }
}
$licenseUsage = @(
    $licenseGroups.Values |
        Sort-Object @{ Expression = { if ($null -ne $_.UsedOverride) { $_.UsedOverride } else { $_.Users.Count } }; Descending = $true }, @{ Expression = { $_.Rows }; Descending = $true } |
        Select-Object -First 4 |
        ForEach-Object {
            $used = if ($null -ne $_.UsedOverride) { $_.UsedOverride } elseif ($_.Users.Count -gt 0) { $_.Users.Count } else { $_.Rows }
            $remaining = if ($null -ne $_.Total) { [Math]::Max(0, [int]$_.Total - [int]$used) } else { 'N/A' }
            [PSCustomObject]@{ License = $_.Name; Used = $used; Total = $(if ($null -ne $_.Total) { [int]$_.Total } else { 'N/A' }); Remaining = $remaining }
        }
)

$mailboxRisks = [System.Collections.Generic.List[object]]::new()
$sharedMailboxRows = [System.Collections.Generic.List[object]]::new()
$mailboxRows = @($datasets | Where-Object { $_.Table -eq 'MailboxStatisticsDCR_CL' } | ForEach-Object { $_.Rows })
foreach ($row in $mailboxRows) {
    $available = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('AvailableSpaceGB', 'AvailableSpaceInGB', 'AvailableSpace') -Default '')
    $quota = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('QuotaLimitGB', 'QuotaGB', 'StorageQuotaGB', 'ProhibitSendReceiveQuotaGB') -Default '')
    $user = Get-UserValue -Row $row -TableName 'MailboxStatisticsDCR_CL'
    $size = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('TotalItemSizeGB', 'MailboxSizeGB', 'SizeGB', 'TotalSizeGB') -Default '')
    $type = Get-AnyFieldValue -Row $row -Names @('RecipientTypeDetails', 'MailboxType', 'RecipientType') -Default ''

    if ($null -ne $available -and $null -ne $quota -and $quota -gt 0 -and ($available / $quota) -lt 0.05) {
        $mailboxRisks.Add([PSCustomObject]@{
            User = Format-UserForReport -User $user
            AvailableGB = [Math]::Round($available, 2)
            QuotaGB = [Math]::Round($quota, 2)
            Usage = '{0:P1}' -f (1 - ($available / $quota))
            Reason = 'AvailableSpaceGB 低于 QuotaLimitGB 的 5%，可能导致无法收发邮件'
        }) | Out-Null
    }

    if ($type -match '(?i)shared') {
        $key = $user.ToLowerInvariant()
        $disabled = if ($script:DisabledUserMap.ContainsKey($key)) { $script:DisabledUserMap[$key] } else { $null }
        $sharedMailboxRows.Add([PSCustomObject]@{
            User = Format-UserForReport -User $user
            SizeGB = if ($null -ne $size) { [Math]::Round($size, 2) } else { 'N/A' }
            Disabled = if ($disabled) { 'Yes' } else { 'No' }
            DisabledTime = if ($disabled) { $disabled.DisabledTime } else { '' }
        }) | Out-Null
    }
}

$suspiciousIpRows = @(
    $suspiciousIpReasons.GetEnumerator() |
        Where-Object { -not (Test-IpInTrustedRules -IP $_.Key -Rules $trustedRules) } |
        Sort-Object Name |
        ForEach-Object { [PSCustomObject]@{ IP = $_.Key; Reason = $_.Value } }
)
$suspiciousIpSet = @{}
foreach ($row in $suspiciousIpRows) { $suspiciousIpSet[$row.IP] = 1 }
$topClientIps = @(
    $clientIpCounts.GetEnumerator() |
        Where-Object { -not $suspiciousIpSet.ContainsKey($_.Key) } |
        Sort-Object Value -Descending |
        Select-Object -First 20 |
        ForEach-Object { [PSCustomObject]@{ IP = $_.Key; Count = $_.Value } }
)

$riskCounts = [PSCustomObject]@{
    FailedSignins = $failedSignins.Count
    FailedOperations = $failedOperations.Count
    DeleteDisable = $deleteDisableEvents.Count
    SuspiciousIPs = $suspiciousIpRows.Count
    SuspiciousSigninSuccess = $suspiciousSigninSuccess.Count
    OffHours = $offHoursEvents.Count
    LicenseTypes = $licenseUsage.Count
    MailboxLowSpace = $mailboxRisks.Count
    SharedMailboxes = $sharedMailboxRows.Count
    MessageTrace = $messageTraceRisks.Count
    IdentityPermissionChanges = $identityPermissionChanges.Count
    DisabledUsers = $azureAdDisabledUsers.Count
}

$totalRecords = ($datasets | ForEach-Object { $_.Rows.Count } | Measure-Object -Sum).Sum
$tableCount = $datasets.Count
$trustedCount = $trustedRules.Count

$failedSigninHtml = New-TableHtml -Rows ($failedSignins | Select-Object -First 50) -Columns @('表', '时间', '主体/用户', '应用/操作', 'IP', '失败原因') -CellBuilder {
    param($r) @($r.Table, $r.Time, $r.User, $r.Operation, $r.IP, $r.Detail)
}
$failedOpsHtml = New-TableHtml -Rows ($failedOperations | Select-Object -First 50) -Columns @('表', '时间', '用户', '操作', '状态/原因') -CellBuilder {
    param($r) @($r.Table, $r.Time, $r.User, $r.Operation, $r.Detail)
}
$deleteDisableHtml = New-TableHtml -Rows ($deleteDisableEvents | Select-Object -First 80) -Columns @('表', '时间', '操作者', '操作', '目标', '结果') -CellBuilder {
    param($r) @($r.Table, $r.Time, $r.User, $r.Operation, $r.Target, $r.Detail)
}
$suspiciousIpHtml = New-TableHtml -Rows $suspiciousIpRows -Columns @('IP', '原因') -CellBuilder {
    param($r) @($r.IP, $r.Reason)
}
$signinSuspiciousHtml = New-TableHtml -Rows ($suspiciousSigninSuccess | Select-Object -First 80) -Columns @('时间', '用户', '应用', 'IP', '说明') -CellBuilder {
    param($r) @($r.Time, $r.User, $r.Operation, $r.IP, $r.Reason)
}
$topClientIpHtml = New-TableHtml -Rows $topClientIps -Columns @('IP', '次数') -CellBuilder {
    param($r) @($r.IP, $r.Count)
}
$offHoursHtml = New-TableHtml -Rows ($offHoursEvents | Select-Object -First 80) -Columns @('表', '时间', '用户/主体', '操作', 'IP') -CellBuilder {
    param($r) @($r.Table, $r.Time, $r.User, $r.Operation, $r.IP)
}
$licenseHtml = New-TableHtml -Rows $licenseUsage -Columns @('License 名称', '已使用', '总数', '剩余') -CellBuilder {
    param($r) @($r.License, $r.Used, $r.Total, $r.Remaining)
}
$mailboxRiskHtml = New-TableHtml -Rows ($mailboxRisks | Sort-Object AvailableGB | Select-Object -First 50) -Columns @('邮箱', 'AvailableSpaceGB', 'QuotaLimitGB', '使用率', '风险') -CellBuilder {
    param($r) @($r.User, $r.AvailableGB, $r.QuotaGB, $r.Usage, $r.Reason)
}
$sharedMailboxHtml = New-TableHtml -Rows ($sharedMailboxRows | Sort-Object SizeGB -Descending | Select-Object -First 80) -Columns @('SharedMailbox', '大小GB', '对应用户Disabled', 'Disable时间') -CellBuilder {
    param($r) @($r.User, $r.SizeGB, $r.Disabled, $r.DisabledTime)
}
$messageTraceHtml = New-TableHtml -Rows ($messageTraceRisks | Select-Object -First 80) -Columns @('时间', '用户/地址', '状态/主题', '目标', '说明') -CellBuilder {
    param($r) @($r.Time, $r.User, $r.Operation, $r.Target, $r.Detail)
}
$permissionHtml = New-TableHtml -Rows ($identityPermissionChanges | Select-Object -First 80) -Columns @('表', '时间', '操作者/主体', '操作', '目标', '说明') -CellBuilder {
    param($r) @($r.Table, $r.Time, $r.User, $r.Operation, $r.Target, $r.Detail)
}
$disabledUsersHtml = New-TableHtml -Rows ($azureAdDisabledUsers | Select-Object -First 80) -Columns @('时间', '用户', '状态/分类') -CellBuilder {
    param($r) @($r.Time, $r.User, $r.Operation)
}
$missingUsersHtml = New-TableHtml -Rows ($azureAdMissingFields | Select-Object -First 80) -Columns @('用户', '缺少字段') -CellBuilder {
    param($r) @($r.User, $r.Missing)
}
$sourceStatusHtml = New-TableHtml -Rows $sourceStatusRows -Columns @('表', '记录数', 'CSV') -CellBuilder {
    param($r) @($r.Table, $r.Records, $r.Source)
}

$html = @"
<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Log Analytics 合并风险报告</title>
<style>
:root {
  --bg: #0f1720;
  --panel: #151f2b;
  --panel2: #1d2a38;
  --text: #eef4fb;
  --muted: #9fb0c2;
  --line: #2d3d4f;
  --red: #ff6b6b;
  --amber: #f3b95f;
  --green: #66d98f;
  --blue: #6bb6ff;
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--text); font-family: "Segoe UI", Arial, sans-serif; }
.wrap { max-width: 1440px; margin: 0 auto; padding: 28px; }
.header { border-bottom: 1px solid var(--line); padding-bottom: 18px; margin-bottom: 22px; }
h1 { margin: 0 0 10px; font-size: 28px; font-weight: 700; }
h2 { margin: 0 0 14px; font-size: 19px; }
.meta { display: flex; gap: 10px; flex-wrap: wrap; color: var(--muted); }
.tag { background: var(--panel2); border: 1px solid var(--line); border-radius: 6px; padding: 6px 10px; font-size: 13px; }
.summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 22px; }
.card { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 15px; }
.label { color: var(--muted); font-size: 12px; margin-bottom: 8px; }
.value { font-size: 26px; font-weight: 700; }
.red { color: var(--red); } .amber { color: var(--amber); } .green { color: var(--green); } .blue { color: var(--blue); }
.section { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 18px; margin-bottom: 18px; }
.note { color: var(--muted); margin: -4px 0 14px; line-height: 1.6; }
.empty { color: var(--green); margin: 0; }
.table-scroll { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; min-width: 760px; }
th, td { border-bottom: 1px solid var(--line); padding: 9px 10px; text-align: left; vertical-align: top; font-size: 13px; }
th { color: var(--muted); font-weight: 600; background: #111a24; position: sticky; top: 0; }
td { color: #e7edf5; }
.risk-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 12px; }
.small { font-size: 12px; color: var(--muted); }
</style>
</head>
<body>
<div class="wrap">
  <div class="header">
    <h1>Log Analytics 合并风险报告</h1>
    <div class="meta">
      <span class="tag">查询时间段: $(Escape-Html $AnalysisDate)</span>
      <span class="tag">数据表: $tableCount</span>
      <span class="tag">总记录数: $totalRecords</span>
      <span class="tag">可信 IP 规则: $trustedCount</span>
      <span class="tag">非工作时间范围: 21:00 - 08:00</span>
    </div>
  </div>

  <div class="summary">
    <div class="card"><div class="label">登录失败</div><div class="value red">$($riskCounts.FailedSignins)</div></div>
    <div class="card"><div class="label">失败/异常操作</div><div class="value red">$($riskCounts.FailedOperations)</div></div>
    <div class="card"><div class="label">删除 / Disable</div><div class="value amber">$($riskCounts.DeleteDisable)</div></div>
    <div class="card"><div class="label">可疑 IP</div><div class="value amber">$($riskCounts.SuspiciousIPs)</div></div>
    <div class="card"><div class="label">可信位置外成功登录</div><div class="value amber">$($riskCounts.SuspiciousSigninSuccess)</div></div>
    <div class="card"><div class="label">邮箱低容量</div><div class="value red">$($riskCounts.MailboxLowSpace)</div></div>
    <div class="card"><div class="label">SharedMailbox</div><div class="value blue">$($riskCounts.SharedMailboxes)</div></div>
    <div class="card"><div class="label">AAD Disabled 用户</div><div class="value amber">$($riskCounts.DisabledUsers)</div></div>
  </div>

  <div class="section">
    <h2>AAD / Managed Identity / Service Principal 登录失败</h2>
    <p class="note">Managed Identity 或 Service Principal 登录失败可能表示依赖该身份的服务无法正常运行。</p>
    $failedSigninHtml
  </div>

  <div class="section">
    <h2>删除 / Disable 操作</h2>
    <p class="note">只统计 delete / remove / disable / deactivate 语义的操作。</p>
    $deleteDisableHtml
  </div>

  <div class="section">
    <h2>可疑成功登录</h2>
    <p class="note">SigninLogs 中 IP 不在可信位置内，且登录应用不属于 Windows Sign In / Microsoft Edge / Sangfor SASE VPN / Microsoft Office 的成功登录。</p>
    $signinSuspiciousHtml
  </div>

  <div class="section">
    <h2>可疑 IP</h2>
    <p class="note">已排除 TrustedLocation_KJ.txt 和 TrustedLocation_IDC_Ali.txt 中的可信 IP。</p>
    $suspiciousIpHtml
  </div>

  <div class="section">
    <h2>客户端 IP 排行</h2>
    <p class="note">此排行已排除所有出现在“可疑 IP”中的 IP。</p>
    $topClientIpHtml
  </div>

  <div class="section">
    <h2>非工作时间活动</h2>
    <p class="note">非工作时间范围：21:00 - 08:00。</p>
    $offHoursHtml
  </div>

  <div class="section">
    <h2>License 使用量与剩余数量</h2>
    <p class="note">自动从 AssignedLicensesDCR_CL 中整理使用量最高的 4 种 License。若源数据没有总量字段，剩余数量显示 N/A。</p>
    $licenseHtml
  </div>

  <div class="section">
    <h2>邮箱容量风险</h2>
    <p class="note">AvailableSpaceGB 低于 QuotaLimitGB 的 5% 时列为风险；邮箱用量耗尽可能导致无法收发邮件。</p>
    $mailboxRiskHtml
  </div>

  <div class="section">
    <h2>SharedMailbox</h2>
    <p class="note">显示 SharedMailbox 数量、大小以及对应 AAD 用户是否被 disable；DisplayName 通过 AzureADUsersDCR_CL 关联。</p>
    $sharedMailboxHtml
  </div>

  <div class="section">
    <h2>MessageTrace 异常 / PBI / SkyGuard 相关</h2>
    <p class="note">关注 failed、blocked、quarantined、rejected、undelivered、PBI、Power BI、SkyGuard 等状态或关键词。</p>
    $messageTraceHtml
  </div>

  <div class="section">
    <h2>Managed Identity / SP / 应用权限变动</h2>
    <p class="note">关注 permission、consent、credential、secret、certificate、app role、service principal、managed identity 等审计关键词。</p>
    $permissionHtml
  </div>

  <div class="section">
    <h2>AAD 用户信息</h2>
    <p class="note">AzureADUsersDCR_CL 用作 AAD 用户信息源；其他表仅含 UPN 但需要 displayName 时会 join 此表查询。</p>
    <div class="risk-grid">
      <div>$disabledUsersHtml</div>
      <div>$missingUsersHtml</div>
    </div>
  </div>

  <div class="section">
    <h2>其他失败/异常操作</h2>
    $failedOpsHtml
  </div>

  <div class="section">
    <h2>数据源查询状态</h2>
    $sourceStatusHtml
  </div>

  <p class="small">Generated at $(Escape-Html ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')))</p>
</div>
</body>
</html>
"@

$utf8Bom = [System.Text.UTF8Encoding]::new($true)
[System.IO.File]::WriteAllText($OutputPath, $html, $utf8Bom)
Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
