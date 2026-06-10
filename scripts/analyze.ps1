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

function Get-RowEventCount {
    param([object]$Row)

    $count = Get-NumberValue (Get-AnyFieldValue -Row $Row -Names @('EventCount', 'Count', 'RecordCount') -Default '')
    if ($null -eq $count -or $count -lt 1) { return 1 }
    return [int]$count
}

function Get-EventCountSum {
    param([object[]]$Rows)

    $sum = 0
    foreach ($row in @($Rows)) {
        $sum += Get-RowEventCount -Row $row
    }
    return $sum
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
    $eventCount = Get-RowEventCount -Row $Row
    $firstTime = Get-AnyFieldValue -Row $Row -Names @('FirstTime', 'StartTime', 'MinTime') -Default ''
    $lastTime = Get-AnyFieldValue -Row $Row -Names @('LastTime', 'EndTime', 'MaxTime') -Default ''
    if (-not $firstTime) { $firstTime = [string]$Row.TimeGenerated }
    if (-not $lastTime) { $lastTime = [string]$Row.TimeGenerated }
    $activityDateTime = Get-AnyFieldValue -Row $Row -Names @('ActivityDateTime') -Default $lastTime
    $detailFields = @('ResultDescription', 'ResultReason', 'FailureReason', 'Status', 'DeliveryStatus', 'ErrorCode', 'ResultType', 'Subject', 'Message', 'ErrorMessage', 'PermissionName', 'TargetResources', 'ModifiedProperties')
    if ($Table -eq 'AuditLogs') {
        $detailFields = @('PermissionName', 'ResultDescription', 'ResultReason', 'FailureReason', 'Status', 'ResultType')
    }
    if ($Table -eq 'DCRLogErrors') {
        $detailFields = @('Message', 'ErrorMessage', 'Details', 'Description', 'Status')
    }
    $detail = Get-AnyFieldValue -Row $Row -Names $detailFields -Default ''
    if ($Table -eq 'MessageTraceDataDCR_CL') {
        $traceStatus = Get-AnyFieldValue -Row $Row -Names @('Status', 'DeliveryStatus', 'EventType', 'Action', 'Result') -Default ''
        $subject = Get-AnyFieldValue -Row $Row -Names @('Subject') -Default ''
        $detail = (@($traceStatus, $subject) | Where-Object { $_ }) -join ' | '
    }
    $target = Get-AnyFieldValue -Row $Row -Names @('Target', 'TargetDisplayName', 'TargetResources', 'TargetResource', 'ObjectId', 'InputStreamId', 'ResourceDisplayName', 'AppDisplayName', 'ServicePrincipalName', 'DisplayName', 'RecipientAddress', 'Subject') -Default ''
    $permissionName = Get-AnyFieldValue -Row $Row -Names @('PermissionName', 'AppRoleDisplayName') -Default ''

    return [PSCustomObject]@{
        Table = $Table
        Time = Get-LocalTimeText -Value $lastTime
        ActivityDateTime = Get-LocalTimeText -Value $activityDateTime
        FirstTime = Get-LocalTimeText -Value $firstTime
        LastTime = Get-LocalTimeText -Value $lastTime
        Count = $eventCount
        User = $user
        Operation = $op
        IP = $ip
        Status = $status
        Target = $target
        PermissionName = $permissionName
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

function Get-TimeValueForSort {
    param([string]$Value)
    try { return [DateTime]::Parse($Value) } catch { return [DateTime]::MinValue }
}

function Get-ShortListText {
    param(
        [string[]]$Values,
        [int]$MaxItems = 3
    )

    $items = @($Values | Where-Object { $_ } | Sort-Object -Unique)
    if ($items.Count -le $MaxItems) { return ($items -join ', ') }
    return (($items | Select-Object -First $MaxItems) -join ', ') + " ... +$($items.Count - $MaxItems)"
}

function Group-EventRecords {
    param(
        [object[]]$Rows,
        [scriptblock]$KeyBuilder
    )

    $groups = @{}
    foreach ($row in @($Rows)) {
        $key = & $KeyBuilder $row
        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = [System.Collections.Generic.List[object]]::new()
        }
        $groups[$key].Add($row) | Out-Null
    }

    $groupedRows = @(
        foreach ($entry in $groups.GetEnumerator()) {
            $items = @($entry.Value)
            $firstTimes = @($items | ForEach-Object { Get-TimeValueForSort -Value $(if ($_.FirstTime) { $_.FirstTime } else { $_.Time }) } | Where-Object { $_ -ne [DateTime]::MinValue } | Sort-Object)
            $lastTimes = @($items | ForEach-Object { Get-TimeValueForSort -Value $(if ($_.LastTime) { $_.LastTime } else { $_.Time }) } | Where-Object { $_ -ne [DateTime]::MinValue } | Sort-Object)
            [PSCustomObject]@{
                Count = Get-EventCountSum -Rows $items
                FirstTime = if ($firstTimes.Count -gt 0) { $firstTimes[0].ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
                LastTime = if ($lastTimes.Count -gt 0) { $lastTimes[-1].ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
                Table = Get-ShortListText -Values @($items | ForEach-Object { $_.Table })
                User = Get-ShortListText -Values @($items | ForEach-Object { $_.User })
                Operation = Get-ShortListText -Values @($items | ForEach-Object { $_.Operation })
                IP = Get-ShortListText -Values @($items | ForEach-Object { $_.IP })
                Target = Get-ShortListText -Values @($items | ForEach-Object { $_.Target }) -MaxItems 2
                PermissionName = Get-ShortListText -Values @($items | ForEach-Object { $_.PermissionName }) -MaxItems 2
                Detail = Get-ShortListText -Values @($items | ForEach-Object { $_.Detail }) -MaxItems 2
                Reason = Get-ShortListText -Values @($items | ForEach-Object { $_.Reason }) -MaxItems 2
                ActivityDateTime = if ($lastTimes.Count -gt 0) { $lastTimes[-1].ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            }
        }
    )
    return @($groupedRows | Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.LastTime }; Descending = $true })
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
    $statusText = @(
        (Get-AnyFieldValue -Row $Row -Names @('Status', 'DeliveryStatus', 'EventType', 'Action', 'Result') -Default ''),
        (Get-AnyFieldValue -Row $Row -Names @('ErrorCode', 'Reason', 'ResultDescription') -Default '')
    ) -join ' '
    $messageText = @(
        (Get-AnyFieldValue -Row $Row -Names @('Subject') -Default ''),
        (Get-AnyFieldValue -Row $Row -Names @('SenderAddress', 'RecipientAddress') -Default '')
    ) -join ' '
    $allText = "$statusText $messageText"
    $criticalStatus = $statusText -match '(?i)\b(fail(ed|ure)?|blocked|quarantined|reject(ed)?|undeliver(ed|able)?|error|timeout|bounced)\b'
    $businessMonitor = $allText -match '(?i)\b(Power\s*BI|PBI|SkyGuard)\b'
    $businessProblem = $allText -match '(?i)\b(fail(ed|ure)?|blocked|quarantined|reject(ed)?|undeliver(ed|able)?|error|timeout|bounced)\b'
    return ($criticalStatus -or ($businessMonitor -and $businessProblem))
}

function Get-MessageTracePlainExplanation {
    param([string]$Text)
    if ($Text -match '(?i)quarantine') { return '邮件被隔离，收件人通常无法直接收到；需要检查安全策略或放行记录。' }
    if ($Text -match '(?i)block') { return '邮件被阻止，通常是安全策略、网关或反垃圾规则拦截。' }
    if ($Text -match '(?i)reject|bounce') { return '邮件被拒收或退信，发件系统认为消息没有成功送达。' }
    if ($Text -match '(?i)undeliver') { return '邮件未送达，可能影响业务通知或报表邮件接收。' }
    if ($Text -match '(?i)timeout') { return '邮件投递超时，可能是目标服务器或网关暂时不可达。' }
    if ($Text -match '(?i)PBI|Power\s*BI') { return 'Power BI / PBI 相关邮件异常，可能影响报表刷新或通知送达。' }
    if ($Text -match '(?i)SkyGuard') { return 'SkyGuard 相关邮件异常，需要结合网关侧告警确认拦截原因。' }
    return '邮件投递失败或异常，需要确认发件方、收件方和邮件网关状态。'
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

function Test-AuditLogUserActor {
    param([object]$Row)

    $actor = Get-AnyFieldValue -Row $Row -Names @('Actor', 'InitiatedByUserPrincipalName', 'ActorUserPrincipalName', 'UserPrincipalName', 'Identity') -Default ''
    $initiated = Get-AnyFieldValue -Row $Row -Names @('InitiatedBy') -Default ''
    $upnFromJson = ''
    if ($initiated) {
        $match = [regex]::Match($initiated, '"userPrincipalName"\s*:\s*"([^"]+)"')
        if ($match.Success) { $upnFromJson = $match.Groups[1].Value }
    }
    return ((-not [string]::IsNullOrWhiteSpace($upnFromJson)) -or ($actor -match '@') -or ($initiated -match '"user"\s*:'))
}

function Test-PimAuditNoise {
    param([object]$Row)

    $text = @(
        (Get-OperationValue -Row $Row -TableName 'AuditLogs'),
        (Get-AnyFieldValue -Row $Row -Names @('ActivityDisplayName', 'ResultReason', 'ResultDescription', 'TargetResources', 'ModifiedProperties') -Default '')
    ) -join ' '
    return ($text -match '(?i)\bPIM\b|PIM activation expired')
}

function Test-ServicePrincipalAuditOperation {
    param([string]$Operation)

    if ([string]::IsNullOrWhiteSpace($Operation)) { return $false }
    return @(
        'Add service principal',
        'Remove service principal',
        'Hard delete service principal',
        'Add app role assignment to service principal',
        'Remove app role assignment from service principal'
    ) -contains $Operation.Trim()
}

function Normalize-LicenseKey {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    return ([regex]::Replace($Name.ToUpperInvariant(), '[^A-Z0-9]', ''))
}

function ConvertFrom-AccessTokenValue {
    param([object]$Token)
    if ($null -eq $Token) { return '' }
    if ($Token -is [securestring]) {
        $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Token)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
    }
    return [string]$Token
}

function Get-LicenseSkuTotalsFromGraph {
    $result = [PSCustomObject]@{
        Success = $false
        Message = ''
        Skus = @{}
    }

    try {
        if (-not (Get-Command Get-AzAccessToken -ErrorAction SilentlyContinue)) {
            $result.Message = '无法获取 License 总量：当前 PowerShell 会话未加载 Az.Accounts / Get-AzAccessToken。'
            return $result
        }

        $resourceUrl = 'https://graph.microsoft.com/'
        $endpoint = 'https://graph.microsoft.com/v1.0/subscribedSkus'
        $context = $null
        if (Get-Command Get-AzContext -ErrorAction SilentlyContinue) {
            $context = Get-AzContext -ErrorAction SilentlyContinue
        }
        if ($context -and $context.Environment -and $context.Environment.Name -eq 'AzureChinaCloud') {
            $resourceUrl = 'https://microsoftgraph.chinacloudapi.cn/'
            $endpoint = 'https://microsoftgraph.chinacloudapi.cn/v1.0/subscribedSkus'
        }

        $tokenResponse = Get-AzAccessToken -ResourceUrl $resourceUrl -ErrorAction Stop
        $token = ConvertFrom-AccessTokenValue -Token $tokenResponse.Token
        if ([string]::IsNullOrWhiteSpace($token)) {
            throw 'Get-AzAccessToken returned an empty token.'
        }

        $response = Invoke-RestMethod -Method Get -Uri $endpoint -Headers @{ Authorization = "Bearer $token" } -ErrorAction Stop
        $map = @{}
        foreach ($sku in @($response.value)) {
            $name = [string]$sku.skuPartNumber
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $enabled = $null
            if ($sku.prepaidUnits -and $null -ne $sku.prepaidUnits.enabled) {
                $enabled = [int]$sku.prepaidUnits.enabled
            }
            $consumed = if ($null -ne $sku.consumedUnits) { [int]$sku.consumedUnits } else { $null }
            $remaining = if ($null -ne $enabled -and $null -ne $consumed) { [Math]::Max(0, $enabled - $consumed) } else { $null }
            $map[$name.ToUpperInvariant()] = [PSCustomObject]@{
                License = $name
                Used = $consumed
                Total = $enabled
                Remaining = $remaining
            }
            $normalizedName = Normalize-LicenseKey -Name $name
            if ($normalizedName) {
                $map[$normalizedName] = $map[$name.ToUpperInvariant()]
            }
        }
        $result.Success = $true
        $result.Message = "License 总量已通过 Microsoft Graph subscribedSkus 获取。"
        $result.Skus = $map
        return $result
    } catch {
        $result.Message = "无法通过 Microsoft Graph 获取 License 总量：$($_.Exception.Message)"
        return $result
    }
}

function New-ReportSection {
    param(
        [string]$Id,
        [string]$Title,
        [string]$Note,
        [string]$Content,
        [bool]$Open = $true
    )
    $openText = if ($Open) { ' open' } else { '' }
    $noteHtml = if ([string]::IsNullOrWhiteSpace($Note)) { '' } else { '<p class="note">' + (Escape-Html $Note) + '</p>' }
    return @"
  <details class="section" id="$(Escape-Html $Id)"$openText>
    <summary><span>$(Escape-Html $Title)</span></summary>
    $noteHtml
    $Content
  </details>
"@
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
$suspiciousSigninSuccess = [System.Collections.Generic.List[object]]::new()
$suspiciousIpReasons = @{}
$clientIpCounts = @{}
$identityPermissionChanges = [System.Collections.Generic.List[object]]::new()
$dcrLogErrorRows = [System.Collections.Generic.List[object]]::new()
$intuneAuditRows = [System.Collections.Generic.List[object]]::new()
$sourceStatusRows = [System.Collections.Generic.List[object]]::new()

foreach ($dataset in $datasets) {
    $table = $dataset.Table
    $sourceStatusRows.Add([PSCustomObject]@{ Table = $table; Records = $dataset.Rows.Count; Source = (Split-Path -Leaf $dataset.Path) }) | Out-Null

    foreach ($row in $dataset.Rows) {
        if ($table -eq 'AuditLogs') {
            if (-not (Test-AuditLogUserActor -Row $row)) { continue }
            if (Test-PimAuditNoise -Row $row) { continue }
        }

        $rowEventCount = Get-RowEventCount -Row $row
        $op = Get-OperationValue -Row $row -TableName $table
        $success = Get-SuccessValue -Row $row -TableName $table
        $ip = Get-NormalizedIpValue -IP (Get-ClientIpValue -Row $row -TableName $table)
        $isUsablePublicIp = -not (Test-PrivateOrInvalidIp -IP $ip)
        $isTrustedIp = Test-IpInTrustedRules -IP $ip -Rules $trustedRules

        if ($isUsablePublicIp) {
            Add-Count -Map $clientIpCounts -Key $ip -By $rowEventCount
        }

        if ($success -eq 'false') {
            $record = New-EventRecord -Table $table -Row $row -Reason '失败/异常'
            if ($table -in @('AADManagedIdentitySignInLogs', 'AADServicePrincipalSignInLogs', 'SigninLogs')) {
                if ($table -ne 'AADServicePrincipalSignInLogs' -or $rowEventCount -gt 10) {
                    $failedSignins.Add($record) | Out-Null
                }
            } elseif ($table -in @('DCRLogErrors', 'IntuneAuditLogsDCR_CL')) {
                # These tables have dedicated sections below.
            } else {
                $failedOperations.Add($record) | Out-Null
            }
        }

        if (Test-DeleteOrDisableOperation -Operation $op -TableName $table) {
            $deleteDisableEvents.Add((New-EventRecord -Table $table -Row $row -Reason '删除 / Disable 操作')) | Out-Null
        }

        if ($table -in @('AADManagedIdentitySignInLogs', 'AADServicePrincipalSignInLogs', 'SigninLogs') -and $success -eq 'true' -and $isUsablePublicIp -and -not $isTrustedIp) {
            $appName = Get-SigninAppName -Row $row
            $isAllowedInteractiveApp = ($table -eq 'SigninLogs' -and (Test-AllowedSigninApp -AppName $appName))
            if (-not $isAllowedInteractiveApp) {
                $event = New-EventRecord -Table $table -Row $row -Reason "可信位置外成功登录，应用：$appName"
                $suspiciousSigninSuccess.Add($event) | Out-Null
                if ($table -eq 'SigninLogs') {
                    $suspiciousIpReasons[$ip] = 'SigninLogs 可信位置外成功登录'
                }
            }
        }

        if ($table -eq 'SigninLogs' -and $isUsablePublicIp -and -not $isTrustedIp) {
            $workload = Get-WorkloadValue -Row $row -TableName $table
            if (-not $suspiciousIpReasons.ContainsKey($ip)) {
                $suspiciousIpReasons[$ip] = "公共 IP，来源表/工作负载：$table / $workload"
            }
        }

        if ($table -eq 'DCRLogErrors') {
            $dcrLogErrorRows.Add((New-EventRecord -Table $table -Row $row -Reason 'DCR 日志采集错误')) | Out-Null
        }

        if ($table -eq 'IntuneAuditLogsDCR_CL') {
            $intuneAuditRows.Add((New-EventRecord -Table $table -Row $row -Reason 'Intune 审计风险')) | Out-Null
        }

        if ($table -eq 'AuditLogs' -and $success -eq 'true' -and (Test-ServicePrincipalAuditOperation -Operation $op)) {
            $identityPermissionChanges.Add((New-EventRecord -Table $table -Row $row -Reason 'Service Principal 对象 / 权限成功变动')) | Out-Null
        }
    }
}

# License usage: infer the four license names from the data, then use Graph subscribedSkus when log totals are missing.
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
            [PSCustomObject]@{ License = $_.Name; Used = $used; Total = $(if ($null -ne $_.Total) { [int]$_.Total } else { 'N/A' }); Remaining = $remaining; Source = $(if ($null -ne $_.Total) { 'Log' } else { 'Pending' }) }
        }
)
$licenseStatusNote = 'License 使用量优先使用 AssignedLicensesDCR_CL；若日志缺少总量，会尝试调用 Microsoft Graph subscribedSkus 获取总量和剩余量。'
$licensesNeedGraph = @($licenseUsage | Where-Object { $_.Total -eq 'N/A' -or $_.Remaining -eq 'N/A' })
if ($licensesNeedGraph.Count -gt 0 -or $licenseUsage.Count -eq 0) {
    $graphLicenseResult = Get-LicenseSkuTotalsFromGraph
    if ($graphLicenseResult.Success) {
        if ($licenseUsage.Count -eq 0) {
            $licenseUsage = @(
                $graphLicenseResult.Skus.Values |
                    Sort-Object @{ Expression = { if ($null -ne $_.Used) { $_.Used } else { 0 } }; Descending = $true }, License |
                    Select-Object -Unique -First 4 |
                    ForEach-Object {
                        [PSCustomObject]@{ License = $_.License; Used = $_.Used; Total = $_.Total; Remaining = $_.Remaining; Source = 'Graph' }
                    }
            )
        }

        foreach ($license in $licenseUsage) {
            $keys = @(([string]$license.License).ToUpperInvariant(), (Normalize-LicenseKey -Name ([string]$license.License))) | Where-Object { $_ } | Select-Object -Unique
            $matchedKey = @($keys | Where-Object { $graphLicenseResult.Skus.ContainsKey($_) } | Select-Object -First 1)
            if ($matchedKey.Count -gt 0) {
                $sku = $graphLicenseResult.Skus[$matchedKey[0]]
                if ($null -ne $sku.Used) { $license.Used = $sku.Used }
                if ($null -ne $sku.Total) { $license.Total = $sku.Total }
                if ($null -ne $sku.Remaining) { $license.Remaining = $sku.Remaining }
                $license.Source = 'Graph'
            } elseif ($license.Total -eq 'N/A') {
                $license.Source = 'Missing'
            }
        }
    }
    $licenseStatusNote = "$licenseStatusNote $($graphLicenseResult.Message)"
}

$mailboxRisks = [System.Collections.Generic.List[object]]::new()
$sharedMailboxRows = [System.Collections.Generic.List[object]]::new()
$mailboxRows = @($datasets | Where-Object { $_.Table -eq 'MailboxStatisticsDCR_CL' } | ForEach-Object { $_.Rows })
foreach ($row in $mailboxRows) {
    $available = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('AvailableSpaceGB', 'AvailableSpaceInGB', 'AvailableSpace') -Default '')
    $quota = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('QuotaLimitGB', 'QuotaGB', 'StorageQuotaGB', 'ProhibitSendReceiveQuotaGB') -Default '')
    $user = Get-UserValue -Row $row -TableName 'MailboxStatisticsDCR_CL'
    $size = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('TotalItemSizeGB', 'TotalItemSizeInGB', 'MailboxSizeGB', 'MailboxSize', 'SizeGB', 'TotalSizeGB', 'TotalItemSize', 'StorageUsedGB', 'StorageUsed') -Default '')
    if ($null -eq $size -and $null -ne $available -and $null -ne $quota -and $quota -ge $available) {
        $size = $quota - $available
    }
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
        $sharedMailboxRows.Add([PSCustomObject]@{
            User = Format-UserForReport -User $user
            Type = if ($type) { $type } else { 'SharedMailbox' }
            SizeGB = if ($null -ne $size) { [Math]::Round($size, 2) } else { 'N/A' }
            AvailableGB = if ($null -ne $available) { [Math]::Round($available, 2) } else { 'N/A' }
            QuotaGB = if ($null -ne $quota) { [Math]::Round($quota, 2) } else { 'N/A' }
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
    FailedSignins = Get-EventCountSum -Rows $failedSignins
    FailedOperations = Get-EventCountSum -Rows $failedOperations
    DeleteDisable = Get-EventCountSum -Rows $deleteDisableEvents
    SuspiciousIPs = $suspiciousIpRows.Count
    SuspiciousSigninSuccess = Get-EventCountSum -Rows $suspiciousSigninSuccess
    LicenseTypes = $licenseUsage.Count
    MailboxLowSpace = $mailboxRisks.Count
    SharedMailboxes = $sharedMailboxRows.Count
    IdentityPermissionChanges = Get-EventCountSum -Rows $identityPermissionChanges
    DcrLogErrors = Get-EventCountSum -Rows $dcrLogErrorRows
    IntuneAudit = Get-EventCountSum -Rows $intuneAuditRows
}

$totalRecords = ($datasets | ForEach-Object { $_.Rows.Count } | Measure-Object -Sum).Sum
$tableCount = $datasets.Count
$trustedCount = $trustedRules.Count
$microsoftTrustedCount = @($trustedRules | Where-Object { $_.Source -eq 'MicrosoftServiceTags' }).Count
$microsoftTrustedNote = if ($microsoftTrustedCount -gt 0) {
    "并已排除 Microsoft Service Tags 中的 $microsoftTrustedCount 条 CIDR 规则。"
} else {
    'Microsoft Service Tags 缓存不可用或未下载，本次仅使用本地可信 IP 文件。'
}

$failedSigninGrouped = Group-EventRecords -Rows $failedSignins -KeyBuilder { param($r) $r.IP }
$failedSigninHtml = New-TableHtml -Rows ($failedSigninGrouped | Select-Object -First 50) -Columns @('次数', '首次时间', '最后时间', 'IP', '主体/应用摘要', '说明') -CellBuilder {
    param($r) @($r.Count, $r.FirstTime, $r.LastTime, $r.IP, (($r.User, $r.Operation) -join ' / '), $r.Detail)
}
$failedOpsGrouped = Group-EventRecords -Rows $failedOperations -KeyBuilder { param($r) "$($r.Table)|$($r.User)|$($r.Operation)|$($r.Detail)" }
$failedOpsHtml = New-TableHtml -Rows ($failedOpsGrouped | Select-Object -First 50) -Columns @('次数', '首次时间', '最后时间', '表', '用户', '操作', '状态/原因') -CellBuilder {
    param($r) @($r.Count, $r.FirstTime, $r.LastTime, $r.Table, $r.User, $r.Operation, $r.Detail)
}
$deleteDisableGrouped = Group-EventRecords -Rows $deleteDisableEvents -KeyBuilder { param($r) "$($r.Table)|$($r.User)|$($r.Operation)|$($r.Detail)" }
$deleteDisableHtml = New-TableHtml -Rows ($deleteDisableGrouped | Select-Object -First 80) -Columns @('次数', '首次时间', '最后时间', '表', '操作者', '操作', '结果/说明') -CellBuilder {
    param($r) @($r.Count, $r.FirstTime, $r.LastTime, $r.Table, $r.User, $r.Operation, $r.Detail)
}
$suspiciousIpHtml = New-TableHtml -Rows $suspiciousIpRows -Columns @('IP', '原因') -CellBuilder {
    param($r) @($r.IP, $r.Reason)
}
$signinSuspiciousGrouped = Group-EventRecords -Rows $suspiciousSigninSuccess -KeyBuilder { param($r) "$($r.User)|$($r.Operation)|$($r.IP)" }
$signinSuspiciousHtml = New-TableHtml -Rows ($signinSuspiciousGrouped | Select-Object -First 80) -Columns @('次数', '首次时间', '最后时间', '用户', '应用', 'IP', '说明') -CellBuilder {
    param($r) @($r.Count, $r.FirstTime, $r.LastTime, $r.User, $r.Operation, $r.IP, $r.Reason)
}
$topClientIpHtml = New-TableHtml -Rows $topClientIps -Columns @('IP', '次数') -CellBuilder {
    param($r) @($r.IP, $r.Count)
}
$licenseHtml = New-TableHtml -Rows $licenseUsage -Columns @('License 名称', '已使用', '总数', '剩余', '来源') -CellBuilder {
    param($r) @($r.License, $r.Used, $r.Total, $r.Remaining, $r.Source)
}
$mailboxRiskHtml = New-TableHtml -Rows ($mailboxRisks | Sort-Object AvailableGB | Select-Object -First 50) -Columns @('邮箱', 'AvailableSpaceGB', 'QuotaLimitGB', '使用率', '风险') -CellBuilder {
    param($r) @($r.User, $r.AvailableGB, $r.QuotaGB, $r.Usage, $r.Reason)
}
$sharedMailboxHtml = New-TableHtml -Rows ($sharedMailboxRows | Sort-Object SizeGB -Descending | Select-Object -First 80) -Columns @('SharedMailbox', '类型', '大小GB', 'AvailableSpaceGB', 'QuotaLimitGB') -CellBuilder {
    param($r) @($r.User, $r.Type, "$($r.SizeGB) GB", $r.AvailableGB, $r.QuotaGB)
}
$permissionGrouped = Group-EventRecords -Rows $identityPermissionChanges -KeyBuilder { param($r) "$($r.User)|$($r.Operation)|$($r.Target)|$($r.PermissionName)" }
$permissionHtml = New-TableHtml -Rows ($permissionGrouped | Select-Object -First 80) -Columns @('Timestamp(ActivityDateTime)', 'Actor', 'Operation', 'Target', 'Permission') -CellBuilder {
    param($r) @($r.ActivityDateTime, $r.User, $r.Operation, $r.Target, $r.PermissionName)
}
$dcrLogErrorGrouped = Group-EventRecords -Rows $dcrLogErrorRows -KeyBuilder { param($r) "$($r.Target)|$($r.Operation)|$($r.Detail)" }
$dcrLogErrorHtml = New-TableHtml -Rows ($dcrLogErrorGrouped | Select-Object -First 80) -Columns @('次数', '首次时间', '最后时间', 'InputStreamId', 'OperationName', 'Message') -CellBuilder {
    param($r) @($r.Count, $r.FirstTime, $r.LastTime, $r.Target, $r.Operation, $r.Detail)
}
$intuneGrouped = Group-EventRecords -Rows $intuneAuditRows -KeyBuilder { param($r) "$($r.User)|$($r.Operation)|$($r.Target)|$($r.Detail)" }
$intuneHtml = New-TableHtml -Rows ($intuneGrouped | Select-Object -First 80) -Columns @('次数', '首次时间', '最后时间', 'Actor', 'Operation', 'Target', '结果/说明') -CellBuilder {
    param($r) @($r.Count, $r.FirstTime, $r.LastTime, $r.User, $r.Operation, $r.Target, $r.Detail)
}
$sourceStatusHtml = New-TableHtml -Rows $sourceStatusRows -Columns @('表', '记录数', 'CSV') -CellBuilder {
    param($r) @($r.Table, $r.Records, $r.Source)
}

$sectionSpecs = @(
    [PSCustomObject]@{ Id = 'failed-signins'; Title = 'AAD / Managed Identity / Service Principal 登录失败'; Note = 'Managed Identity 或 Service Principal 登录失败可能表示依赖该身份的服务无法正常运行。相同 IP 的多次失败已合并。'; Content = $failedSigninHtml; Open = $true },
    [PSCustomObject]@{ Id = 'delete-disable'; Title = '删除 / Disable 操作'; Note = '只统计 delete / remove / disable / deactivate 语义的操作；AuditLogs 已去掉目标字段，并按除时间外相同的记录合并。'; Content = $deleteDisableHtml; Open = $true },
    [PSCustomObject]@{ Id = 'suspicious-success'; Title = '可疑成功登录'; Note = '关注 AADManagedIdentitySignInLogs / AADServicePrincipalSignInLogs / SigninLogs 三张表；SigninLogs 仍排除 Windows Sign In / Microsoft Edge / Sangfor SASE VPN / Microsoft Office。相同主体、应用、IP 已合并。'; Content = $signinSuspiciousHtml; Open = $true },
    [PSCustomObject]@{ Id = 'suspicious-ip'; Title = '可疑 IP'; Note = "仅统计 SigninLogs 中的可疑 IP；已排除 TrustedLocation_KJ.txt、TrustedLocation_IDC_Ali.txt 中的可信 IP，$microsoftTrustedNote"; Content = $suspiciousIpHtml; Open = $true },
    [PSCustomObject]@{ Id = 'client-ip-rank'; Title = '客户端 IP 排行'; Note = '此排行已排除所有出现在“可疑 IP”中的 IP。'; Content = $topClientIpHtml; Open = $false },
    [PSCustomObject]@{ Id = 'license'; Title = 'License 使用量与剩余数量'; Note = $licenseStatusNote; Content = $licenseHtml; Open = $true },
    [PSCustomObject]@{ Id = 'mailbox-low-space'; Title = '邮箱容量风险'; Note = 'AvailableSpaceGB 低于 QuotaLimitGB 的 5% 时列为风险；邮箱用量耗尽可能导致无法收发邮件。'; Content = $mailboxRiskHtml; Open = $true },
    [PSCustomObject]@{ Id = 'shared-mailbox'; Title = 'SharedMailbox'; Note = '显示 SharedMailbox 数量、大小 GB、可用空间和配额。'; Content = $sharedMailboxHtml; Open = $false },
    [PSCustomObject]@{ Id = 'identity-permission'; Title = 'Service Principal 对象 / 权限成功变动'; Note = '仅显示用户操作者触发的 Add/Remove/Hard delete service principal，以及 Add/Remove app role assignment to service principal；PIM 相关记录已排除。'; Content = $permissionHtml; Open = $true },
    [PSCustomObject]@{ Id = 'dcr-log-errors'; Title = 'DCRLogErrors'; Note = '按最近 30 天的 InputStreamId、OperationName、Message 去重统计。'; Content = $dcrLogErrorHtml; Open = $true },
    [PSCustomObject]@{ Id = 'intune-audit'; Title = 'Intune 审计风险'; Note = '显示 IntuneAuditLogsDCR_CL 中失败、异常、删除或禁用类审计记录，字段已按 Actor / Operation / Target 提取。'; Content = $intuneHtml; Open = $true },
    [PSCustomObject]@{ Id = 'failed-ops'; Title = '其他失败/异常操作'; Note = '登录失败和删除 / Disable 已单独列出，这里保留其他失败、异常记录，并按相似内容合并。'; Content = $failedOpsHtml; Open = $false },
    [PSCustomObject]@{ Id = 'source-status'; Title = '数据源查询状态'; Note = '显示本次参与生成合并报告的 CSV 数据源。'; Content = $sourceStatusHtml; Open = $false }
)

$sideNavHtml = '<nav class="side-nav"><div class="nav-title">目录</div>'
foreach ($section in $sectionSpecs) {
    $sideNavHtml += '<a href="#' + (Escape-Html $section.Id) + '">' + (Escape-Html $section.Title) + '</a>'
}
$sideNavHtml += '</nav>'

$reportSectionsHtml = @(
    foreach ($section in $sectionSpecs) {
        New-ReportSection -Id $section.Id -Title $section.Title -Note $section.Note -Content $section.Content -Open $section.Open
    }
) -join "`r`n"

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
html { scroll-behavior: smooth; }
body { margin: 0; background: var(--bg); color: var(--text); font-family: "Segoe UI", Arial, sans-serif; }
.layout { display: grid; grid-template-columns: 250px minmax(0, 1fr); gap: 22px; max-width: 1680px; margin: 0 auto; padding: 28px; }
.wrap { min-width: 0; }
.side-nav { position: sticky; top: 18px; align-self: start; max-height: calc(100vh - 36px); overflow: auto; background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 12px; }
.nav-title { color: var(--muted); font-size: 12px; margin: 2px 6px 10px; }
.side-nav a { display: block; color: var(--text); text-decoration: none; border-radius: 6px; padding: 8px 9px; font-size: 13px; line-height: 1.35; }
.side-nav a:hover { background: var(--panel2); }
.header { border-bottom: 1px solid var(--line); padding-bottom: 18px; margin-bottom: 22px; }
h1 { margin: 0 0 10px; font-size: 28px; font-weight: 700; }
.meta { display: flex; gap: 10px; flex-wrap: wrap; color: var(--muted); }
.tag { background: var(--panel2); border: 1px solid var(--line); border-radius: 6px; padding: 6px 10px; font-size: 13px; }
.summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 22px; }
.card { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 15px; }
.label { color: var(--muted); font-size: 12px; margin-bottom: 8px; }
.value { font-size: 26px; font-weight: 700; }
.red { color: var(--red); } .amber { color: var(--amber); } .green { color: var(--green); } .blue { color: var(--blue); }
.section { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 0; margin-bottom: 18px; scroll-margin-top: 18px; }
.section summary { cursor: pointer; list-style: none; padding: 16px 18px; border-bottom: 1px solid transparent; }
.section summary::-webkit-details-marker { display: none; }
.section summary span { font-size: 19px; font-weight: 700; }
.section summary span::before { content: "▸"; color: var(--blue); display: inline-block; margin-right: 8px; transform: translateY(-1px); }
.section[open] summary { border-bottom-color: var(--line); }
.section[open] summary span::before { content: "▾"; }
.section > .note, .section > .table-scroll, .section > .risk-grid, .section > .empty { margin: 14px 18px 18px; }
.note { color: var(--muted); line-height: 1.6; }
.empty { color: var(--green); margin: 0; }
.table-scroll { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; min-width: 760px; }
th, td { border-bottom: 1px solid var(--line); padding: 9px 10px; text-align: left; vertical-align: top; font-size: 13px; }
th { color: var(--muted); font-weight: 600; background: #111a24; position: sticky; top: 0; }
td { color: #e7edf5; }
.risk-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 12px; }
.small { font-size: 12px; color: var(--muted); }
@media (max-width: 980px) {
  .layout { display: block; padding: 18px; }
  .side-nav { position: static; max-height: none; margin-bottom: 18px; }
}
</style>
</head>
<body>
<div class="layout">
$sideNavHtml
<div class="wrap">
  <div class="header">
    <h1>Log Analytics 合并风险报告</h1>
    <div class="meta">
      <span class="tag">查询时间段: $(Escape-Html $AnalysisDate)</span>
      <span class="tag">数据表: $tableCount</span>
      <span class="tag">总记录数: $totalRecords</span>
      <span class="tag">可信 IP 规则: $trustedCount</span>
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
    <div class="card"><div class="label">DCRLogErrors</div><div class="value red">$($riskCounts.DcrLogErrors)</div></div>
    <div class="card"><div class="label">Intune 审计风险</div><div class="value amber">$($riskCounts.IntuneAudit)</div></div>
  </div>

  $reportSectionsHtml

  <p class="small">Generated at $(Escape-Html ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')))</p>
</div>
</div>
</body>
</html>
"@

$utf8Bom = [System.Text.UTF8Encoding]::new($true)
[System.IO.File]::WriteAllText($OutputPath, $html, $utf8Bom)
Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
