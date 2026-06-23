param(
    [Parameter(Mandatory = $true)]
    [string[]]$CsvPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [string]$AnalysisDate,

    [Parameter(Mandatory = $true)]
    [string[]]$TableName,

    [Parameter(Mandatory = $false)]
    [int[]]$TotalCounts,

    [Parameter(Mandatory = $false)]
    [string]$StartUtc = '',

    [Parameter(Mandatory = $false)]
    [string]$EndUtc = '',

    [Parameter(Mandatory = $false)]
    [switch]$SkipLicenseGraph
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $ScriptDir 'log-analyzer-shared.ps1')

if ($CsvPath.Count -ne $TableName.Count) {
    throw 'CsvPath and TableName must have the same number of items.'
}

function Escape-Html {
    param([object]$Text)
    if ($null -eq $Text) { return '' }
    $s = [string]$Text
    return $s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&#39;'
}

function Get-I18nAttr {
    param([string]$Key)
    if ([string]::IsNullOrWhiteSpace($Key)) { return '' }
    return ' data-i18n="' + (Escape-Html $Key) + '"'
}

function Get-I18nKeyFromText {
    param(
        [string]$Prefix,
        [string]$Text
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $slug = [regex]::Replace($Text.Trim(), '[^\p{L}\p{Nd}]+', '_').Trim('_')
    if ([string]::IsNullOrWhiteSpace($slug)) { return '' }
    return "$Prefix.$slug"
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

    $detailFields = @('ResultSignature', 'ResultDescription', 'ResultReason', 'FailureReason', 'Status', 'DeliveryStatus', 'ErrorCode', 'ResultType', 'Subject', 'Message', 'ErrorMessage', 'PermissionName', 'TargetResources', 'ModifiedProperties', 'Result', 'IsSuccess', 'ResultStatus')
    if ($Table -eq 'AuditLogs') {
        # 优先从 ModifiedProperties 中提取权限显示名称
        $modifiedProps = Get-AnyFieldValue -Row $Row -Names @('ModifiedProperties') -Default ''
        if ($modifiedProps) {
            # 尝试提取 appRoleValue (Role 名称，如 User.Read.All)
            if ($modifiedProps -match '"appRoleValue"\s*:\s*"([^"]+)"') {
                $detailFields = @($matches[1])
            }
            # 尝试提取 displayName
            elseif ($modifiedProps -match '"displayName"\s*:\s*"([^"]+)"') {
                $detailFields = @($matches[1])
            }
            # 如果都没有，使用 PermissionName
            else {
                $detailFields = @('PermissionName')
            }
        }
        else {
            $detailFields = @('ResultReason', 'ResultDescription', 'TargetResources', 'Result')
        }
    }

    if ($Table -eq 'DCRLogErrors') {
        $detailFields = @('Message', 'ErrorMessage', 'Details', 'Description', 'Status')
    }
    
    # Intune 审计记录特殊处理
    if ($Table -eq 'IntuneAuditLogsDCR_CL') {
        $intuneResult = Get-AnyFieldValue -Row $Row -Names @('Result', 'ResultStatus', 'Status') -Default ''
        $intuneDesc = Get-AnyFieldValue -Row $Row -Names @('ResultDescription', 'FailureReason', 'Message', 'ErrorMessage') -Default ''
        if ($intuneResult -and $intuneDesc -and $intuneResult -ne $intuneDesc) {
            $detail = "$intuneResult - $intuneDesc"
        } elseif ($intuneDesc) {
            $detail = $intuneDesc
        } elseif ($intuneResult) {
            $detail = $intuneResult
        } else {
            $detail = ''
        }
    } else {
        $detail = Get-AnyFieldValue -Row $Row -Names $detailFields -Default ''
    }
    if ($Table -eq 'MessageTraceDataDCR_CL') {
        $traceStatus = Get-AnyFieldValue -Row $Row -Names @('Status', 'DeliveryStatus', 'EventType', 'Action', 'Result') -Default ''
        $subject = Get-AnyFieldValue -Row $Row -Names @('Subject') -Default ''
        $detail = (@($traceStatus, $subject) | Where-Object { $_ }) -join ' | '
    }
    $target = Get-AnyFieldValue -Row $Row -Names @('Target', 'TargetDeviceName', 'TargetResources', 'TargetResource', 'ObjectId', 'InputStreamId', 'ResourceDisplayName', 'AppDisplayName', 'ServicePrincipalName', 'DisplayName', 'RecipientAddress', 'Subject') -Default ''
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
        [scriptblock]$CellBuilder,
        [string[]]$RawHtmlColumns = @()
    )

    if (-not $Rows -or @($Rows).Count -eq 0) {
        return '<p class="empty" data-i18n="empty.noRisk">未发现相关风险。</p>'
    }

    $html = '<div class="table-scroll"><table><thead><tr>'
    foreach ($col in $Columns) {
        $html += '<th' + (Get-I18nAttr -Key (Get-I18nKeyFromText -Prefix 'field' -Text $col)) + '>' + (Escape-Html $col) + '</th>'
    }
    $html += '</tr></thead><tbody>'
    foreach ($row in $Rows) {
        $html += '<tr>'
        $cells = & $CellBuilder $row
        $cellIndex = 0
        foreach ($cell in $cells) {
            $colName = if ($cellIndex -lt $Columns.Count) { $Columns[$cellIndex] } else { '' }
            if ($colName -in $RawHtmlColumns) {
                $html += '<td>' + $cell + '</td>'
            } else {
                $html += '<td>' + (Escape-Html $cell) + '</td>'
            }
            $cellIndex++
        }
        $html += '</tr>'
    }
    $html += '</tbody></table></div>'
    return $html
}

function New-CodeBlockHtml {
    param([string]$Text)
    return '<details class="kql-block"><summary data-i18n="label.kql">KQL 语句</summary><code>' + (Escape-Html $Text) + '</code></details>'
}

function Get-TimeValueForSort {
    param([string]$Value)
    try { return [DateTime]::Parse($Value) } catch { return [DateTime]::MinValue }
}

function Get-RowTimeValue {
    param([object]$Row)

    $timeText = Get-AnyFieldValue -Row $Row -Names @('LastTime', 'TimeGenerated', 'StartTime', 'FirstTime', 'CreatedDateTime') -Default ''
    return Get-TimeValueForSort -Value $timeText
}

function Get-SuspiciousIpSlidingWindowRows {
    param(
        [object[]]$Rows,
        [int]$WindowDays = 3,
        [int]$Threshold = 10,
        [DateTime]$StartUtc = [DateTime]::MinValue,
        [DateTime]$EndUtc = [DateTime]::MaxValue
    )

    # 计算实际时间范围（天）
    $totalDays = if ($StartUtc -ne [DateTime]::MinValue -and $EndUtc -ne [DateTime]::MaxValue) {
        ($EndUtc - $StartUtc).TotalDays
    } else { 0 }

    # 如果时间范围不足3天，使用每天5次的阈值
    $effectiveThreshold = $Threshold
    $useDailyThreshold = $false
    if ($totalDays -gt 0 -and $totalDays -lt $WindowDays) {
        $effectiveThreshold = 5
        $useDailyThreshold = $true
    }

    $result = [System.Collections.Generic.List[object]]::new()
    foreach ($group in @($Rows | Group-Object -Property IP)) {
        $events = @($group.Group | Where-Object { $_.TimeValue -ne [DateTime]::MinValue } | Sort-Object TimeValue)
        if ($events.Count -eq 0) { continue }

        $maxWindowCount = 0

        if ($useDailyThreshold) {
            # 时间范围不足3天时，按天统计，判断每天是否多于5次
            $dailyCounts = @{}
            foreach ($event in $events) {
                $dayKey = $event.TimeValue.ToString('yyyy-MM-dd')
                if (-not $dailyCounts.ContainsKey($dayKey)) {
                    $dailyCounts[$dayKey] = 0
                }
                $dailyCounts[$dayKey] += [int]$event.Count
            }
            # 找出最大的单日登录次数
            foreach ($count in $dailyCounts.Values) {
                if ($count -gt $maxWindowCount) { $maxWindowCount = $count }
            }
        } else {
            # 时间范围>=3天时，使用滑动窗口算法
            $left = 0
            $windowCount = 0
            for ($right = 0; $right -lt $events.Count; $right++) {
                $windowCount += [int]$events[$right].Count
                $windowEnd = $events[$right].TimeValue
                while ($left -le $right -and $events[$left].TimeValue -lt $windowEnd.AddDays(-$WindowDays)) {
                    $windowCount -= [int]$events[$left].Count
                    $left++
                }
                if ($windowCount -gt $maxWindowCount) { $maxWindowCount = $windowCount }
            }
        }

        if ($maxWindowCount -ge $effectiveThreshold) {
            # 获取该IP的首次和最近访问时间
            $firstAccess = ($events | Sort-Object TimeValue | Select-Object -First 1).TimeValue
            $lastAccess = ($events | Sort-Object TimeValue -Descending | Select-Object -First 1).TimeValue
            $firstAccessStr = if ($firstAccess -ne [DateTime]::MinValue) { $firstAccess.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            $lastAccessStr = if ($lastAccess -ne [DateTime]::MinValue) { $lastAccess.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            $result.Add([PSCustomObject]@{ 
                IP = $group.Name
                Count = $maxWindowCount
                FirstAccess = $firstAccessStr
                LastAccess = $lastAccessStr
            }) | Out-Null
        }
    }

    return @($result | Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, IP)
}

function Get-MailboxIdentityKey {
    param([object]$Row)

    $key = Get-AnyFieldValue -Row $Row -Names @('UserPrincipalName', 'MailboxOwnerUPN', 'PrimarySmtpAddress', 'EmailAddress', 'Mail', 'Identity', 'DisplayName') -Default ''
    if ([string]::IsNullOrWhiteSpace($key)) {
        return [guid]::NewGuid().ToString()
    }
    return $key.ToLowerInvariant()
}

function Get-LatestMailboxRows {
    param([object[]]$Rows)

    $latestByMailbox = @{}
    foreach ($row in @($Rows)) {
        $key = Get-MailboxIdentityKey -Row $row
        $timeText = Get-AnyFieldValue -Row $row -Names @('TimeGenerated', 'LastTime', 'FirstTime') -Default ''
        $time = Get-TimeValueForSort -Value $timeText
        if (-not $latestByMailbox.ContainsKey($key) -or $time -ge $latestByMailbox[$key].Time) {
            $latestByMailbox[$key] = [PSCustomObject]@{ Time = $time; Row = $row }
        }
    }
    return @($latestByMailbox.Values | ForEach-Object { $_.Row })
}

function Get-MailboxTypeText {
    param([object]$Row)

    return Get-AnyFieldValue -Row $Row -Names @('RecipientTypeDetails', 'RecipientTypeDetails_s', 'RecipientTypeDetail', 'RecipientTypeDetail_s', 'MailboxRecipientType', 'MailboxRecipientType_s', 'MailboxType', 'MailboxType_s', 'RecipientType', 'RecipientType_s') -Default ''
}

function Test-SharedMailboxRow {
    param([object]$Row)

    $type = Get-MailboxTypeText -Row $Row
    if ($type -match '(?i)shared') { return $true }

    $flag = Get-AnyFieldValue -Row $Row -Names @('IsSharedMailbox', 'IsSharedMailbox_s', 'IsSharedMailBox', 'IsSharedMailBox_s', 'IsShared', 'IsShared_s', 'SharedMailbox', 'SharedMailbox_s', 'SharedMailBox', 'SharedMailBox_s') -Default ''
    if ($flag -match '(?i)^(true|1|yes|y|shared|shared\s*mail\s*box|sharedmailbox)$') { return $true }
    return $false
}

function Get-MailboxDisplayName {
    param([object]$Row)

    # 直接使用 MailboxStatisticsDCR_CL 查询结果中的 DisplayName 字段
    # 查询命令 New-MailboxStatisticsOptimizedQuery 已确保返回 DisplayName 字段
    $displayName = Get-AnyFieldValue -Row $Row -Names @('DisplayName') -Default ''
    
    # 如果 DisplayName 存在，直接返回（如 "Shayne Wang 基恩士"）
    if ($displayName) {
        return $displayName
    }
    
    # 如果 DisplayName 为空，尝试从 DisplayNameMap 中查找（通过 UserPrincipalName）
    $upn = Get-AnyFieldValue -Row $Row -Names @('UserPrincipalName') -Default ''
    if ($upn -and $upn -match '@') {
        $key = $upn.ToLowerInvariant()
        if ($script:DisplayNameMap.ContainsKey($key)) {
            return $script:DisplayNameMap[$key]
        }
    }
    
    return ''
}

function Get-MailboxEmailAddress {
    param([object]$Row)

    # 直接使用 MailboxStatisticsDCR_CL 查询结果中的 UserPrincipalName 字段
    # 查询命令 New-MailboxStatisticsOptimizedQuery 已确保返回 UserPrincipalName 字段
    $upn = Get-AnyFieldValue -Row $Row -Names @('UserPrincipalName') -Default ''
    
    if ($upn) { return $upn }
    
    # 如果没有 UserPrincipalName，尝试其他邮箱字段
    $primaryEmail = Get-AnyFieldValue -Row $Row -Names @('EmailAddress', 'PrimarySmtpAddress', 'Mail', 'WindowsEmailAddress', 'ExternalEmailAddress', 'MailboxOwnerUPN') -Default ''
    if ($upn) { return $upn }
    
    return ''
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

function Format-CompactTargetForReport {
    param([string]$Target)

    if ([string]::IsNullOrWhiteSpace($Target)) { return '' }
    $text = $Target.Trim()
    $values = [System.Collections.Generic.List[string]]::new()

    try {
        $jsonItems = @($text | ConvertFrom-Json -ErrorAction Stop)
        foreach ($item in $jsonItems) {
            $name = Get-AnyFieldValue -Row $item -Names @('displayName', 'userPrincipalName', 'appDisplayName', 'appId', 'id', 'type') -Default ''
            if ($name) { $values.Add($name) | Out-Null }
        }
    } catch {
        foreach ($fieldName in @('displayName', 'userPrincipalName', 'appDisplayName', 'appId', 'id')) {
            foreach ($match in [regex]::Matches($text, '"' + $fieldName + '"\s*:\s*"([^"]+)"')) {
                if ($match.Groups[1].Value) { $values.Add($match.Groups[1].Value) | Out-Null }
            }
        }
    }

    if ($values.Count -gt 0) {
        return Get-ShortListText -Values @($values.ToArray()) -MaxItems 2
    }
    if ($text.Length -le 50) { return $text }
    return $text.Substring(0, 50) + '...'
}

function Format-DeleteTargetForReport {
    <#
    .SYNOPSIS
        根据操作类型智能格式化被删除者信息，显示清晰的删除对象
    .DESCRIPTION
        针对不同的删除操作类型，提取并格式化被删除者的信息：
        - Delete user: 显示 "用户: xxx"
        - Delete device / Unregister device: 显示 "设备: xxx"
        - Delete application: 显示 "应用: xxx"
        - Remove/Hard delete service principal: 显示 "服务主体: xxx"
        
        TargetResources JSON 结构示例：
        [{"id":"xxx","displayName":"Is Hard Deleted","userType":"Member","accountEnabled":true,"groupType":null,"userPrincipalName":"user@keyence.com.cn","mimeType":"metadata","type":"User"}]
        
        注意：在删除操作中，displayName 字段通常是操作状态（如"Is Hard Deleted"），不是用户名称
    #>
    param(
        [string]$Target,
        [string]$Operation
    )

    if ([string]::IsNullOrWhiteSpace($Target)) { return '' }
    $text = $Target.Trim()
    
    # 根据操作类型确定要提取的字段和显示前缀
    $opLower = if ($Operation) { $Operation.ToLowerInvariant() } else { '' }
    
    $isUserDelete = $opLower -match 'delete\s+user|remove\s+user'
    $isDeviceDelete = $opLower -match 'delete\s+device|unregister\s+device|remove\s+device'
    $isAppDelete = $opLower -match 'delete\s+application|remove\s+application'
    $isSpDelete = $opLower -match '(hard\s+delete|remove)\s+service\s+principal|delete\s+service\s+principal'
    
    # 确定显示前缀
    $prefix = ''
    if ($isUserDelete) { $prefix = '用户: ' }
    elseif ($isDeviceDelete) { $prefix = '设备: ' }
    elseif ($isAppDelete) { $prefix = '应用: ' }
    elseif ($isSpDelete) { $prefix = '服务主体: ' }
    else { $prefix = '目标: ' }
    
    $displayNames = [System.Collections.Generic.List[string]]::new()
    
    # 无效值列表（displayName 字段可能包含的操作状态）
    $invalidDisplayNames = @('Is Hard Deleted', 'Is Deleted', 'Deleted', 'Hard Deleted', 'Remove', 'Removed')
    
    # 根据操作类型优先提取不同字段
    if ($isUserDelete) {
        # 用户删除：优先 userPrincipalName（邮箱格式），其次 id
        # 注意：displayName 在删除操作中通常是操作状态（如"Is Hard Deleted"），不是用户名称
        $fieldPriority = @('userPrincipalName', 'id')
    } elseif ($isDeviceDelete) {
        # 设备删除：优先 displayName, deviceId
        $fieldPriority = @('displayName', 'deviceId', 'devicePhysicalIds', 'id')
    } elseif ($isAppDelete) {
        # 应用删除：优先 displayName, appDisplayName, appId
        $fieldPriority = @('displayName', 'appDisplayName', 'appId', 'id')
    } elseif ($isSpDelete) {
        # 服务主体删除：优先 displayName, servicePrincipalName, appDisplayName
        $fieldPriority = @('displayName', 'servicePrincipalName', 'appDisplayName', 'appId', 'id')
    } else {
        # 其他删除：通用字段
        $fieldPriority = @('displayName', 'userPrincipalName', 'appDisplayName', 'servicePrincipalName', 'appId', 'id')
    }

    # 尝试解析 JSON
    try {
        $jsonItems = @($text | ConvertFrom-Json -ErrorAction Stop)
        foreach ($item in $jsonItems) {
            $displayName = ''
            foreach ($fieldName in $fieldPriority) {
                $value = Get-AnyFieldValue -Row $item -Names @($fieldName) -Default ''
                if ($value) {
                    # 对于用户删除，跳过无效的 displayName 值
                    if ($isUserDelete -and $fieldName -eq 'displayName' -and $value -in $invalidDisplayNames) {
                        continue
                    }
                    $displayName = $value
                    break
                }
            }
            if (-not $displayName) {
                # JSON 解析失败，使用正则提取
                foreach ($fieldName in $fieldPriority) {
                    $matchesArr = [regex]::Matches($text, '"' + $fieldName + '"\s*:\s*"([^"]+)"')
                    if ($matchesArr.Count -gt 0) {
                        $displayName = $matchesArr[0].Groups[1].Value
                        # 对于用户删除，跳过无效的 displayName 值
                        if ($isUserDelete -and $fieldName -eq 'displayName' -and $displayName -in $invalidDisplayNames) {
                            $displayName = ''
                            continue
                        }
                        break
                    }
                }
            }
            if ($displayName) {
                $displayNames.Add($displayName) | Out-Null
            }
        }
    } catch {
        # JSON 解析失败，使用正则提取所有匹配
        foreach ($fieldName in $fieldPriority) {
            $matchesArr = [regex]::Matches($text, '"' + $fieldName + '"\s*:\s*"([^"]+)"')
            foreach ($match in $matchesArr) {
                if ($match.Groups[1].Value) {
                    $value = $match.Groups[1].Value
                    # 对于用户删除，跳过无效的 displayName 值
                    if ($isUserDelete -and $fieldName -eq 'displayName' -and $value -in $invalidDisplayNames) {
                        continue
                    }
                    $displayNames.Add($value) | Out-Null
                }
            }
        }
    }

    if ($displayNames.Count -gt 0) {
        $uniqueNames = @($displayNames | Sort-Object -Unique)
        # 对显示名称进行后处理，使输出更简洁易读
        $formattedNames = @($uniqueNames | ForEach-Object {
            Format-DeleteTargetDisplayName -DisplayName $_ -Operation $opLower
        })
        if ($formattedNames.Count -le 3) {
            return $prefix + ($formattedNames -join ', ')
        } else {
            return $prefix + (($formattedNames | Select-Object -First 3) -join ', ') + " 等$($formattedNames.Count)个"
        }
    }
    
    # 如果没有提取到任何值，返回截断的原始文本
    if ($text.Length -le 80) { return $prefix + $text }
    return $prefix + $text.Substring(0, 80) + '...'
}

function Format-DeleteTargetDisplayName {
    <#
    .SYNOPSIS
        格式化被删除者的显示名称，使输出更简洁易读
    .DESCRIPTION
        对从 TargetResources 提取的显示名称进行后处理：
        - 对于用户：优先显示标准邮箱格式的用户名（如 zhang.san@keyence.com.cn）
        - 对于 OID 格式（如 777ad59df6214f0fa67f333438bc9e74TC202021@china.keyence.com.cn），尝试提取 TC 后面的用户ID
        - 对于设备/应用/服务主体：直接显示 displayName
    #>
    param(
        [string]$DisplayName,
        [string]$Operation
    )

    if ([string]::IsNullOrWhiteSpace($DisplayName)) { return '' }
    
    $opLower = if ($Operation) { $Operation.ToLowerInvariant() } else { '' }
    $isUserDelete = $opLower -match 'delete\s+user|remove\s+user'
    
    # 对于用户删除，尝试解析各种格式
    if ($isUserDelete) {
        # OID 格式：777ad59df6214f0fa67f333438bc9e74TC202021@china.keyence.com.cn
        # 提取 TC 后面的用户ID部分，如 TC202021
        if ($DisplayName -match '[0-9a-f]{32}([A-Z]{2}\d+)@') {
            $userId = $matches[1]
            return "用户: $userId"
        }
        
        # 如果是标准邮箱格式（如 zhang.san@keyence.com.cn），直接显示
        if ($DisplayName -match '@') {
            return $DisplayName
        }
        
        # 如果是 GUID 格式（如 732c874b-0497-464e-a11b-55715833f291），显示为 "用户 (GUID)"
        if ($DisplayName -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
            return "用户 (GUID: $($DisplayName.Substring(0, 8))...)"
        }
        
        # 其他情况，直接显示
        return $DisplayName
    }
    
    # 对于非用户删除，直接返回 displayName
    return $DisplayName
}

function Format-CompactTextForReport {
    param(
        [string]$Text,
        [int]$MaxLength = 80
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $compact = [regex]::Replace($Text.Trim(), '\s+', ' ')
    if ($compact.Length -le $MaxLength) { return $compact }
    return $compact.Substring(0, $MaxLength) + '...'
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

function Get-StrictEventMergeKey {
    param([object]$Row)

    $time = if ($Row.LastTime) { $Row.LastTime } elseif ($Row.Time) { $Row.Time } else { $Row.ActivityDateTime }
    $operationContent = @(
        $Row.Operation,
        $Row.Target,
        $Row.PermissionName,
        $Row.Detail,
        $Row.IP,
        $Row.Status,
        $Row.Reason,
        $Row.Table
    ) -join '|'

    return "$($Row.User)|$operationContent|$time"
}

function Get-DeleteDisableMergeKey {
    param([object]$Row)

    # 删除/Disable 操作栏只根据 最后时间、表、操作者、操作、结果/说明 这5个字段来判断是否合并
    $time = if ($Row.LastTime) { $Row.LastTime } elseif ($Row.Time) { $Row.Time } else { $Row.ActivityDateTime }
    return "$($Row.User)|$($Row.Table)|$($Row.Operation)|$($Row.Detail)|$time"
}

function Test-CachedPrivateOrInvalidIp {
    param([string]$IP)

    if ([string]::IsNullOrWhiteSpace($IP)) { return $true }
    if (-not $script:PrivateIpCache.ContainsKey($IP)) {
        $script:PrivateIpCache[$IP] = Test-PrivateOrInvalidIp -IP $IP
    }
    return [bool]$script:PrivateIpCache[$IP]
}

function Test-CachedTrustedIp {
    param(
        [string]$IP,
        [object[]]$Rules
    )

    if ([string]::IsNullOrWhiteSpace($IP)) { return $false }
    if (-not $script:TrustedIpCache.ContainsKey($IP)) {
        $script:TrustedIpCache[$IP] = Test-IpInTrustedRules -IP $IP -Rules $Rules
    }
    return [bool]$script:TrustedIpCache[$IP]
}

function Format-UserForReport {
    param([string]$User)

    if ([string]::IsNullOrWhiteSpace($User)) { return 'Unknown' }
    
    # 处理已经是 "displayName / email" 或 "email / something" 格式的情况
    if ($User -match '\s/\s') {
        $parts = $User -split '\s/\s', 2
        $leftPart = $parts[0].Trim()
        $rightPart = if ($parts.Count -gt 1) { $parts[1].Trim() } else { '' }
        
        # 确定哪部分是邮箱，哪部分是显示名称
        $emailPart = ''
        $displayPart = ''
        
        if ($leftPart -match '@') {
            $emailPart = $leftPart
            $displayPart = $rightPart
        } elseif ($rightPart -match '@') {
            $emailPart = $rightPart
            $displayPart = $leftPart
        } else {
            # 两部分都不是邮箱格式，直接返回原值
            return $User.Trim()
        }
        
        # 从邮箱中提取 key 来查找 DisplayName
        $emailKey = $emailPart.ToLowerInvariant()
        $displayNameFromMap = if ($script:DisplayNameMap.ContainsKey($emailKey)) { $script:DisplayNameMap[$emailKey] } else { '' }
        
        # 如果显示名称为空或看起来像邮箱地址（包含@），使用 DisplayNameMap 中的值
        if ([string]::IsNullOrWhiteSpace($displayPart) -or $displayPart -match '@') {
            if ($displayNameFromMap) {
                $displayPart = $displayNameFromMap
            } else {
                # 如果 DisplayNameMap 中没有找到，从邮箱地址中提取用户名部分作为显示名称
                # 例如：c250126@china.keyence.com.cn -> c250126
                $usernamePart = $emailPart -split '@' | Select-Object -First 1
                if ($usernamePart) {
                    $displayPart = $usernamePart
                }
            }
        }
        
        # 检查是否有第二种邮箱格式
        $secondaryEmail = ''
        if ($script:UserEmailMap.ContainsKey($emailKey)) {
            $secondaryEmail = $script:UserEmailMap[$emailKey].Secondary
        }
        
        # 构建显示格式
        if ($displayPart -and $secondaryEmail) {
            return "$displayPart / $emailPart ($secondaryEmail)"
        } elseif ($displayPart) {
            return "$displayPart / $emailPart"
        } elseif ($secondaryEmail) {
            return "$emailPart ($secondaryEmail)"
        }
        return $emailPart
    }
    
    # 处理纯邮箱地址或 "DisplayName" 格式的情况
    $emailPattern = '[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}'
    return [regex]::Replace($User, $emailPattern, {
        param($match)
        $email = $match.Value
        $key = $email.ToLowerInvariant()
        $displayName = if ($script:DisplayNameMap.ContainsKey($key)) { $script:DisplayNameMap[$key] } else { '' }
        
        # 如果 DisplayNameMap 中没有找到，从邮箱地址中提取用户名部分作为显示名称
        if (-not $displayName) {
            $usernamePart = $email -split '@' | Select-Object -First 1
            if ($usernamePart) {
                $displayName = $usernamePart
            }
        }
        
        # 检查是否有第二种邮箱格式
        $secondaryEmail = ''
        if ($script:UserEmailMap.ContainsKey($key)) {
            $secondaryEmail = $script:UserEmailMap[$key].Secondary
        }
        # 构建显示格式: DisplayName / PrimaryEmail (SecondaryEmail)
        if ($displayName -and $secondaryEmail) {
            return "$displayName / $email ($secondaryEmail)"
        } elseif ($displayName) {
            return "$displayName / $email"
        } elseif ($secondaryEmail) {
            return "$email ($secondaryEmail)"
        }
        return $email
    }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Add-DisplayNameFromRow {
    param([object]$Row)

    $displayName = Get-AnyFieldValue -Row $Row -Names @('displayName', 'DisplayName', 'UserDisplayName') -Default ''
    $isDisplayNameEmail = $displayName -match '@'
    
    $upn = Get-AnyFieldValue -Row $Row -Names @('userPrincipalName', 'UserPrincipalName', 'UserUPN', 'UPN') -Default ''
    $mail = Get-AnyFieldValue -Row $Row -Names @('mail', 'Mail', 'EmailAddress') -Default ''
    
    # 记录 DisplayName 映射
    if ($displayName) {
        if ($isDisplayNameEmail) {
            # 如果 displayName 是邮箱地址格式，从邮箱中提取用户名部分作为显示名称
            # 例如：c250126@china.keyence.com.cn -> c250126
            $emailKey = $displayName.ToLowerInvariant()
            $usernamePart = $displayName -split '@' | Select-Object -First 1
            # 只有当用户名部分不是纯数字时，才用作显示名称（如 c250126 可以用，但 12345 不行）
            if ($usernamePart -match '[a-zA-Z]') {
                foreach ($id in @($upn, $mail)) {
                    if ($id -and $id -match '@') {
                        $script:DisplayNameMap[$id.ToLowerInvariant()] = $usernamePart
                    }
                }
            }
        } else {
            # displayName 不是邮箱地址，直接记录映射
            foreach ($id in @($upn, $mail)) {
                if ($id -and $id -match '@') {
                    $script:DisplayNameMap[$id.ToLowerInvariant()] = $displayName
                }
            }
        }
    }
    
    # 记录两种邮箱格式的映射关系（UPN <-> Mail）
    if ($upn -and $upn -match '@' -and $mail -and $mail -match '@' -and $upn -ne $mail) {
        $upnKey = $upn.ToLowerInvariant()
        $mailKey = $mail.ToLowerInvariant()
        # 以 UPN 为主键，记录 Secondary（Mail）
        if (-not $script:UserEmailMap.ContainsKey($upnKey)) {
            $script:UserEmailMap[$upnKey] = [PSCustomObject]@{ Primary = $upn; Secondary = $mail }
        }
        # 以 Mail 为主键，记录 Secondary（UPN）
        if (-not $script:UserEmailMap.ContainsKey($mailKey)) {
            $script:UserEmailMap[$mailKey] = [PSCustomObject]@{ Primary = $mail; Secondary = $upn }
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
    return Get-AnyFieldValue -Row $Row -Names @('AppDisplayName', 'ServicePrincipalName', 'Application', 'ApplicationDisplayName', 'ClientAppUsed') -Default ''
}

function Test-AllowedSigninApp {
    param([string]$AppName)
    if ([string]::IsNullOrWhiteSpace($AppName)) { return $false }
    $normalizedAppName = $AppName.Trim()
    return (@('Windows Sign In', 'Microsoft Edge', 'Microsoft Office') -contains $normalizedAppName) -or ($normalizedAppName -match '(?i)Sangfor')
}

function Test-SangforRelatedSigninRow {
    param(
        [object]$Row,
        [object]$EventRecord = $null
    )

    $text = @(
        (Get-AnyFieldValue -Row $Row -Names @('UserPrincipalName', 'UserDisplayName', 'Identity', 'AppDisplayName', 'ServicePrincipalName', 'Application', 'ApplicationDisplayName', 'ClientAppUsed', 'ManagedIdentityName') -Default ''),
        $(if ($null -ne $EventRecord) { $EventRecord.User } else { '' }),
        $(if ($null -ne $EventRecord) { $EventRecord.Operation } else { '' }),
        $(if ($null -ne $EventRecord) { $EventRecord.Target } else { '' }),
        $(if ($null -ne $EventRecord) { $EventRecord.Detail } else { '' })
    ) -join ' '

    return ($text -match '(?i)Sangfor')
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

function Get-WebExceptionResponseText {
    param([object]$ErrorRecord)

    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        return [string]$ErrorRecord.ErrorDetails.Message
    }
    try {
        $response = $ErrorRecord.Exception.Response
        if ($null -eq $response) { return '' }
        $stream = $response.GetResponseStream()
        if ($null -eq $stream) { return '' }
        $reader = [System.IO.StreamReader]::new($stream)
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } catch {
        return ''
    }
}

function Get-OAuthErrorCode {
    param([string]$Body)
    if ([string]::IsNullOrWhiteSpace($Body)) { return '' }
    try {
        $json = $Body | ConvertFrom-Json
        if ($json.error) { return [string]$json.error }
    } catch {
    }
    if ($Body -match 'authorization_pending') { return 'authorization_pending' }
    if ($Body -match 'slow_down') { return 'slow_down' }
    if ($Body -match 'authorization_declined') { return 'authorization_declined' }
    if ($Body -match 'expired_token') { return 'expired_token' }
    return ''
}

function Confirm-LicenseGraphVerification {
    <#
    .SYNOPSIS
        确认是否继续执行 Microsoft Graph License API 验证
    .DESCRIPTION
        默认直接继续验证，不再要求用户确认。
        如果需要在非交互环境中跳过验证，请在调用 main.ps1 时添加 -SkipLicenseGraph 参数。
    .EXAMPLE
        Confirm-LicenseGraphVerification
    #>
    # 默认直接继续验证，不再要求用户确认
    # 如需跳过验证，请使用 main.ps1 的 -SkipLicenseGraph 参数
    return $true
}

function Get-LicenseFromMgGraph {
    <#
    .SYNOPSIS
        使用 Connect-MgGraph + Get-MgSubscribedSku 获取 License 数据
    .DESCRIPTION
        优先使用 Microsoft Graph PowerShell SDK（中国版 21v）获取订阅的 SKU 信息
    .EXAMPLE
        Get-LicenseFromMgGraph
    #>
    param(
        [string]$ClientId = '5bbea6de-1297-488f-aff5-9b55f4c61c3e',
        [string]$TenantId = '420c4dab-8603-402f-afe0-75bc28c51c13'
    )

    if (-not (Confirm-LicenseGraphVerification)) {
        throw '用户选择跳过 Microsoft Graph License API 验证。'
    }

    # 检查 Microsoft.Graph.Authentication 模块
    $mgAuthModule = Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication'
    
    if (-not $mgAuthModule) {
        Write-Host 'Microsoft.Graph.Authentication 模块未找到，尝试安装...' -ForegroundColor Yellow
        try {
            Install-Module -Name 'Microsoft.Graph.Authentication' -Scope CurrentUser -Force -AllowClobber -Confirm:$false -ErrorAction Stop
            Write-Host 'Microsoft.Graph.Authentication 模块安装成功。' -ForegroundColor Green
        } catch {
            throw "无法安装 Microsoft.Graph.Authentication 模块：$($_.Exception.Message)`n请手动运行: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force"
        }
        $mgAuthModule = Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication'
    }
    
    if (-not $mgAuthModule) {
        throw 'Microsoft.Graph.Authentication PowerShell 模块不可用。请运行: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }

    # 导入 Microsoft.Graph 模块
    Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop

    try {
        Write-Host ''
        Write-Host '=== 开始通过 Microsoft Graph PowerShell SDK (China 21v) 获取 License 数据 ===' -ForegroundColor Cyan
        Write-Host "ClientId: $ClientId" -ForegroundColor Cyan
        Write-Host "TenantId: $TenantId" -ForegroundColor Cyan
        Write-Host '使用命令: Connect-MgGraph -Environment China -ClientId "..."' -ForegroundColor Cyan

        # 使用 Connect-MgGraph 连接中国版 21v
        Connect-MgGraph -Environment China -ClientId $ClientId -TenantId $TenantId -NoWelcome -ErrorAction Stop | Out-Null
        Write-Host 'Microsoft Graph 连接成功！' -ForegroundColor Green

        # 获取 subscribedSkus
        Write-Host '正在获取 License 数据...' -ForegroundColor Cyan
        
        $allSkus = @(Get-MgSubscribedSku -ErrorAction Stop)

        if ($allSkus.Count -eq 0) {
            throw 'Microsoft.Graph Get-MgSubscribedSku 未返回任何许可证 SKU。'
        }

        # 断开连接
        Disconnect-MgGraph -ErrorAction SilentlyContinue 2>$null

        if ($allSkus.Count -eq 0) {
            throw 'Microsoft.Graph Get-MgSubscribedSku 未返回任何许可证 SKU。'
        }

        Write-Host "成功获取 $($allSkus.Count) 个 License SKU。" -ForegroundColor Green

        # 输出所有获取到的 License 详细信息
        Write-Host ''
        Write-Host '=== 获取到的 License 详情 ===' -ForegroundColor Cyan
        foreach ($sku in $allSkus) {
            $name = [string]$sku.SkuPartNumber
            $enabledVal = if ($null -ne $sku.PrepaidUnits -and $null -ne $sku.PrepaidUnits.Enabled) { [int]$sku.PrepaidUnits.Enabled } else { 0 }
            $consumedVal = if ($null -ne $sku.ConsumedUnits) { [int]$sku.ConsumedUnits } else { 0 }
            Write-Host "  ${name}: Enabled=${enabledVal}, Consumed=${consumedVal}" -ForegroundColor Gray
        }
        Write-Host '=============================' -ForegroundColor Cyan

        # 构建 SKU 列表（按照用户指定的格式）
        $skuList = [System.Collections.Generic.List[object]]::new()
        foreach ($sku in $allSkus) {
            $name = [string]$sku.SkuPartNumber
            if ([string]::IsNullOrWhiteSpace($name)) { continue }

            # 按照用户指定的格式计算：
            # SkuPartNumber = 许可证名
            # Total = PrepaidUnits.Enabled (已购买且处于活跃状态的许可证总数)
            # ConsumedUnits = 已分配给用户的许可证数量
            # Suspended = 已挂起的许可证数量
            # Warning = 处于宽限期或过期警告状态的许可证数量
            # Remaining = PrepaidUnits.Enabled - ConsumedUnits - Warning - Suspended
            $enabled = if ($null -ne $sku.PrepaidUnits -and $null -ne $sku.PrepaidUnits.Enabled) { [int]$sku.PrepaidUnits.Enabled } else { 0 }
            $warning = if ($null -ne $sku.PrepaidUnits -and $null -ne $sku.PrepaidUnits.Warning) { [int]$sku.PrepaidUnits.Warning } else { 0 }
            $suspended = if ($null -ne $sku.PrepaidUnits -and $null -ne $sku.PrepaidUnits.Suspended) { [int]$sku.PrepaidUnits.Suspended } else { 0 }
            $total = $enabled
            $consumed = if ($null -ne $sku.ConsumedUnits) { [int]$sku.ConsumedUnits } else { 0 }
            $remaining = $enabled - $consumed - $warning - $suspended
            if ($remaining -lt 0) { $remaining = 0 }

            Write-Host "  License: $name | Total: $total | Consumed: $consumed | Warning: $warning | Suspended: $suspended | Remaining: $remaining" -ForegroundColor Gray

            $skuList.Add([PSCustomObject]@{
                License = $name
                Total = $total
                Enabled = $enabled
                ConsumedUnits = $consumed
                Warning = $warning
                Suspended = $suspended
                Remaining = $remaining
            }) | Out-Null
        }

        if ($skuList.Count -eq 0) {
            throw '未找到有效的 License SKU。'
        }

        return [PSCustomObject]@{
            Success = $true
            Message = "License 总量已通过 Microsoft.Graph PowerShell SDK (Connect-MgGraph -Environment China) Get-MgSubscribedSku 获取。"
            SkuList = $skuList.ToArray()
        }
    } catch {
        # 断开可能存在的连接
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
        throw "无法通过 Microsoft.Graph PowerShell SDK 获取 License 信息：$($_.Exception.Message)"
    }
}

function Get-GraphDeviceCodeTokenCandidate {
    param(
        [string]$TenantId = '420c4dab-8603-402f-afe0-75bc28c51c13'
    )

    if (-not (Confirm-LicenseGraphVerification)) {
        throw '用户选择跳过 Microsoft Graph License API 验证。'
    }

    $clientIds = @(
        '14d82eec-204b-4c2f-b7e8-296a70dab67e',
        '1950a258-227b-4e31-a9cf-717495945fc2'
    )
    $loginBase = 'https://login.chinacloudapi.cn'
    $graphScope = 'https://microsoftgraph.chinacloudapi.cn/Organization.Read.All offline_access'
    $endpoint = 'https://microsoftgraph.chinacloudapi.cn/v1.0/subscribedSkus'
    $tenant = if ([string]::IsNullOrWhiteSpace($TenantId)) { 'common' } else { $TenantId }
    $deviceCodeUri = "$loginBase/$tenant/oauth2/v2.0/devicecode"
    $tokenUri = "$loginBase/$tenant/oauth2/v2.0/token"

    $device = $null
    $clientId = ''
    $deviceErrors = [System.Collections.Generic.List[string]]::new()
    foreach ($candidateClientId in $clientIds) {
        try {
            $device = Invoke-RestMethod -Method Post -Uri $deviceCodeUri -Body @{
                client_id = $candidateClientId
                scope = $graphScope
            } -ErrorAction Stop
            $clientId = $candidateClientId
            break
        } catch {
            $deviceErrors.Add("$candidateClientId`: $($_.Exception.Message)") | Out-Null
        }
    }
    if ($null -eq $device) {
        throw "无法启动 Microsoft Graph device code 登录。$($deviceErrors -join '；')"
    }

    Write-Host ''
    Write-Host '=== Microsoft Graph License API Login Required ===' -ForegroundColor Yellow
    Write-Host $device.message -ForegroundColor Cyan

    $deadline = (Get-Date).AddSeconds([int]$device.expires_in)
    $interval = [Math]::Max(5, [int]$device.interval)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUri -Body @{
                grant_type = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id = $clientId
                device_code = $device.device_code
            } -ErrorAction Stop
            if ($tokenResponse.access_token) {
                return [PSCustomObject]@{ Source = 'DeviceCode:ChinaGraph'; Endpoint = $endpoint; Token = [string]$tokenResponse.access_token }
            }
        } catch {
            $body = Get-WebExceptionResponseText -ErrorRecord $_
            $oauthError = Get-OAuthErrorCode -Body $body
            if ($oauthError -eq 'authorization_pending') { continue }
            if ($oauthError -eq 'slow_down') { $interval += 5; continue }
            if ($oauthError -eq 'authorization_declined') { throw '用户拒绝了 Microsoft Graph device code 登录。' }
            if ($oauthError -eq 'expired_token') { throw 'Microsoft Graph device code 已过期，请重新运行脚本获取新的 code。' }
            if ($body) { throw "Microsoft Graph device code token 请求失败：$body" }
            throw
        }
    }
    throw 'Device code 登录超时。'
}

function Get-GraphAccessTokenCandidates {
    $candidates = [System.Collections.Generic.List[object]]::new()
    $errors = [System.Collections.Generic.List[string]]::new()
    $resources = @(
        [PSCustomObject]@{ Name = 'ChinaGraph'; ResourceUrl = 'https://microsoftgraph.chinacloudapi.cn/'; Endpoint = 'https://microsoftgraph.chinacloudapi.cn/v1.0/subscribedSkus' },
        [PSCustomObject]@{ Name = 'GlobalGraph'; ResourceUrl = 'https://graph.microsoft.com/'; Endpoint = 'https://graph.microsoft.com/v1.0/subscribedSkus' }
    )

    if (-not (Get-Command Get-AzAccessToken -ErrorAction SilentlyContinue)) {
        Import-Module Az.Accounts -Force -ErrorAction SilentlyContinue
    }
    if (Get-Command Get-AzAccessToken -ErrorAction SilentlyContinue) {
        foreach ($resource in $resources) {
            try {
                $tokenResponse = Get-AzAccessToken -ResourceUrl $resource.ResourceUrl -ErrorAction Stop
                $token = ConvertFrom-AccessTokenValue -Token $tokenResponse.Token
                if (-not [string]::IsNullOrWhiteSpace($token)) {
                    $candidates.Add([PSCustomObject]@{ Source = "Az:$($resource.Name)"; Endpoint = $resource.Endpoint; Token = $token }) | Out-Null
                }
            } catch {
                $errors.Add("Az $($resource.Name): $($_.Exception.Message)") | Out-Null
            }
        }
    } else {
        $errors.Add('Az.Accounts / Get-AzAccessToken 不可用') | Out-Null
    }

    if (Get-Command az -ErrorAction SilentlyContinue) {
        foreach ($resource in $resources) {
            try {
                $jsonText = & az account get-access-token --resource $resource.ResourceUrl --output json 2>$null
                if ($LASTEXITCODE -eq 0 -and $jsonText) {
                    $tokenResponse = ($jsonText -join "`n") | ConvertFrom-Json
                    if ($tokenResponse.accessToken) {
                        $candidates.Add([PSCustomObject]@{ Source = "AzureCLI:$($resource.Name)"; Endpoint = $resource.Endpoint; Token = [string]$tokenResponse.accessToken }) | Out-Null
                    }
                }
            } catch {
                $errors.Add("Azure CLI $($resource.Name): $($_.Exception.Message)") | Out-Null
            }
        }
    } else {
        $errors.Add('Azure CLI az 不可用') | Out-Null
    }

    if ($candidates.Count -eq 0) {
        # 只使用 AuthCodeFlow（中国版 21v），不再回退到 DeviceCodeFlow
        try {
            $candidates.Add((Get-GraphAuthCodeTokenCandidate)) | Out-Null
        } catch {
            $errors.Add("AuthCode ChinaGraph (21v): $($_.Exception.Message)") | Out-Null
        }
    }

    return [PSCustomObject]@{ Candidates = @($candidates.ToArray()); Errors = @($errors.ToArray()) }
}

function Get-LicenseSkuTotalsFromGraph {
    # 只使用 Microsoft Graph PowerShell SDK (Connect-MgGraph + Get-MgSubscribedSku)
    # 不再使用 DeviceCode 或 REST API 回退
    # 使用 AuthCode 认证方式（浏览器登录）
    $result = [PSCustomObject]@{
        Success = $false
        Message = ''
        Skus = @{}
        SkuList = @()
    }

    $ClientId = '5bbea6de-1297-488f-aff5-9b55f4c61c3e'
    $TenantId = '420c4dab-8603-402f-afe0-75bc28c51c13'

    try {
        Write-Host ''
        Write-Host '=== 开始通过 Microsoft Graph (AuthCode) 获取 License 数据 ===' -ForegroundColor Cyan
        Write-Host "ClientId: $ClientId" -ForegroundColor Cyan
        Write-Host "TenantId: $TenantId" -ForegroundColor Cyan
        
        # 直接使用 Connect-MgGraph + Get-MgSubscribedSku 获取 License 数据
        $mgResult = Get-LicenseFromMgGraph -ClientId $ClientId -TenantId $TenantId
        
        if ($mgResult.Success) {
            $map = @{}
            $skuList = [System.Collections.Generic.List[object]]::new()
            foreach ($sku in $mgResult.SkuList) {
                $name = [string]$sku.License
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                $skuRecord = [PSCustomObject]@{
                    License = $name
                    Used = $sku.ConsumedUnits
                    Total = $sku.Total
                    Warning = $sku.Warning
                    Suspended = $sku.Suspended
                    Remaining = $sku.Remaining
                }
                $skuList.Add($skuRecord) | Out-Null
                $map[$name.ToUpperInvariant()] = $skuRecord
                $normalizedName = Normalize-LicenseKey -Name $name
                if ($normalizedName) {
                    $map[$normalizedName] = $skuRecord
                }
            }
            if ($skuList.Count -eq 0) { throw 'Microsoft Graph Get-MgSubscribedSku 未返回任何许可证 SKU。' }
            $result.Success = $true
            $result.Message = $mgResult.Message
            $result.Skus = $map
            $result.SkuList = $skuList.ToArray()
            Write-Host ''
            Write-Host "=== License 数据获取成功，共 $($skuList.Count) 个 License ===" -ForegroundColor Green
            return $result
        }
    } catch {
        Write-Host "无法通过 Microsoft Graph (AuthCode) 获取 License 总量：$($_.Exception.Message)" -ForegroundColor Red
        $result.Message = "无法通过 Microsoft Graph (AuthCode) 获取 License 总量：$($_.Exception.Message)"
        return $result
    }
}

function Test-LicenseMetricMissing {
    param([object]$License)
    if ($null -eq $License) { return $true }
    foreach ($name in @('Used', 'Total', 'Remaining')) {
        $value = $License.$name
        if ($null -eq $value) { return $true }
        if ([string]::IsNullOrWhiteSpace([string]$value)) { return $true }
        if ([string]$value -eq 'N/A') { return $true }
    }
    return $false
}

function Get-RiskLevel {
    param(
        [string]$Category,
        [object]$Row
    )
    
    switch ($Category) {
        'FailedSignins' {
            $count = 0
            if ($Row.PSObject.Properties.Name -contains 'EventCount') { $count = [int]$Row.EventCount }
            elseif ($Row.PSObject.Properties.Name -contains 'Count') { $count = [int]$Row.Count }
            if ($count -gt 50) { return 'high' }
            if ($count -gt 10) { return 'medium' }
            return 'low'
        }
        'SuspiciousIP' {
            $count = 0
            if ($Row.PSObject.Properties.Name -contains 'Count') { $count = [int]$Row.Count }
            if ($count -gt 20) { return 'high' }
            if ($count -gt 5) { return 'medium' }
            return 'low'
        }
        'SuspiciousSuccess' {
            $count = 0
            if ($Row.PSObject.Properties.Name -contains 'Count') { $count = [int]$Row.Count }
            if ($count -gt 10) { return 'high' }
            if ($count -gt 3) { return 'medium' }
            return 'low'
        }
        'DeleteDisable' {
            $op = ''
            if ($Row.PSObject.Properties.Name -contains 'Operation') { $op = $Row.Operation }
            if ($op -match 'delete|remove|disable') { return 'high' }
            return 'medium'
        }
        'IdentityPermission' { return 'high' }
        'MailboxLowSpace' {
            # 根据 CapacityRisk 字段判断风险等级
            $capacityRisk = ''
            if ($Row.PSObject.Properties.Name -contains 'CapacityRisk') { 
                $capacityRisk = [string]$Row.CapacityRisk
            }
            # 当邮箱容量是否风险为"是"时，风险等级为"中"
            # 当邮箱容量是否风险为"否"时，风险等级为"无"
            if ($capacityRisk -eq '是') { return 'medium' }
            return 'none'
        }
        'DcrLogErrors' { return 'medium' }
        'IntuneAudit' { return 'low' }
        default { return 'low' }
    }
}

function Get-RiskLevelBadge {
    param([string]$Level)
    
    $colorMap = @{
        'high' = '#dc3545'
        'medium' = '#ffc107'
        'low' = '#28a745'
        'none' = '#6c757d'
    }
    $textMap = @{
        'high' = '高'
        'medium' = '中'
        'low' = '低'
        'none' = '无'
    }
    $color = $colorMap[$Level]
    if (-not $color) { $color = '#6c757d' }
    $text = $textMap[$Level]
    if (-not $text) { $text = '低' }
    return "<span style='background-color: $color; color: white; padding: 2px 8px; border-radius: 4px; font-size: 12px;'>$text</span>"
}

function Get-AiAnalysis {
    param(
        [string]$Category,
        [object[]]$Data
    )
    
    $count = ($Data | Measure-Object).Count
    switch ($Category) {
        'FailedSignins' {
            $totalEvents = 0
            foreach ($d in $Data) { 
                if ($d.PSObject.Properties.Name -contains 'EventCount') { $totalEvents += [int]$d.EventCount }
                elseif ($d.PSObject.Properties.Name -contains 'Count') { $totalEvents += [int]$d.Count }
            }
            return "检测到 $count 个应用/身份存在登录失败，共计 $totalEvents 次失败事件。主要风险：可能存在暴力破解攻击、凭证泄露或配置错误。建议检查失败原因并采取相应安全措施。"
        }
        'SuspiciousIP' {
            return "检测到 $count 个可疑IP地址尝试访问。主要风险：可能来自恶意来源的未授权访问尝试。建议核实IP来源，必要时添加防火墙规则或加入可信IP列表。"
        }
        'SuspiciousSuccess' {
            return "检测到 $count 个来自非可信位置的成功登录。主要风险：可能是凭证泄露后的异地登录，或用户使用了未授权的网络。建议验证用户身份并确认登录合法性。"
        }
        'DeleteDisable' {
            return "检测到 $count 个删除或禁用操作。主要风险：关键资源被意外或恶意删除/禁用可能导致服务中断。建议审核操作者权限并确认操作合法性。"
        }
        'IdentityPermission' {
            return "检测到 $count 个身份权限变更操作。主要风险：权限提升或不当授权可能导致安全漏洞。建议严格审核权限变更，确保符合最小权限原则。"
        }
        'MailboxLowSpace' {
            return "检测到 $count 个邮箱容量不足。主要风险：邮箱满可能导致邮件丢失或业务中断。建议清理邮箱或增加配额。"
        }
        'DcrLogErrors' {
            return "检测到 $count 个DCR日志错误。主要风险：数据采集规则异常可能导致日志丢失，影响安全监控。建议检查DCR配置和网络连接。"
        }
        'IntuneAudit' {
            return "检测到 $count 个Intune审计记录。主要风险：设备管理变更可能影响终端安全策略。建议审核变更内容确保合规。"
        }
        default {
            return "检测到相关活动，请关注潜在风险。"
        }
    }
}

function New-ReportSection {
    param(
        [string]$Id,
        [string]$Title,
        [string]$Note,
        [string]$Content,
        [bool]$Open = $true,
        [string]$AiAnalysis = ''
    )
    $openText = if ($Open) { ' open' } else { '' }
    $noteHtml = if ([string]::IsNullOrWhiteSpace($Note)) { '' } else { 
        # 先 HTML 转义，再将 \n 转换为 <br> 实现换行
        $escapedNote = Escape-Html $Note -replace '\r?\n', '<br>'
        '<p class="note">' + $escapedNote + '</p>' 
    }
    $aiAnalysisHtml = ''
    if (-not [string]::IsNullOrWhiteSpace($AiAnalysis)) {
        $aiAnalysisHtml = '<div class="ai-analysis"><strong>🤖 AI 分析（主要风险）：</strong>' + (Escape-Html $AiAnalysis) + '</div>'
    }
    $titleKey = "section.$Id.title"
    return @"
  <details class="section" id="$(Escape-Html $Id)"$openText>
    <summary><span data-i18n="$(Escape-Html $titleKey)">$(Escape-Html $Title)</span></summary>
    $aiAnalysisHtml
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
$script:UserEmailMap = @{}  # 存储用户的两种邮箱格式: key=主邮箱, value=PSCustomObject(Primary, Secondary)
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
$script:PrivateIpCache = @{}
$script:TrustedIpCache = @{}
$failedSignins = [System.Collections.Generic.List[object]]::new()
$failedOperations = [System.Collections.Generic.List[object]]::new()
$deleteDisableEvents = [System.Collections.Generic.List[object]]::new()
$suspiciousSigninSuccess = [System.Collections.Generic.List[object]]::new()
$suspiciousIpReasons = @{}
$suspiciousIpRecords = [System.Collections.Generic.List[object]]::new()
$clientIpCounts = @{}
$identityPermissionChanges = [System.Collections.Generic.List[object]]::new()
$dcrLogErrorRows = [System.Collections.Generic.List[object]]::new()
$intuneAuditRows = [System.Collections.Generic.List[object]]::new()

for ($i = 0; $i -lt $datasets.Count; $i++) {
    $dataset = $datasets[$i]
    $table = $dataset.Table
    $filteredCount = $dataset.Rows.Count
    $totalCount = if ($TotalCounts -and $i -lt $TotalCounts.Count) { $TotalCounts[$i] } else { $filteredCount }

    if ($table -in @('AssignedLicensesDCR_CL', 'MailboxStatisticsDCR_CL')) {
        continue
    }

    foreach ($row in $dataset.Rows) {
        if ($table -eq 'AuditLogs') {
            # 根据 OperationName 和 AADOperationType 字段自动分类记录类型
            $operationName = Get-AnyFieldValue -Row $row -Names @('OperationName', 'Operation') -Default ''
            $aadOpType = Get-AnyFieldValue -Row $row -Names @('AADOperationType') -Default ''
            $result = Get-AnyFieldValue -Row $row -Names @('Result') -Default ''
            
            # 权限变更操作名列表（与 KQL 中的 __isPermissionChange 一致）
            $permissionChangeOps = @(
                'Add delegated permission grant',
                'Consent to application',
                'Create application – Certificates and secrets management',
                'Add owner to application',
                'Add app role assignment to service principal',
                'Update application – Certificates and secrets management',
                'Remove delegated permission grant',
                'Remove app role assignment from service principal'
            )
            
            # 判断是否为删除操作：OperationName 包含 delete/remove/disable 关键词，或 AADOperationType 为 Delete
            $normalizedOp = $operationName.ToLowerInvariant()
            $isDeleteOperation = ($aadOpType -eq 'Delete') -or 
                                 ($normalizedOp -match '(^|[^a-z])(delete|deleted|remove|removed|disable|disabled|deactivate|deactivated)([^a-z]|$)')
            
            # 判断是否为权限变更操作：Result 为 success 且 OperationName 在权限变更操作名列表中
            $isPermissionChange = ($result -eq 'success') -and ($operationName -in $permissionChangeOps)
            
            if ($isDeleteOperation) {
                $deleteDisableEvents.Add((New-EventRecord -Table $table -Row $row -Reason '删除操作')) | Out-Null
            } elseif ($isPermissionChange) {
                $identityPermissionChanges.Add((New-EventRecord -Table $table -Row $row -Reason '权限变更审计')) | Out-Null
            }
            # 其他 AuditLogs 记录不添加到任何风险列表，跳过处理
            continue
        }

        $rowEventCount = Get-RowEventCount -Row $row
        $recordKind = Get-AnyFieldValue -Row $row -Names @('__RecordKind') -Default ''
        $op = Get-OperationValue -Row $row -TableName $table
        $success = Get-SuccessValue -Row $row -TableName $table
        if ($recordKind -eq 'AggregatedSuspiciousSigninSuccess') {
            $success = 'true'
        } elseif ($recordKind -eq 'AggregatedFailedSignin') {
            $success = 'false'
        }
        $ip = Get-NormalizedIpValue -IP (Get-ClientIpValue -Row $row -TableName $table)
        $isUsablePublicIp = -not (Test-CachedPrivateOrInvalidIp -IP $ip)
        $isTrustedIp = if ($isUsablePublicIp) { Test-CachedTrustedIp -IP $ip -Rules $trustedRules } else { $false }

        if ($isUsablePublicIp) {
            Add-Count -Map $clientIpCounts -Key $ip -By $rowEventCount
        }

        if ($success -eq 'false') {
            $record = New-EventRecord -Table $table -Row $row -Reason '失败/异常'
            if ($table -in @('AADManagedIdentitySignInLogs', 'AADServicePrincipalSignInLogs', 'SigninLogs')) {
                # 不再在记录级别过滤 keyence.com.cn，改为在分组后过滤（与KQL逻辑一致）
                # 不再在记录级别过滤 EventCount > 10，改为在分组后过滤（与KQL逻辑一致）
                $failedSignins.Add($record) | Out-Null
            } elseif ($table -in @('DCRLogErrors', 'IntuneAuditLogsDCR_CL')) {
                # These tables have dedicated sections below.
            } else {
                $failedOperations.Add($record) | Out-Null
            }
        }

        if ($table -in @('AADManagedIdentitySignInLogs', 'AADServicePrincipalSignInLogs', 'SigninLogs') -and $success -eq 'true' -and $isUsablePublicIp -and -not $isTrustedIp) {
            $appName = Get-SigninAppName -Row $row
            $isAllowedInteractiveApp = ($table -eq 'SigninLogs' -and (Test-AllowedSigninApp -AppName $appName))
            if (-not $isAllowedInteractiveApp -and -not (Test-SangforRelatedSigninRow -Row $row)) {
                $event = New-EventRecord -Table $table -Row $row -Reason "可信位置外成功登录，应用：$appName"
                if (-not (Test-SangforRelatedSigninRow -Row $row -EventRecord $event)) {
                    $suspiciousSigninSuccess.Add($event) | Out-Null
                    if ($table -eq 'SigninLogs') {
                        $suspiciousIpReasons[$ip] = 'SigninLogs 可信位置外成功登录'
                    }
                }
            }
        }

        if ($table -eq 'SigninLogs' -and $isUsablePublicIp -and -not $isTrustedIp) {
            $firstTime = Get-AnyFieldValue -Row $row -Names @('FirstTime', 'StartTime', 'MinTime') -Default ''
            $lastTime = Get-AnyFieldValue -Row $row -Names @('LastTime', 'EndTime', 'MaxTime') -Default ''
            if (-not $firstTime) { $firstTime = [string]$row.TimeGenerated }
            if (-not $lastTime) { $lastTime = [string]$row.TimeGenerated }
            $suspiciousIpRecords.Add([PSCustomObject]@{
                IP = $ip
                Count = $rowEventCount
                TimeValue = Get-RowTimeValue -Row $row
                FirstTime = Get-LocalTimeText -Value $firstTime
                LastTime = Get-LocalTimeText -Value $lastTime
            }) | Out-Null
        }

        if ($table -eq 'DCRLogErrors') {
            $dcrLogErrorRows.Add((New-EventRecord -Table $table -Row $row -Reason 'DCR 日志采集错误')) | Out-Null
        }

        if ($table -eq 'IntuneAuditLogsDCR_CL') {
            $intuneAuditRows.Add((New-EventRecord -Table $table -Row $row -Reason 'Intune 审计风险')) | Out-Null
        }

    }
}

# ==================== License 产品名称映射（必须在使用前定义）====================
$licenseProductMap = @{
    'AAD_PREMIUM_P2_CN' = 'Microsoft Entra ID P2 (中国版)'
    'POWER_BI_PRO' = 'Power BI Pro'
    'SPE_E3_NO_WIN' = 'Office 365 E3 (不含Windows)'
    'EXCHANGEENTERPRISE' = 'Exchange Online Enterprise'
}

# License usage: 优先使用 Microsoft Graph API 获取 License 列表，然后从日志中获取使用量
$licenseStatusNote = 'License 列表优先使用 Microsoft Graph subscribedSkus 获取（SkuPartNumber 更准确）；使用量优先从 AssignedLicensesDCR_CL 日志获取。'

# 首先尝试从 Graph API 获取 License 列表
$graphLicenseResult = if ($SkipLicenseGraph) {
    [PSCustomObject]@{ Success = $false; Message = '已跳过 Microsoft Graph License 校验以加快报告生成。'; Skus = @{}; SkuList = @() }
} else {
    Get-LicenseSkuTotalsFromGraph
}
$licenseUsage = @()

if ($graphLicenseResult.Success) {
    # 从日志中获取使用量映射
    $licenseRows = @($datasets | Where-Object { $_.Table -eq 'AssignedLicensesDCR_CL' } | ForEach-Object { $_.Rows })
    $logUsageMap = @{}
    foreach ($row in $licenseRows) {
        $licenseName = Get-AnyFieldValue -Row $row -Names @('SkuPartNumber', 'LicenseName', 'SkuDisplayName', 'ServicePlanName', 'AssignedLicenses', 'Licenses') -Default 'Unknown License'
        $user = Get-UserValue -Row $row -TableName 'AssignedLicensesDCR_CL'
        $usedUsers = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('UsedUsers', 'Used') -Default '')
        
        $normalizedLogName = Normalize-LicenseKey -Name $licenseName
        if (-not $logUsageMap.ContainsKey($normalizedLogName)) {
            $logUsageMap[$normalizedLogName] = [PSCustomObject]@{
                Users = @{}
                UsedOverride = $null
            }
        }
        $mapEntry = $logUsageMap[$normalizedLogName]
        if ($user) { $mapEntry.Users[$user.ToLowerInvariant()] = 1 }
        if ($null -ne $usedUsers) {
            if ($null -eq $mapEntry.UsedOverride -or $usedUsers -gt $mapEntry.UsedOverride) {
                $mapEntry.UsedOverride = [int]$usedUsers
            }
        }
    }
    
    # 固定显示指定的 4 个 License：AAD_PREMIUM_P2_CN、POWER_BI_PRO、SPE_E3_NO_WIN、EXCHANGEENTPRISE
    # 使用 $graphSkuToFixedLicenseMap 将 Graph API 返回的 SKU 名称映射为固定的 License 名称
    $graphSkuToFixedLicenseMap = @{
        'AAD_PREMIUM_P2' = 'AAD_PREMIUM_P2_CN'
        'AADPREMIUM_P2' = 'AAD_PREMIUM_P2_CN'
        'BI_AZURE_P2' = 'POWER_BI_PRO'
        'BI_AZURE_P2_ALIAS' = 'POWER_BI_PRO'
        'OFFICESUBSCRIPTION' = 'SPE_E3_NO_WIN'
        'SPE_E3_NO_WIN' = 'SPE_E3_NO_WIN'
        'O365_E3' = 'SPE_E3_NO_WIN'
        'EXCHANGE_S_ENTERPRISE' = 'EXCHANGEENTERPRISE'
        'EXCHANGEENTERPRISE' = 'EXCHANGEENTERPRISE'
    }
    $targetLicenseNames = @('AAD_PREMIUM_P2_CN', 'POWER_BI_PRO', 'SPE_E3_NO_WIN', 'EXCHANGEENTERPRISE')
    $targetLicenseNamesNormalized = @($targetLicenseNames | ForEach-Object { Normalize-LicenseKey -Name $_ })
    
    # 安全获取 SkuList，防止 Null 值错误
    $safeSkuList = @()
    if ($null -ne $graphLicenseResult -and $null -ne $graphLicenseResult.SkuList) {
        $safeSkuList = @($graphLicenseResult.SkuList)
    }
    
    $licenseUsage = @(
        foreach ($item in $safeSkuList) {
            if ($null -eq $item) { continue }
            $skuName = [string]$item.License
            # 检查是否映射到目标 License 名称
            $fixedName = if ($graphSkuToFixedLicenseMap.ContainsKey($skuName)) { $graphSkuToFixedLicenseMap[$skuName] } else { $skuName }
            $normalizedFixedName = Normalize-LicenseKey -Name $fixedName
            if (-not ($targetLicenseNames -contains $fixedName -or $targetLicenseNames -contains $skuName -or $targetLicenseNamesNormalized -contains $normalizedFixedName)) {
                continue
            }
            
            # 将 Graph API 返回的 SKU 名称映射为固定的 License 名称
            $graphSkuName = $skuName
            $fixedLicenseName = if ($graphSkuToFixedLicenseMap.ContainsKey($graphSkuName)) { $graphSkuToFixedLicenseMap[$graphSkuName] } else { $graphSkuName }
            $normalizedGraphName = Normalize-LicenseKey -Name $graphSkuName
            $graphConsumed = if ($null -ne $item.Used) { [int]$item.Used } else { 0 }
            $graphTotal = if ($null -ne $item.Total) { [int]$item.Total } else { 0 }
            
            # 优先使用 Graph API 的 ConsumedUnits 作为已分配数
            $used = $graphConsumed
            $total = $graphTotal
            
            # 尝试从日志中精确匹配使用量（仅用于验证）
            $logUsed = $null
            if ($logUsageMap.ContainsKey($normalizedGraphName)) {
                $logEntry = $logUsageMap[$normalizedGraphName]
                $logUsed = if ($null -ne $logEntry.UsedOverride) { $logEntry.UsedOverride } elseif ($logEntry.Users.Count -gt 0) { $logEntry.Users.Count } else { $null }
                Write-Host "  License $graphSkuName -> $fixedLicenseName : 日志匹配 $normalizedGraphName, 日志使用量=$logUsed, Graph Consumed=$graphConsumed, Graph Total=$graphTotal" -ForegroundColor Gray
            } else {
                Write-Host "  License $graphSkuName -> $fixedLicenseName : 无日志匹配, 使用 Graph Consumed=$graphConsumed, Graph Total=$graphTotal" -ForegroundColor Yellow
            }
            
            # 如果日志使用量可用且大于 Graph Consumed，使用日志值
            if ($null -ne $logUsed -and $logUsed -gt $used) {
                $used = $logUsed
            }
            
            # 合理性检查：如果 Used > Total，将 Total 调整为 Used
            if ($used -gt $total) {
                Write-Host "  警告: License $fixedLicenseName Used($used) > Total($total), 调整 Total=$used" -ForegroundColor Yellow
                $total = $used
            }
            
            $remaining = [Math]::Max(0, $total - $used)
            
            [PSCustomObject]@{ 
                License = $fixedLicenseName
                Used = $used
                Total = $total
                Remaining = $remaining
                Source = 'Graph'
            }
        }
    )
    $licenseStatusNote = "$licenseStatusNote $($graphLicenseResult.Message)"
}

# ==================== License 数据统计（已通过上面的代码获取）====================
# 注意：上面的代码已经处理了 License 数据的获取和映射
# 这里不再重复定义 $licenseUsage，避免覆盖上面的修复结果
$licenseStatusNote = 'License 数据通过 Microsoft Graph PowerShell SDK (Connect-MgGraph -Environment China) 获取。'

# 输出统计摘要
Write-Host ''
Write-Host '=== License 统计摘要 ===' -ForegroundColor Cyan
foreach ($lic in $licenseUsage) {
    Write-Host "  $($lic.License): Total=$($lic.Total), Used=$($lic.Used), Remaining=$($lic.Remaining), Source=$($lic.Source)" -ForegroundColor White
}
Write-Host '========================' -ForegroundColor Cyan

$sharedMailboxRows = [System.Collections.Generic.List[object]]::new()
$mailboxRows = Get-LatestMailboxRows -Rows @($datasets | Where-Object { $_.Table -eq 'MailboxStatisticsDCR_CL' } | ForEach-Object { $_.Rows })
foreach ($row in $mailboxRows) {
    $available = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('AvailableSpaceGB', 'AvailableSpaceGB_d', 'AvailableSpaceGB_r', 'AvailableSpaceGB_s', 'AvailableSpaceInGB', 'AvailableSpaceInGB_d', 'AvailableSpaceInGB_r', 'AvailableSpaceInGB_s', 'AvailableSpace', 'AvailableSpace_d', 'AvailableSpace_r', 'AvailableSpace_s') -Default '')
    $quota = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('QuotaLimitGB', 'QuotaLimitGB_d', 'QuotaLimitGB_r', 'QuotaLimitGB_s', 'QuotaGB', 'QuotaGB_d', 'QuotaGB_r', 'QuotaGB_s', 'StorageQuotaGB', 'StorageQuotaGB_d', 'StorageQuotaGB_r', 'StorageQuotaGB_s', 'ProhibitSendReceiveQuotaGB', 'ProhibitSendReceiveQuotaGB_d', 'ProhibitSendReceiveQuotaGB_r', 'ProhibitSendReceiveQuotaGB_s') -Default '')
    $usagePercent = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('UsagePercent') -Default '')
    $user = Get-UserValue -Row $row -TableName 'MailboxStatisticsDCR_CL'
    $size = Get-NumberValue (Get-AnyFieldValue -Row $row -Names @('TotalItemSizeGB', 'TotalItemSizeInGB', 'MailboxSizeGB', 'MailboxSize', 'SizeGB', 'TotalSizeGB', 'TotalItemSize', 'StorageUsedGB', 'StorageUsed') -Default '')
    if ($null -eq $size -and $null -ne $available -and $null -ne $quota -and $quota -ge $available) {
        $size = $quota - $available
    }
    $type = Get-MailboxTypeText -Row $row
    $isCapacityRisk = ($null -ne $available -and $null -ne $quota -and $quota -gt 0 -and ($available / $quota) -lt 0.05)

    if ((Test-SharedMailboxRow -Row $row) -or $isCapacityRisk) {
        $sharedMailboxRows.Add([PSCustomObject]@{
            DisplayName = Get-MailboxDisplayName -Row $row
            EmailAddress = Get-MailboxEmailAddress -Row $row
            Type = if ($type) { $type } else { 'SharedMailbox' }
            TotalCapacityGB = if ($null -ne $quota) { [Math]::Round($quota, 2) } else { 'N/A' }
            RemainingCapacityGB = if ($null -ne $available) { [Math]::Round($available, 2) } else { 'N/A' }
            Usage = if ($null -ne $usagePercent) { "$usagePercent%" } elseif ($isCapacityRisk) { '{0:P1}' -f (1 - ($available / $quota)) } else { 'N/A' }
            CapacityText = if ($null -ne $available -and $null -ne $quota) { "$([Math]::Round($available, 2))/$([Math]::Round($quota, 2))" } else { 'N/A' }
            CapacityRisk = if ($isCapacityRisk) { '是' } else { '否' }
            CapacityRiskSort = if ($isCapacityRisk) { 0 } else { 1 }
            RemainingSort = if ($null -ne $available) { [double]$available } else { [double]::MaxValue }
        }) | Out-Null
    }
}

# 解析时间范围用于滑动窗口算法
$startUtcForWindow = if ($actualStartUtc) { [DateTime]::Parse($actualStartUtc) } else { [DateTime]::MinValue }
$endUtcForWindow = if ($actualEndUtc) { [DateTime]::Parse($actualEndUtc) } else { [DateTime]::MaxValue }
$suspiciousIpRows = Get-SuspiciousIpSlidingWindowRows -Rows @($suspiciousIpRecords | Where-Object { -not (Test-CachedTrustedIp -IP $_.IP -Rules $trustedRules) }) -WindowDays 3 -Threshold 10 -StartUtc $startUtcForWindow -EndUtc $endUtcForWindow
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
    MailboxLowSpace = @($sharedMailboxRows | Where-Object { $_.CapacityRisk -eq '是' }).Count
    SharedMailboxes = $sharedMailboxRows.Count
    IdentityPermissionChanges = Get-EventCountSum -Rows $identityPermissionChanges
    DcrLogErrors = $dcrLogErrorRows.Count
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

# 替换 KQL 语句中的时间变量为实际值
$actualStartUtc = if ($StartUtc) { $StartUtc } else { '2026-06-08T02:00:00Z' }
$actualEndUtc = if ($EndUtc) { $EndUtc } else { '2026-06-15T02:00:00Z' }
$trustedIpKqlLiteral = Get-TrustedIpKqlDynamicLiteral

$failedSigninKql = @"
union withsource=TableName AADManagedIdentitySignInLogs, AADServicePrincipalSignInLogs, SigninLogs
| where TimeGenerated >= datetime($actualStartUtc) and TimeGenerated < datetime($actualEndUtc)
| extend __principal = case(
    TableName == "SigninLogs", tostring(coalesce(column_ifexists("UserPrincipalName", ""), column_ifexists("UserDisplayName", ""), column_ifexists("Identity", ""), "Unknown")),
    tostring(coalesce(column_ifexists("ServicePrincipalName", ""), column_ifexists("ManagedIdentityName", ""), column_ifexists("Identity", ""), column_ifexists("AppDisplayName", ""), column_ifexists("ServicePrincipalId", ""), column_ifexists("AppId", ""), "Unknown"))
)
| extend __operation = case(TableName == "SigninLogs", tostring(coalesce(column_ifexists("AppDisplayName", ""), "Unknown")), __principal)
| extend __ip = tostring(coalesce(column_ifexists("IPAddress", ""), column_ifexists("IpAddress", ""), column_ifexists("ClientIP", ""), column_ifexists("ClientIpAddress", ""), ""))
| extend __result = tostring(coalesce(column_ifexists("ResultSignature", ""), column_ifexists("Result", ""), column_ifexists("ResultType", ""), column_ifexists("Status", ""), column_ifexists("ResultDescription", ""), ""))
| extend __detail = tostring(coalesce(column_ifexists("ResultDescription", ""), column_ifexists("FailureReason", ""), column_ifexists("Status", ""), __result))
| extend __status = tolower(__result)
| extend __isSuccess = (__status in ("true","success","succeeded","completed","complete","ok","pass","passed","0"))
| extend __isFailed = (isnotempty(__status) and not(__isSuccess) and __status != "unknown") or tolower(__detail) has_any ("fail","failed","failure","denied","error","timeout")
| where __isFailed
| where not(strcat(__principal, " ", __operation, " ", __detail) contains "keyence.com.cn")
| summarize Count=count(), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated) by TableName, IP=__ip, User=__principal, Operation=__operation, Detail=__detail, Result=__result
| where TableName != "AADServicePrincipalSignInLogs" or Count > 10
| project Count, LastTime, IP, 主体_应用摘要=strcat(User, " / ", Operation), Detail, TableName, FirstTime, Result
| sort by Count desc, LastTime desc
| take 50
"@
$deleteDisableKql = @"
AuditLogs
| where TimeGenerated >= datetime($actualStartUtc) and TimeGenerated < datetime($actualEndUtc)
| extend __isDeleteOperation = tostring(AADOperationType) == "Delete"
| where __isDeleteOperation
| where OperationName !contains "PIM"
| extend __RecordKind = "DeleteOperation"
| project TimeGenerated, OperationName, AADOperationType, Actor=tostring(InitiatedBy.user.userPrincipalName), Target=tostring(TargetResources), Result, CorrelationId, __RecordKind
| order by TimeGenerated desc
"@
$suspiciousSuccessKql = @"
union withsource=TableName AADManagedIdentitySignInLogs, AADServicePrincipalSignInLogs, SigninLogs
| where TimeGenerated >= datetime($actualStartUtc) and TimeGenerated < datetime($actualEndUtc)
| extend __ResultSignatureRaw = tostring(column_ifexists("ResultSignature", "")), __ResultRaw = tostring(column_ifexists("Result", "")), __ResultTypeRaw = tostring(column_ifexists("ResultType", "")), __StatusRaw = tostring(column_ifexists("Status", "")), __ResultDescriptionRaw = tostring(column_ifexists("ResultDescription", ""))
| extend ResultType = case(isnotempty(__ResultSignatureRaw), __ResultSignatureRaw, isnotempty(__ResultRaw), __ResultRaw, isnotempty(__ResultTypeRaw), __ResultTypeRaw, isnotempty(__StatusRaw), __StatusRaw, isnotempty(__ResultDescriptionRaw), __ResultDescriptionRaw, "")
| extend AppDisplayName = case(isnotempty(tostring(column_ifexists("AppDisplayName", ""))), tostring(column_ifexists("AppDisplayName", "")), isnotempty(tostring(column_ifexists("ServicePrincipalName", ""))), tostring(column_ifexists("ServicePrincipalName", "")), isnotempty(tostring(column_ifexists("ManagedIdentityName", ""))), tostring(column_ifexists("ManagedIdentityName", "")), "Unknown")
| extend UserOrApp = strcat(tostring(column_ifexists("UserPrincipalName", "")), " ", tostring(column_ifexists("UserDisplayName", "")), " ", tostring(column_ifexists("Identity", "")), " ", AppDisplayName)
| extend IPAddressText = case(isnotempty(tostring(column_ifexists("IPAddress", ""))), tostring(column_ifexists("IPAddress", "")), isnotempty(tostring(column_ifexists("IpAddress", ""))), tostring(column_ifexists("IpAddress", "")), isnotempty(tostring(column_ifexists("ClientIP", ""))), tostring(column_ifexists("ClientIP", "")), tostring(column_ifexists("ClientIpAddress", "")))
| extend __status = tolower(ResultType), __resultCode = tolong(ResultType)
| extend __isSuccess = __status in ("true","success","succeeded","completed","complete","ok","pass","passed","0") or __resultCode == 0
| extend __ip = extract(@"(?:^|[^0-9])((?:[0-9]{1,3}\.){3}[0-9]{1,3})(?:$|[^0-9])", 1, IPAddressText)
| extend __isPublicIp = isnotempty(__ip) and not(ipv4_is_private(__ip))
| extend __isTrustedIp = __isPublicIp and ipv4_is_in_any_range(__ip, $trustedIpKqlLiteral)
| where __isSuccess and __isPublicIp and not(__isTrustedIp)
| where UserOrApp !has "Sangfor"
| where TableName != "SigninLogs" or AppDisplayName !in~ ("Windows Sign In", "Microsoft Edge", "Microsoft Office")
| project TimeGenerated, TableName, UserOrApp, AppDisplayName, IPAddress=__ip, ResultType, ResultDescription=__ResultDescriptionRaw
| sort by TimeGenerated desc
"@
# 注意：实际的可疑IP计算在PowerShell中完成（Get-SuspiciousIpSlidingWindowRows函数）
# 因为实际查询返回的数据已经按用户/应用/IP聚合过了，不能直接用KQL的自连接来计算滑动窗口
# 以下KQL仅用于展示目的，说明数据筛选逻辑
$suspiciousIpKql = @"
// 可疑IP检测逻辑（实际计算在PowerShell中完成）
// 1. 查询SigninLogs中非信任IP的公共IP登录记录
// 2. 按IP分组，使用3天滑动窗口计算登录次数
// 3. 只显示窗口内登录次数 >= 10 的IP
//
// 基础数据查询（与New-SigninLogsOptimizedQuery一致）：
SigninLogs
| where TimeGenerated >= datetime($actualStartUtc) and TimeGenerated < datetime($actualEndUtc)
| extend IPAddress = tostring(column_ifexists("IPAddress", ""))
| extend __ip = extract(@"(\d{1,3}(?:\.\d{1,3}){3})", 1, IPAddress)
| extend __isPublicIp = isnotempty(__ip) and not(ipv4_is_private(__ip))
| extend __isTrustedIp = __isPublicIp and ipv4_is_in_any_range(__ip, $trustedIpKqlLiteral)
| where __isPublicIp and not(__isTrustedIp)
| summarize EventCount=count() by IPAddress=__ip, TimeGenerated
| order by TimeGenerated desc
// 然后PowerShell中使用Get-SuspiciousIpSlidingWindowRows函数计算3天滑动窗口
// 筛选条件：窗口内登录次数 >= 10
"@
$clientIpRankLogic = @"
union withsource=TableName AADManagedIdentitySignInLogs, AADServicePrincipalSignInLogs, SigninLogs, AuditLogs, DCRLogErrors, IntuneAuditLogsDCR_CL
| where TimeGenerated >= datetime($actualStartUtc) and TimeGenerated < datetime($actualEndUtc)
| extend ClientIp = tostring(coalesce(column_ifexists("IPAddress", ""), column_ifexists("IpAddress", ""), column_ifexists("ClientIP", ""), column_ifexists("ClientIpAddress", ""), column_ifexists("CallerIpAddress", "")))
| extend __ip = extract(@"(?:^|[^0-9])((?:[0-9]{1,3}\.){3}[0-9]{1,3})(?:$|[^0-9])", 1, ClientIp)
| where isnotempty(__ip) and not(ipv4_is_private(__ip))
| project-away __ip
"@

# License 数据永远只查询当天（本地时间今天）的 AssignedLicensesDCR_CL 记录
# 不管用户查询的时间范围是什么
$todayDate = Get-Date -Format 'yyyy-MM-dd'
$licenseLogic = @"
AssignedLicensesDCR_CL
| where TimeGenerated >= datetime('$todayDate 00:00:00') and TimeGenerated < datetime('$todayDate 23:59:59')
| where ServicePlanName in ("AAD_PREMIUM_P2", "BI_AZURE_P2", "OFFICESUBSCRIPTION", "EXCHANGE_S_ENTERPRISE")
| where ProvisioningStatus in ("Success", "PendingInput")
| summarize count() by ServicePlanName, ProvisioningStatus
"@

$permissionKql = @"
AuditLogs
| where TimeGenerated >= datetime($actualStartUtc) and TimeGenerated < datetime($actualEndUtc)
| extend __isPermissionChange = tostring(Result) =~ "success" and OperationName in (
    "Add delegated permission grant",
    "Consent to application",
    "Create application – Certificates and secrets management",
    "Add owner to application",
    "Add app role assignment to service principal",
    "Update application – Certificates and secrets management",
    "Remove delegated permission grant",
    "Remove app role assignment from service principal"
)
| where __isPermissionChange
| where OperationName !contains "PIM"
| project TimeGenerated, 
    OperationName, 
    AADOperationType,
    Actor = tostring(InitiatedBy.user.userPrincipalName),
    Target = tostring(TargetResources),
    Result,
    CorrelationId
| order by TimeGenerated desc
"@


$failedSigninGrouped = Group-EventRecords -Rows $failedSignins -KeyBuilder { param($r) Get-StrictEventMergeKey -Row $r }
# 在分组后应用过滤（与KQL逻辑一致）：
# 1. 过滤掉主体/应用摘要中含有 keyence.com.cn 的行
# 2. 对于 AADServicePrincipalSignInLogs，只保留 Count > 10 的行
$failedSigninFiltered = @($failedSigninGrouped | Where-Object {
    $concatText = "$($_.User) $($_.Operation) $($_.Detail)"
    # 使用 -like 通配符匹配，与 KQL contains 操作符行为一致（子串匹配，非分词匹配）
    $containsKeyence = $concatText -like '*keyence.com.cn*'
    if ($containsKeyence) { return $false }
    if ($_.Table -eq 'AADServicePrincipalSignInLogs' -and $_.Count -le 10) { return $false }
    return $true
})
$failedSigninHtml = (New-CodeBlockHtml -Text $failedSigninKql) + (New-TableHtml -Rows ($failedSigninFiltered | Select-Object -First 10000) -Columns @('风险等级', '次数', '最后时间', 'IP', '主体/应用摘要', '说明') -CellBuilder {
    param($r) 
    $riskLevel = Get-RiskLevel -Category 'FailedSignins' -Row $r
    $riskBadge = Get-RiskLevelBadge -Level $riskLevel
    @($riskBadge, $r.Count, $r.LastTime, $r.IP, (($r.User, $r.Operation) -join ' / '), $r.Detail)
} -RawHtmlColumns @('风险等级'))
# 删除/Disable 操作栏不做任何合并，每条记录独立显示，使用每条记录自己的发生
$deleteDisableHtml = (New-CodeBlockHtml -Text $deleteDisableKql) + (New-TableHtml -Rows ($deleteDisableEvents | Select-Object -First 10000) -Columns @('风险等级', '时间', '操作者', '操作', '被删除者') -CellBuilder {
    param($r) 
    $riskLevel = Get-RiskLevel -Category 'DeleteDisable' -Row $r
    $riskBadge = Get-RiskLevelBadge -Level $riskLevel
    @($riskBadge, $r.Time, (Format-UserForReport -User $r.User), $r.Operation, (Format-DeleteTargetForReport -Target $r.Target -Operation $r.Operation))
} -RawHtmlColumns @('风险等级'))
$suspiciousIpHtml = (New-CodeBlockHtml -Text $suspiciousIpKql) + (New-TableHtml -Rows $suspiciousIpRows -Columns @('风险等级', 'IP', '次数', '首次访问时间', '最近访问时间') -CellBuilder {
    param($r) 
    $riskLevel = Get-RiskLevel -Category 'SuspiciousIP' -Row $r
    $riskBadge = Get-RiskLevelBadge -Level $riskLevel
    $firstAccess = if ($r.PSObject.Properties.Name -contains 'FirstAccess') { $r.FirstAccess } else { '' }
    $lastAccess = if ($r.PSObject.Properties.Name -contains 'LastAccess') { $r.LastAccess } else { '' }
    @($riskBadge, $r.IP, $r.Count, $firstAccess, $lastAccess)
} -RawHtmlColumns @('风险等级'))
$signinSuspiciousGrouped = Group-EventRecords -Rows $suspiciousSigninSuccess -KeyBuilder { param($r) Get-StrictEventMergeKey -Row $r }
$signinSuspiciousHtml = (New-CodeBlockHtml -Text $suspiciousSuccessKql) + (New-TableHtml -Rows ($signinSuspiciousGrouped | Select-Object -First 10000) -Columns @('风险等级', '次数', '最后时间', '用户', '应用', 'IP', '说明') -CellBuilder {
    param($r) 
    $riskLevel = Get-RiskLevel -Category 'SuspiciousSuccess' -Row $r
    $riskBadge = Get-RiskLevelBadge -Level $riskLevel
    @($riskBadge, $r.Count, $r.LastTime, $r.User, $r.Operation, $r.IP, $r.Reason)
} -RawHtmlColumns @('风险等级'))
$topClientIpHtml = (New-CodeBlockHtml -Text $clientIpRankLogic) + (New-TableHtml -Rows $topClientIps -Columns @('IP', '次数') -CellBuilder {
    param($r) @($r.IP, $r.Count)
})

# ==================== License 状态判断函数
function Get-LicenseStatus {
    param([int]$Remaining)
    if ($Remaining -eq 0) {
        return '<span style="color:#ff6b6b;">🔴 已耗尽</span>'
    } elseif ($Remaining -lt 5) {
        return '<span style="color:#f3b95f;">⚠️ 即将耗尽</span>'
    } else {
        return '<span style="color:#66d98f;">✅ 充足</span>'
    }
}

# 构建 License 表格行
$licenseTableRows = @(
    foreach ($license in $licenseUsage) {
        $skuPartNumber = [string]$license.License
        $productName = if ($licenseProductMap.ContainsKey($skuPartNumber)) { $licenseProductMap[$skuPartNumber] } else { $skuPartNumber }
        $total = if ($null -ne $license.Total -and $license.Total -ne 'N/A') { [int]$license.Total } else { 0 }
        $used = if ($null -ne $license.Used) { [int]$license.Used } else { 0 }
        $remaining = if ($null -ne $license.Remaining -and $license.Remaining -ne 'N/A') { [int]$license.Remaining } else { 0 }
        $status = Get-LicenseStatus -Remaining $remaining
        [PSCustomObject]@{
            SkuPartNumber = $skuPartNumber
            ProductName = $productName
            Total = $total
            Used = $used
            Remaining = $remaining
            Status = $status
        }
    }
)

$licenseHtml = (New-CodeBlockHtml -Text $licenseLogic) + (New-TableHtml -Rows $licenseTableRows -Columns @('SkuPartNumber', '产品名称', '总数', '已分配', '剩余', '状态') -CellBuilder {
    param($r) @($r.SkuPartNumber, $r.ProductName, $r.Total, $r.Used, $r.Remaining, $r.Status)
} -RawHtmlColumns @('状态'))
# 使用与实际查询相同的函数生成KQL，确保报告中显示的KQL可以直接执行
$sharedMailboxKql = New-MailboxStatisticsOptimizedQuery -StartUtc $actualStartUtc -EndUtc $actualEndUtc
$sharedMailboxHtml = (New-CodeBlockHtml -Text $sharedMailboxKql) + (New-TableHtml -Rows ($sharedMailboxRows | Sort-Object CapacityRiskSort, RemainingSort, DisplayName) -Columns @('风险等级', '用户名', '邮箱', '剩余容量/总容量', '邮箱容量是否风险') -CellBuilder {
    param($r) 
    $riskLevel = Get-RiskLevel -Category 'MailboxLowSpace' -Row $r
    $riskBadge = Get-RiskLevelBadge -Level $riskLevel
    @($riskBadge, $r.DisplayName, $r.EmailAddress, $r.CapacityText, $r.CapacityRisk)
} -RawHtmlColumns @('风险等级'))
# DCRLogErrors 的 KQL 需要反映实际的聚合逻辑
$dcrLogErrorsKql = @"
DCRLogErrors
| where TimeGenerated >= datetime($actualStartUtc) and TimeGenerated < datetime($actualEndUtc)
| summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by InputStreamId, OperationName, Message
| project TimeGenerated, FirstTime, LastTime, EventCount, InputStreamId, OperationName, Message
| order by TimeGenerated desc
"@
$dcrLogErrorHtml = (New-CodeBlockHtml -Text $dcrLogErrorsKql) + (New-TableHtml -Rows ($dcrLogErrorRows | Select-Object -First 10000) -Columns @('风险等级', '次数', '时间', '输入流ID', '操作名称', '消息') -CellBuilder {
    param($r) 
    $riskLevel = Get-RiskLevel -Category 'DcrLogErrors' -Row $r
    $riskBadge = Get-RiskLevelBadge -Level $riskLevel
    @($riskBadge, $r.Count, $r.Time, $r.Target, $r.Operation, $r.Detail)
} -RawHtmlColumns @('风险等级'))
$permissionHtml = (New-CodeBlockHtml -Text $permissionKql) + (New-TableHtml -Rows ($identityPermissionChanges | Sort-Object -Property ActivityDateTime -Descending) -Columns @('风险等级', '活动时间', '操作者', '操作', '结果/说明') -CellBuilder {
    param($r) 
    $riskLevel = Get-RiskLevel -Category 'IdentityPermission' -Row $r
    $riskBadge = Get-RiskLevelBadge -Level $riskLevel
    $permValue = $r.Detail
    
    # 如果 Detail 看起来像 GUID，回退到 PermissionName
    if ($permValue -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
        $permValue = $r.PermissionName
    }
    # 如果 PermissionName 也像是 GUID，尝试从 Detail 中提取可读名称
    if ($permValue -match '^[0-9a-f]{8}') {
        if ($r.Detail -match '"displayName"\s*:\s*"([^"]+)"') {
            $permValue = $matches[1]
        }
        elseif ($r.Detail -match '"appRoleValue"\s*:\s*"([^"]+)"') {
            $permValue = $matches[1]
        }
    }
    
    @($riskBadge, $r.ActivityDateTime, (Format-UserForReport -User $r.User), $r.Operation, (Format-CompactTextForReport -Text $permValue -MaxLength 80))
} -RawHtmlColumns @('风险等级'))
# IntuneAuditLogs 的 KQL 需要反映实际的聚合逻辑
$intuneAuditKql = @"
IntuneAuditLogsDCR_CL
| where TimeGenerated >= datetime($actualStartUtc) and TimeGenerated < datetime($actualEndUtc)
| extend ActorDisplayName = tostring(coalesce(column_ifexists("InitiatorDisplayName", ""), column_ifexists("InitiatorDisplayName_s", ""), column_ifexists("ActorDisplayName", ""), column_ifexists("ActorDisplayName_s", ""), column_ifexists("DisplayName", ""), column_ifexists("DisplayName_s", ""), column_ifexists("InitiatedByUserDisplayName", ""), column_ifexists("InitiatedByUserDisplayName_s", ""), column_ifexists("UserDisplayName", ""), column_ifexists("UserDisplayName_s", ""), ""))
| extend ActorUserPrincipalName = tostring(coalesce(column_ifexists("InitiatorUserPrincipalName", ""), column_ifexists("InitiatorUserPrincipalName_s", ""), column_ifexists("ActorInitiator", ""), column_ifexists("ActorUPN", ""), column_ifexists("ActorUPN_s", ""), column_ifexists("ActorUserPrincipalName", ""), column_ifexists("ActorUserPrincipalName_s", ""), column_ifexists("InitiatedByUserPrincipalName", ""), column_ifexists("InitiatedByUserPrincipalName_s", ""), column_ifexists("UserPrincipalName", ""), column_ifexists("UserPrincipalName_s", ""), column_ifexists("UPN", ""), column_ifexists("UPN_s", ""), column_ifexists("Actor", ""), column_ifexists("Actor_s", ""), column_ifexists("UserId", ""), column_ifexists("UserId_s", ""), column_ifexists("Identity", ""), column_ifexists("Identity_s", ""), ""))
| extend Actor = case(isnotempty(ActorDisplayName) and isnotempty(ActorUserPrincipalName) and ActorDisplayName != ActorUserPrincipalName, strcat(ActorDisplayName, " / ", ActorUserPrincipalName), isnotempty(ActorDisplayName), ActorDisplayName, ActorUserPrincipalName)
| extend OperationName = tostring(coalesce(column_ifexists("OperationName", ""), column_ifexists("OperationName_s", ""), column_ifexists("ActivityDisplayName", ""), column_ifexists("ActivityDisplayName_s", ""), column_ifexists("Activity", ""), column_ifexists("Activity_s", ""), column_ifexists("Operation", ""), column_ifexists("Operation_s", ""), column_ifexists("Action", ""), column_ifexists("Action_s", ""), column_ifexists("AuditEventType", ""), column_ifexists("AuditEventType_s", ""), "Intune Audit Event"))
| extend TargetDeviceName = tostring(coalesce(column_ifexists("TargetDeviceName", ""), column_ifexists("TargetDeviceName_s", ""), column_ifexists("DeviceName", ""), column_ifexists("DeviceName_s", ""), column_ifexists("ManagedDeviceName", ""), column_ifexists("ManagedDeviceName_s", ""), ""))
| extend Result = tostring(coalesce(column_ifexists("Result", ""), column_ifexists("Result_s", ""), column_ifexists("ResultStatus", ""), column_ifexists("ResultStatus_s", ""), column_ifexists("Status", ""), column_ifexists("Status_s", ""), column_ifexists("ActivityResult", ""), column_ifexists("ActivityResult_s", ""), column_ifexists("OperationStatus", ""), column_ifexists("OperationStatus_s", ""), ""))
| extend ResultDescription = tostring(coalesce(column_ifexists("ResultDescription", ""), column_ifexists("ResultDescription_s", ""), column_ifexists("FailureReason", ""), column_ifexists("FailureReason_s", ""), column_ifexists("Message", ""), column_ifexists("Message_s", ""), column_ifexists("ErrorMessage", ""), column_ifexists("ErrorMessage_s", ""), ""))
| summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by Actor, ActorDisplayName, ActorUserPrincipalName, OperationName, TargetDeviceName, Result, ResultDescription
| order by TimeGenerated desc
"@
$intuneGrouped = Group-EventRecords -Rows $intuneAuditRows -KeyBuilder { param($r) Get-StrictEventMergeKey -Row $r }
$intuneHtml = (New-CodeBlockHtml -Text $intuneAuditKql) + (New-TableHtml -Rows ($intuneGrouped | Select-Object -First 10000) -Columns @('风险等级', '次数', '最后时间', '操作者', '操作', '目标', '结果/说明') -CellBuilder {
    param($r) 
    $riskLevel = Get-RiskLevel -Category 'IntuneAudit' -Row $r
    $riskBadge = Get-RiskLevelBadge -Level $riskLevel
    @($riskBadge, $r.Count, $r.LastTime, (Format-UserForReport -User $r.User), $r.Operation, $r.Target, $r.Detail)
} -RawHtmlColumns @('风险等级'))

$sectionSpecs = @(
    [PSCustomObject]@{ Id = 'failed-signins'; Title = '应用登录失败'; Note = @"
Managed Identity / Service Principal 登录失败 → 依赖该身份的服务可能无法运行。
合并规则：操作者、操作内容、时间戳完全相同才合并。
"@; Content = $failedSigninHtml; Open = $true; AiAnalysis = Get-AiAnalysis -Category 'FailedSignins' -Data $failedSigninFiltered },
    [PSCustomObject]@{ Id = 'identity-permission'; Title = '应用权限变更'; Note = '显示所选时间范围内 AuditLogs 表中 Result 为 success 的全部记录，不再按 Service Principal 操作类型或权限字段额外过滤。'; Content = $permissionHtml; Open = $true; AiAnalysis = Get-AiAnalysis -Category 'IdentityPermission' -Data $identityPermissionChanges },
    [PSCustomObject]@{ Id = 'delete-disable'; Title = '删除操作'; Note = '只统计 delete 语义的操作；每条记录独立显示，不做合并。'; Content = $deleteDisableHtml; Open = $true; AiAnalysis = Get-AiAnalysis -Category 'DeleteDisable' -Data $deleteDisableEvents },
    [PSCustomObject]@{ Id = 'suspicious-success'; Title = '可疑成功登录'; Note = '关注 AADManagedIdentitySignInLogs / AADServicePrincipalSignInLogs / SigninLogs 三张表；SigninLogs 仍排除 Windows Sign In / Microsoft Edge / Sangfor SASE VPN / Microsoft Office。仅当操作者、操作内容、时间戳完全相同时合并。'; Content = $signinSuspiciousHtml; Open = $true; AiAnalysis = Get-AiAnalysis -Category 'SuspiciousSuccess' -Data $signinSuspiciousGrouped },
    [PSCustomObject]@{ Id = 'suspicious-ip'; Title = '可疑 IP'; Note = "仅统计 SigninLogs 中的可疑 IP；同一 IP 在任意连续 3 天窗口内出现 10 次及以上时展示；已排除 TrustedLocation_KJ.txt、TrustedLocation_IDC_Ali.txt 中的可信 IP，$microsoftTrustedNote"; Content = $suspiciousIpHtml; Open = $true; AiAnalysis = Get-AiAnalysis -Category 'SuspiciousIP' -Data $suspiciousIpRows },
    [PSCustomObject]@{ Id = 'license'; Title = 'License 使用量与剩余数量'; Note = $licenseStatusNote; Content = $licenseHtml; Open = $true; AiAnalysis = '' },
    [PSCustomObject]@{ Id = 'shared-mailbox'; Title = 'SharedMailbox'; Note = "显示 MailboxStatisticsDCR_CL 中最新快照识别出的 SharedMailbox，剩余容量不足5%的邮箱视为有风险，有风险的邮箱排在前面，共 $($sharedMailboxRows.Count) 个邮箱。每个邮箱只保留最新记录。"; Content = $sharedMailboxHtml; Open = $true; AiAnalysis = Get-AiAnalysis -Category 'MailboxLowSpace' -Data @($sharedMailboxRows | Where-Object { $_.CapacityRisk -eq '是' }) },
    [PSCustomObject]@{ Id = 'dcr-log-errors'; Title = 'DCRLogErrors'; Note = '固定观察 DCRLogErrors 表，并按最近 30 天的时间、InputStreamId、OperationName、Message 展示。'; Content = $dcrLogErrorHtml; Open = $true; AiAnalysis = Get-AiAnalysis -Category 'DcrLogErrors' -Data $dcrLogErrorRows },
    [PSCustomObject]@{ Id = 'intune-audit'; Title = 'Intune 审计记录'; Note = '显示 IntuneAuditLogsDCR_CL 在所选时间范围内的审计记录，并兼容自定义日志常见的 _s 后缀字段。'; Content = $intuneHtml; Open = $true; AiAnalysis = Get-AiAnalysis -Category 'IntuneAudit' -Data $intuneGrouped }
)

# 定义二级分类目录结构
$categoryOrder = @(
    [PSCustomObject]@{ Key = 'login-security'; Label = '登录与身份安全'; Icon = '🔐'; Sections = @('suspicious-success', 'suspicious-ip', 'failed-signins') },
    [PSCustomObject]@{ Key = 'operation-audit'; Label = '操作审计'; Icon = '📋'; Sections = @('identity-permission', 'delete-disable') },
    [PSCustomObject]@{ Key = 'mailbox'; Label = '邮箱安全'; Icon = '📧'; Sections = @('shared-mailbox') },
    [PSCustomObject]@{ Key = 'data-source'; Label = '数据源与许可证'; Icon = '💾'; Sections = @('dcr-log-errors', 'intune-audit', 'license') }
)

# 构建二级目录 HTML（使用 details/summary 实现一级分类折叠）
$sideNavHtml = '<nav class="side-nav"><div class="nav-title" data-i18n="nav.title">目录</div>'
foreach ($category in $categoryOrder) {
    $sideNavHtml += '<details class="nav-category">'
    $sideNavHtml += '<summary class="nav-category-summary"><span>' + (Escape-Html $category.Icon) + ' </span><span data-i18n="category.' + (Escape-Html $category.Key) + '">' + (Escape-Html $category.Label) + '</span></summary>'
    $sideNavHtml += '<div class="nav-submenu">'
    foreach ($sectionId in $category.Sections) {
        $section = $sectionSpecs | Where-Object { $_.Id -eq $sectionId }
        if ($section) {
            $sideNavHtml += '<a href="#' + (Escape-Html $section.Id) + '" class="nav-item" data-i18n="section.' + (Escape-Html $section.Id) + '.title">' + (Escape-Html $section.Title) + '</a>'
        }
    }
    $sideNavHtml += '</div>'
    $sideNavHtml += '</details>'
}
$sideNavHtml += '</nav>'

$reportSectionsHtml = @(
    foreach ($section in $sectionSpecs) {
        New-ReportSection -Id $section.Id -Title $section.Title -Note $section.Note -Content $section.Content -Open $section.Open -AiAnalysis $section.AiAnalysis
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
  --bg: #ffffff;
  --panel: #ffffff;
  --panel2: #f5f7fb;
  --text: #111827;
  --muted: #4b5563;
  --line: #d9e0ea;
  --red: #b91c1c;
  --amber: #b45309;
  --green: #047857;
  --blue: #1d4ed8;
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body { margin: 0; background: var(--bg); color: var(--text); font-family: "Segoe UI", Arial, sans-serif; }
.layout { display: grid; grid-template-columns: 250px minmax(0, 1fr); gap: 22px; max-width: 1680px; margin: 0 auto; padding: 28px; }
.wrap { min-width: 0; }
.side-nav { position: sticky; top: 18px; align-self: start; max-height: calc(100vh - 36px); overflow: auto; background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 12px; }
.nav-title { color: var(--muted); font-size: 12px; margin: 2px 6px 10px; }
.nav-category { margin: 4px 0; }
.nav-category summary {
    display: block;
    color: var(--blue);
    cursor: pointer;
    list-style: none;
    border-radius: 6px;
    padding: 8px 9px;
    font-size: 13px;
    font-weight: 600;
    line-height: 1.35;
    transition: background 0.2s;
}
.nav-category summary::-webkit-details-marker { display: none; }
.nav-category summary::before { content: "▸ "; transition: transform 0.2s; display: inline-block; }
.nav-category[open] > summary::before { transform: rotate(90deg); }
.nav-category summary:hover { background: var(--panel2); }
.nav-submenu { padding-left: 16px; }
.nav-item { display: block; color: var(--text); text-decoration: none; border-radius: 6px; padding: 6px 9px; font-size: 12px; line-height: 1.4; margin: 1px 0; }
.nav-item:hover { background: var(--panel2); }
.header { border-bottom: 1px solid var(--line); padding-bottom: 18px; margin-bottom: 22px; }
.header-top { display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; flex-wrap: wrap; }
h1 { margin: 0 0 10px; font-size: 28px; font-weight: 700; }
.language-switcher { display: flex; align-items: center; gap: 8px; background: var(--panel2); border: 1px solid var(--line); border-radius: 8px; padding: 7px 10px; color: var(--muted); font-size: 13px; }
.language-switcher select { background: #ffffff; color: var(--text); border: 1px solid var(--line); border-radius: 6px; padding: 5px 8px; }
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
th { color: var(--muted); font-weight: 600; background: #f3f4f6; position: sticky; top: 0; }
td { color: var(--text); }
.risk-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 12px; }
.small { font-size: 12px; color: var(--muted); }
.kql-block { margin: 14px 18px 0; background: #f8fafc; border: 1px solid var(--line); border-radius: 8px; overflow-x: auto; color: var(--text); font-size: 12px; line-height: 1.55; }
.kql-block summary { cursor: pointer; list-style: none; padding: 10px 14px; font-weight: 600; color: var(--blue); }
.kql-block summary::-webkit-details-marker { display: none; }
.kql-block summary::before { content: "▸"; display: inline-block; margin-right: 6px; transition: transform 0.2s; }
.kql-block[open] summary::before { transform: rotate(90deg); }
.kql-block code { display: block; padding: 10px 14px; background: #ffffff; border-top: 1px solid var(--line); border-radius: 0 0 4px 4px; font-family: Consolas, "Cascadia Mono", monospace; white-space: pre; overflow-x: auto; }
.ip-group { margin: 10px 0; background: var(--panel2); border: 1px solid var(--line); border-radius: 6px; }
.ip-group summary { cursor: pointer; list-style: none; padding: 10px 14px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
.ip-group summary::-webkit-details-marker { display: none; }
.ip-group summary::before { content: "▸"; color: var(--blue); transition: transform 0.2s; }
.ip-group[open] summary::before { transform: rotate(90deg); }
.ip-address { font-weight: 600; color: var(--text); font-family: Consolas, monospace; }
.ip-count { background: var(--red); color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px; font-weight: 600; }
.ip-identities { color: var(--muted); font-size: 12px; flex: 1; min-width: 200px; }
.ip-details { padding: 0 14px 14px; }
.ip-details .note { margin: 8px 0; }
.ai-analysis { background: #fff3cd; border: 1px solid #ffc107; border-radius: 6px; padding: 10px 14px; margin: 10px 18px; color: #856404; font-size: 13px; line-height: 1.5; }
.ai-analysis strong { color: #d63384; }
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
    <div class="header-top">
      <h1 data-i18n="report.title">Log Analytics 合并风险报告</h1>
      <label class="language-switcher"><span data-i18n="language.label">语言</span><select id="languageSelect" aria-label="Language"><option value="zh-CN">中文</option><option value="en-US">English</option><option value="ja-JP">日本語</option></select></label>
    </div>
    <div class="meta">
      <span class="tag"><span data-i18n="meta.timeRange">查询时间段</span>: $(Escape-Html $AnalysisDate)</span>
      <span class="tag"><span data-i18n="meta.tables">数据表</span>: $tableCount</span>
      <span class="tag"><span data-i18n="meta.totalRecords">总记录数</span>: $totalRecords</span>
      <span class="tag"><span data-i18n="meta.trustedIpRules">可信 IP 规则</span>: $trustedCount</span>
    </div>
  </div>

  <div class="summary">
    <div class="card"><div class="label" data-i18n="summary.failedSignins">登录失败</div><div class="value red">$($riskCounts.FailedSignins)</div></div>
    <div class="card"><div class="label" data-i18n="summary.deleteDisable">删除 / Disable</div><div class="value amber">$($riskCounts.DeleteDisable)</div></div>
    <div class="card"><div class="label" data-i18n="summary.suspiciousIp">可疑 IP</div><div class="value amber">$($riskCounts.SuspiciousIPs)</div></div>
    <div class="card"><div class="label" data-i18n="summary.suspiciousSigninSuccess">可信位置外成功登录</div><div class="value amber">$($riskCounts.SuspiciousSigninSuccess)</div></div>
    <div class="card"><div class="label" data-i18n="summary.mailboxLowSpace">邮箱低容量</div><div class="value red">$($riskCounts.MailboxLowSpace)</div></div>
    <div class="card"><div class="label" data-i18n="summary.sharedMailbox">SharedMailbox</div><div class="value blue">$($riskCounts.SharedMailboxes)</div></div>
    <div class="card"><div class="label" data-i18n="summary.dcrLogErrors">DCRLogErrors</div><div class="value red">$($riskCounts.DcrLogErrors)</div></div>
    <div class="card"><div class="label" data-i18n="summary.intuneAudit">Intune 审计记录</div><div class="value amber">$($riskCounts.IntuneAudit)</div></div>
  </div>

  $reportSectionsHtml

  <p class="small"><span data-i18n="footer.generatedAt">Generated at</span> $(Escape-Html ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')))</p>
</div>
</div>
<script>
const i18n = {
  'zh-CN': {
    'report.title': 'Log Analytics 合并风险报告',
    'language.label': '语言',
    'nav.title': '目录',
    'meta.timeRange': '查询时间段',
    'meta.tables': '数据表',
    'meta.totalRecords': '总记录数',
    'meta.trustedIpRules': '可信 IP 规则',
    'footer.generatedAt': 'Generated at',
    'label.kql': 'KQL 语句',
    'empty.noRisk': '未发现相关风险。',
    'category.login-security': '登录与身份安全',
    'category.operation-audit': '操作审计',
    'category.mailbox': '邮箱安全',
    'category.data-source': '数据源与许可证',
    'section.failed-signins.title': '应用登录失败',
    'section.identity-permission.title': '应用权限变更',
    'section.delete-disable.title': '删除操作',
    'section.suspicious-success.title': '可疑成功登录',
    'section.suspicious-ip.title': '可疑 IP',
    'section.license.title': 'License 使用量与剩余数量',
    'section.shared-mailbox.title': 'SharedMailbox',
    'section.dcr-log-errors.title': 'DCRLogErrors',
    'section.intune-audit.title': 'Intune 审计记录',
    'summary.failedSignins': '登录失败',
    'summary.deleteDisable': '删除 / Disable',
    'summary.suspiciousIp': '可疑 IP',
    'summary.suspiciousSigninSuccess': '可信位置外成功登录',
    'summary.mailboxLowSpace': '邮箱低容量',
    'summary.sharedMailbox': 'SharedMailbox',
    'summary.dcrLogErrors': 'DCRLogErrors',
    'summary.intuneAudit': 'Intune 审计记录',
    'field.次数': '次数', 'field.最后时间': '最后时间', 'field.IP': 'IP', 'field.主体_应用摘要': '主体/应用摘要', 'field.说明': '说明',
    'field.表': '表', 'field.用户': '用户', 'field.操作': '操作', 'field.状态_原因': '状态/原因', 'field.时间': '时间', 'field.操作者': '操作者',
    'field.用户身份': '用户身份', 'field.应用名称': '应用名称', 'field.原因': '原因', 'field.应用': '应用', 'field.SkuPartNumber': 'SkuPartNumber',
    'field.产品名称': '产品名称', 'field.总数': '总数', 'field.已分配': '已分配', 'field.剩余': '剩余', 'field.状态': '状态',
    'field.用户名': '用户名', 'field.邮箱': '邮箱', 'field.类型': '类型', 'field.总容量': '总容量', 'field.剩余容量': '剩余容量', 'field.使用率': '使用率', 'field.邮箱容量是否风险': '邮箱容量是否风险',
    'field.输入流ID': '输入流ID', 'field.操作名称': '操作名称', 'field.消息': '消息', 'field.活动时间': '活动时间', 'field.目标': '目标', 'field.权限': '权限',
    'field.结果_说明': '结果/说明', 'field.总记录数': '总记录数', 'field.筛选后记录数': '筛选后记录数', 'field.CSV': 'CSV', 'field.被删除者': '被删除者',
    'field.风险等级': '风险等级', 'field.首次访问时间': '首次访问时间', 'field.最近访问时间': '最近访问时间'
  },
  'en-US': {
    'report.title': 'Log Analytics Merged Risk Report',
    'language.label': 'Language',
    'nav.title': 'Contents',
    'meta.timeRange': 'Time Range',
    'meta.tables': 'Tables',
    'meta.totalRecords': 'Total Records',
    'meta.trustedIpRules': 'Trusted IP Rules',
    'footer.generatedAt': 'Generated at',
    'label.kql': 'KQL Query',
    'empty.noRisk': 'No related risks found.',
    'category.login-security': 'Sign-in & Identity Security',
    'category.operation-audit': 'Operation Audit',
    'category.mailbox': 'Mailbox Security',
    'category.data-source': 'Data Sources & Licenses',
    'section.failed-signins.title': 'AAD / Managed Identity / Service Principal Sign-in Failures',
    'section.identity-permission.title': 'Service Principal Object / Permission Changes',
    'section.delete-disable.title': 'Delete / Disable Operations',
    'section.suspicious-success.title': 'Suspicious Successful Sign-ins',
    'section.suspicious-ip.title': 'Suspicious IPs',
    'section.license.title': 'License Usage and Remaining Count',
    'section.shared-mailbox.title': 'Shared Mailboxes',
    'section.dcr-log-errors.title': 'DCR Log Errors',
    'section.intune-audit.title': 'Intune Audit Records',
    'summary.failedSignins': 'Sign-in Failures',
    'summary.deleteDisable': 'Delete / Disable',
    'summary.suspiciousIp': 'Suspicious IPs',
    'summary.suspiciousSigninSuccess': 'Successful Sign-ins Outside Trusted Locations',
    'summary.mailboxLowSpace': 'Low Mailbox Capacity',
    'summary.sharedMailbox': 'Shared Mailboxes',
    'summary.dcrLogErrors': 'DCR Log Errors',
    'summary.intuneAudit': 'Intune Audit Records',
    'field.次数': 'Count', 'field.最后时间': 'Last Time', 'field.IP': 'IP', 'field.主体_应用摘要': 'Principal / App Summary', 'field.说明': 'Description',
    'field.表': 'Table', 'field.用户': 'User', 'field.操作': 'Operation', 'field.状态_原因': 'Status / Reason', 'field.时间': 'Time', 'field.操作者': 'Actor',
    'field.用户身份': 'Identity', 'field.应用名称': 'App Name', 'field.原因': 'Reason', 'field.应用': 'App', 'field.SkuPartNumber': 'SkuPartNumber',
    'field.产品名称': 'Product Name', 'field.总数': 'Total', 'field.已分配': 'Assigned', 'field.剩余': 'Remaining', 'field.状态': 'Status',
    'field.用户名': 'User Name', 'field.邮箱': 'Email', 'field.类型': 'Type', 'field.总容量': 'Total Capacity', 'field.剩余容量': 'Remaining Capacity', 'field.使用率': 'Usage', 'field.邮箱容量是否风险': 'Mailbox Capacity Risk',
    'field.输入流ID': 'Input Stream ID', 'field.操作名称': 'Operation Name', 'field.消息': 'Message', 'field.活动时间': 'Activity Time', 'field.目标': 'Target', 'field.权限': 'Permission',
    'field.结果_说明': 'Result / Description', 'field.总记录数': 'Total Records', 'field.筛选后记录数': 'Filtered Records', 'field.CSV': 'CSV', 'field.被删除者': 'Deleted Target',
    'field.风险等级': 'Risk Level', 'field.首次访问时间': 'First Access Time', 'field.最近访问时间': 'Last Access Time'
  },
  'ja-JP': {
    'report.title': 'Log Analytics 統合リスクレポート',
    'language.label': '言語',
    'nav.title': '目次',
    'meta.timeRange': '期間',
    'meta.tables': 'テーブル',
    'meta.totalRecords': '総レコード数',
    'meta.trustedIpRules': '信頼済み IP ルール',
    'footer.generatedAt': '生成日時',
    'label.kql': 'KQL クエリ',
    'empty.noRisk': '関連するリスクは見つかりませんでした。',
    'category.login-security': 'サインインと ID セキュリティ',
    'category.operation-audit': '操作監査',
    'category.mailbox': 'メールボックス セキュリティ',
    'category.data-source': 'データソースとライセンス',
    'section.failed-signins.title': 'AAD / マネージド ID / サービス プリンシパルのサインイン失敗',
    'section.identity-permission.title': 'サービス プリンシパル オブジェクト / 権限の変更',
    'section.delete-disable.title': '削除 / 無効化操作',
    'section.suspicious-success.title': '疑わしい成功サインイン',
    'section.suspicious-ip.title': '疑わしい IP',
    'section.license.title': 'ライセンス使用量と残数',
    'section.shared-mailbox.title': '共有メールボックス',
    'section.dcr-log-errors.title': 'DCR ログ エラー',
    'section.intune-audit.title': 'Intune 監査レコード',
    'summary.failedSignins': 'サインイン失敗',
    'summary.deleteDisable': '削除 / 無効化',
    'summary.suspiciousIp': '疑わしい IP',
    'summary.suspiciousSigninSuccess': '信頼済み場所外の成功サインイン',
    'summary.mailboxLowSpace': 'メールボックス容量不足',
    'summary.sharedMailbox': '共有メールボックス',
    'summary.dcrLogErrors': 'DCR ログ エラー',
    'summary.intuneAudit': 'Intune 監査レコード',
    'field.次数': '件数', 'field.最后时间': '最終時刻', 'field.IP': 'IP', 'field.主体_应用摘要': '主体 / アプリ概要', 'field.说明': '説明',
    'field.表': 'テーブル', 'field.用户': 'ユーザー', 'field.操作': '操作', 'field.状态_原因': '状態 / 理由', 'field.时间': '時刻', 'field.操作者': '実行者',
    'field.用户身份': 'ID', 'field.应用名称': 'アプリ名', 'field.原因': '理由', 'field.应用': 'アプリ', 'field.SkuPartNumber': 'SkuPartNumber',
    'field.产品名称': '製品名', 'field.总数': '合計', 'field.已分配': '割り当て済み', 'field.剩余': '残数', 'field.状态': '状態',
    'field.用户名': 'ユーザー名', 'field.邮箱': 'メール', 'field.类型': '種類', 'field.总容量': '総容量', 'field.剩余容量': '残容量', 'field.使用率': '使用率', 'field.邮箱容量是否风险': 'メールボックス容量リスク',
    'field.输入流ID': '入力ストリーム ID', 'field.操作名称': '操作名', 'field.消息': 'メッセージ', 'field.活动时间': 'アクティビティ時刻', 'field.目标': '対象', 'field.权限': '権限',
    'field.结果_说明': '結果 / 説明', 'field.总记录数': '総レコード数', 'field.筛选后记录数': 'フィルター後レコード数', 'field.CSV': 'CSV', 'field.被删除者': '削除対象',
    'field.风险等级': 'リスクレベル', 'field.首次访问时间': '初回アクセス時刻', 'field.最近访问时间': '最終アクセス時刻'
  }
};
function applyLanguage(lang) {
  const dictionary = i18n[lang] || i18n['zh-CN'];
  document.documentElement.lang = lang;
  document.title = dictionary['report.title'] || document.title;
  document.querySelectorAll('[data-i18n]').forEach((element) => {
    const key = element.getAttribute('data-i18n');
    if (dictionary[key]) element.textContent = dictionary[key];
  });
  const selector = document.getElementById('languageSelect');
  if (selector) selector.value = lang;
  localStorage.setItem('logAnalyticsReportLanguage', lang);
}
const savedLanguage = localStorage.getItem('logAnalyticsReportLanguage') || 'zh-CN';
applyLanguage(savedLanguage);
document.getElementById('languageSelect')?.addEventListener('change', (event) => applyLanguage(event.target.value));
</script>
</body>
</html>
"@

$utf8Bom = [System.Text.UTF8Encoding]::new($true)
[System.IO.File]::WriteAllText($OutputPath, $html, $utf8Bom)
Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
