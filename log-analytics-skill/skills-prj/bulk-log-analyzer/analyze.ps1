param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [string]$AnalysisDate,

    [Parameter(Mandatory = $true)]
    [string]$TableName
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $ScriptDir 'log-analyzer-core.ps1')

# Generates self-contained HTML report from exported CSV
$csvPath = $CsvPath
$outputPath = $OutputPath
$analysisDate = $AnalysisDate
$reportTitle = "$TableName Report"
$sourceName = Split-Path -Leaf $csvPath

Write-Host "Loading CSV data..." -ForegroundColor Cyan
$data = @(Import-Csv -Path $csvPath -Encoding UTF8)
$totalEvents = $data.Count
Write-Host "Loaded $totalEvents records" -ForegroundColor Green

Write-Host "Computing statistics..." -ForegroundColor Cyan
$analysisProfile = Get-TableAnalysisProfile -TableName $TableName
$timelineTitleZh = '活动时间线'
$timelineTitleJa = 'アクティビティタイムライン'
$timelineNoteKey = ''
$timelineNoteZh = ''
$timelineNoteJa = ''
$clientIpEmptyKey = 'clientIpNoDataGeneric'
$statusNoteKey = ''
$statusNoteZh = ''
$statusNoteJa = ''
if ($TableName -eq 'AzureADUsersDCR_CL') {
    $timelineTitleZh = '用户目录快照时间'
    $timelineTitleJa = 'ユーザーディレクトリスナップショット時刻'
    $timelineNoteKey = 'timelineNoteAzureAD'
    $timelineNoteZh = 'AzureADUsersDCR_CL 是用户目录快照表；TimeGenerated 是本批数据写入 Log Analytics 的时间，不是用户真实活动时间。'
    $timelineNoteJa = 'AzureADUsersDCR_CL はユーザーディレクトリのスナップショットテーブルです。TimeGenerated は Log Analytics への取り込み時刻であり、実際のユーザー操作時刻ではありません。'
    $clientIpEmptyKey = 'clientIpEmptyAzureAD'
}
elseif ($TableName -eq 'AssignedLicensesDCR_CL') {
    $timelineTitleZh = '许可证快照时间'
    $timelineTitleJa = 'ライセンススナップショット時刻'
    $timelineNoteKey = 'timelineNoteAssignedLicenses'
    $timelineNoteZh = 'AssignedLicensesDCR_CL 是许可证分配快照表；TimeGenerated 是本批许可证状态写入 Log Analytics 的时间，不是用户真实活动时间。'
    $timelineNoteJa = 'AssignedLicensesDCR_CL はライセンス割り当てのスナップショットテーブルです。TimeGenerated はライセンス状態が Log Analytics に取り込まれた時刻であり、実際のユーザー操作時刻ではありません。'
    $clientIpEmptyKey = 'clientIpEmptyAssignedLicenses'
    $statusNoteKey = 'statusNoteAssignedLicenses'
    $statusNoteZh = 'AssignedLicensesDCR_CL 使用 ProvisioningStatus 判断状态：Success 计为成功，其他非空状态计为失败/需关注，空值计为未知。'
    $statusNoteJa = 'AssignedLicensesDCR_CL は ProvisioningStatus で状態を判定します。Success は成功、その他の空でない状態は失敗または確認対象、空値は不明として扱います。'
}
elseif ($TableName -eq 'MailboxStatisticsDCR_CL') {
    $timelineTitleZh = '邮箱统计快照时间'
    $timelineTitleJa = 'メールボックス統計スナップショット時刻'
    $timelineNoteKey = 'timelineNoteMailbox'
    $timelineNoteZh = 'MailboxStatisticsDCR_CL 是邮箱容量快照表；TimeGenerated 是本批统计数据写入 Log Analytics 的时间，不是邮箱用户活动时间。'
    $timelineNoteJa = 'MailboxStatisticsDCR_CL はメールボックス容量統計のスナップショットテーブルです。TimeGenerated は統計データが Log Analytics に取り込まれた時刻であり、メールボックス利用者の操作時刻ではありません。'
    $clientIpEmptyKey = 'clientIpEmptyMailbox'
}
elseif ($TableName -eq 'WQCLogDCR_CL') {
    $statusNoteKey = 'statusNoteWQC'
    $statusNoteZh = 'WQCLogDCR_CL 不提供单条记录的失败状态；成功/失败比率按已采集规则记录展示，真实日志类型来自 OperationType。'
    $statusNoteJa = 'WQCLogDCR_CL は各レコードの失敗状態を提供しません。成功/失敗比率は収集済みルールレコードとして表示し、実際のログ種別は OperationType から取得します。'
}

function Get-FieldValue {
    param(
        [object]$Row,
        [string[]]$Names,
        [string]$Default = 'Unknown'
    )

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

function Get-UserValue {
    param([object]$Row)

    $profile = $analysisProfile
    return Get-FieldValue -Row $Row -Names $profile.UserFields -Default $profile.DefaultUser
}

function Get-OperationValue {
    param([object]$Row)

    $profile = $analysisProfile
    if ($TableName -eq 'AzureADUsersDCR_CL') {
        $enabledRaw = (Get-FieldValue -Row $Row -Names @('accountEnabled', 'AccountEnabled') -Default '').ToLowerInvariant()
        $status = if ($enabledRaw -eq 'false') { 'Disabled Account' } elseif ($enabledRaw -eq 'true') { 'Enabled Account' } else { 'Account Status Unknown' }
        $department = Get-FieldValue -Row $Row -Names @('department', 'Department') -Default ''
        if ($department) { return "$status | Department: $department" }
        $company = Get-FieldValue -Row $Row -Names @('companyName', 'CompanyName') -Default ''
        if ($company) { return "$status | Company: $company" }
        $jobTitle = Get-FieldValue -Row $Row -Names @('jobTitle', 'JobTitle') -Default ''
        if ($jobTitle) { return "$status | JobTitle: $jobTitle" }
        return "$status | Department: Unassigned"
    }
    if ($TableName -eq 'AssignedLicensesDCR_CL') {
        $status = Get-FieldValue -Row $Row -Names @('ProvisioningStatus') -Default ''
        $servicePlan = Get-FieldValue -Row $Row -Names @('ServicePlanName', 'SkuPartNumber', 'LicenseName') -Default ''
        if ($status -and $servicePlan) { return "$status | $servicePlan" }
        if ($status) { return $status }
        if ($servicePlan) { return $servicePlan }
    }
    if (-not $profile.UseCompositeOperationGroup) {
        return Get-FieldValue -Row $Row -Names $profile.OperationFields -Default $profile.DefaultOperation
    }

    $parts = @()
    foreach ($field in $profile.GroupFields) {
        $value = Get-FieldValue -Row $Row -Names @($field) -Default ''
        if ($value) { $parts += $value }
    }

    if ($parts.Count -eq 0) {
        return Get-FieldValue -Row $Row -Names $profile.OperationFields -Default $profile.DefaultOperation
    }

    return ($parts -join ' | ')
}

function Get-WorkloadValue {
    param([object]$Row)

    $profile = $analysisProfile
    return Get-FieldValue -Row $Row -Names $profile.WorkloadFields -Default $profile.DefaultWorkload
}

function Get-ClientIpValue {
    param([object]$Row)

    $profile = $analysisProfile
    return Get-FieldValue -Row $Row -Names $profile.ClientIpFields -Default $profile.DefaultClientIp
}

function Get-SuccessValue {
    param([object]$Row)

    $profile = $analysisProfile
    $value = (Get-FieldValue -Row $Row -Names $profile.SuccessFields -Default $profile.DefaultSuccess).ToLowerInvariant()
    if ($TableName -eq 'AssignedLicensesDCR_CL') {
        if ([string]::IsNullOrWhiteSpace($value) -or $value -eq 'unknown') { return 'unknown' }
        if ($value -eq 'success') { return 'true' }
        return 'false'
    }
    if ($value -match '^(true|success|succeeded|delivered|expanded|completed|complete|ok|pass|passed|0)$') { return 'true' }
    if ($value -match '^(false|fail|failed|failure|undelivered|blocked|rejected|denied|error|timeout|quarantined|1)$') { return 'false' }
    if ($profile.DefaultSuccess -eq 'true') { return 'true' }
    return 'unknown'
}

function Test-OperationContains {
    param(
        [object]$Row,
        [string[]]$Operations
    )

    $operation = Get-OperationValue -Row $Row
    foreach ($op in $Operations) {
        if ($operation -eq $op -or $operation -like "*| $op |*" -or $operation -like "*$op*") {
            return $true
        }
    }

    return $false
}

function Test-UsableIpValue {
    param([string]$IP)

    if ([string]::IsNullOrWhiteSpace($IP)) { return $false }
    $normalized = $IP.Trim()
    if ($normalized -in @('Unknown', 'N/A', '-', '0.0.0.0', '::', '::1', '127.0.0.1', '255.255.255.255')) { return $false }
    return $true
}

function Get-SafeCount {
    param([object]$Value)

    if ($null -eq $Value) { return 0 }
    return @($Value).Count
}

function Add-UserDisplayNameMapping {
    param(
        [hashtable]$Map,
        [object]$Row
    )

    $displayName = Get-FieldValue -Row $Row -Names @('displayName', 'DisplayName') -Default ''
    if (-not $displayName) { return }

    $identifiers = @(
        (Get-FieldValue -Row $Row -Names @('userPrincipalName', 'UserPrincipalName', 'UserUPN', 'UPN') -Default ''),
        (Get-FieldValue -Row $Row -Names @('mail', 'Mail', 'EmailAddress') -Default '')
    )

    foreach ($identifier in $identifiers) {
        if ($identifier -and $identifier -match '@') {
            $key = $identifier.ToLowerInvariant()
            if (-not $Map.ContainsKey($key)) {
                $Map[$key] = $displayName
            }
        }
    }
}

function Get-UserDisplayNameMap {
    param(
        [array]$CurrentData,
        [string]$CsvPath,
        [string]$TableName
    )

    $map = @{}
    if ($TableName -eq 'AzureADUsersDCR_CL') {
        foreach ($row in $CurrentData) {
            Add-UserDisplayNameMapping -Map $map -Row $row
        }
    }

    $csvDir = Split-Path -Parent $CsvPath
    $candidateFiles = @()
    if (Test-Path $csvDir) {
        $candidateFiles += @(Get-ChildItem -Path $csvDir -Filter 'AzureADUsersDCR_CL*.csv' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 3)
        $cacheDir = Join-Path $csvDir 'cache'
        if (Test-Path $cacheDir) {
            $candidateFiles += @(Get-ChildItem -Path $cacheDir -Filter 'AzureADUsersDCR_CL*.csv' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 3)
        }
    }

    foreach ($file in $candidateFiles) {
        try {
            foreach ($row in @(Import-Csv -Path $file.FullName -Encoding UTF8)) {
                Add-UserDisplayNameMapping -Map $map -Row $row
            }
        } catch {}
    }

    return $map
}

function Format-UserDisplayValue {
    param([string]$User)

    if (-not $User) { return $User }
    $emailPattern = '[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}'
    return [regex]::Replace($User, $emailPattern, {
        param($match)
        $email = $match.Value
        $key = $email.ToLowerInvariant()
        if ($script:userDisplayNameMap.ContainsKey($key)) {
            $name = $script:userDisplayNameMap[$key]
            if ($User -like "$name ($email)*") { return $email }
            return "$name ($email)"
        }
        return $email
    }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

$script:userDisplayNameMap = Get-UserDisplayNameMap -CurrentData $data -CsvPath $csvPath -TableName $TableName

# Unique users
$allUsersList = [System.Collections.Generic.List[string]]::new()
foreach ($row in $data) {
    $u = Format-UserDisplayValue -User (Get-UserValue -Row $row)
    $allUsersList.Add($u) | Out-Null
}
$allUsers = $allUsersList.ToArray()
$uniqueUsers = ($allUsers | Select-Object -Unique).Count

# Unique operations
$allOps = @($data | ForEach-Object { Get-OperationValue -Row $_ })
$uniqueOps = ($allOps | Select-Object -Unique).Count

# Workload distribution
$workloadMap = @{}
foreach ($row in $data) {
    $wl = Get-WorkloadValue -Row $row
    $workloadMap[$wl] = ($workloadMap[$wl] + 1)
}

# Top users
$userMap = @{}
foreach ($u in $allUsers) {
    $userMap[$u] = ($userMap[$u] + 1)
}
$topUsers = $userMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 15

# Top operations
$opMap = @{}
foreach ($o in $allOps) {
    $opName = if ($o) { $o } else { 'Unknown' }
    $opMap[$opName] = ($opMap[$opName] + 1)
}
$topOps = $opMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 15

# Top ClientIPs
$ipMap = @{}
if ($TableName -ne 'AssignedLicensesDCR_CL') {
    foreach ($row in $data) {
        $ip = Get-ClientIpValue -Row $row
        if (-not (Test-UsableIpValue -IP $ip)) { continue }
        $ipMap[$ip] = ($ipMap[$ip] + 1)
    }
}
$topIPs = $ipMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10

# Success/Failure
$successCount = 0
$failCount = 0
$unknownCount = 0
foreach ($row in $data) {
    $s = Get-SuccessValue -Row $row
    if ($s -eq 'true') { $successCount++ }
    elseif ($s -eq 'false') { $failCount++ }
    else { $unknownCount++ }
}
if ($unknownCount -gt 0 -and -not $statusNoteKey) {
    $statusNoteKey = 'statusUnknownNote'
    $statusNoteZh = '未知表示源日志没有提供 IsSuccess、ResultStatus、Status 或 Result 等可判断成功/失败的字段，或该事件类型本身没有成功/失败语义。'
    $statusNoteJa = '不明は、元ログに IsSuccess、ResultStatus、Status、Result など成功/失敗を判断できるフィールドがない、またはそのイベント種別自体に成功/失敗の意味がないことを示します。'
}

# Activity timeline (by hour)
$hourMap = @{}
foreach ($row in $data) {
    $tg = $row.TimeGenerated
    if ($tg -and $tg -ne '') {
        try {
            $dt = [DateTime]::Parse($tg)
            $h = $dt.ToString('yyyy-MM-dd HH:00')
            $hourMap[$h] = ($hourMap[$h] + 1)
        } catch {}
    }
}
$timelineSorted = $hourMap.GetEnumerator() | Sort-Object Name

# Off-hours activity (00:00-07:00)
$offHoursEventsList = [System.Collections.Generic.List[object]]::new()
if ($TableName -ne 'AssignedLicensesDCR_CL') {
    foreach ($row in $data) {
        $tg = $row.TimeGenerated
        if ($tg -and $tg -ne '') {
            try {
                $dt = [DateTime]::Parse($tg)
                $localDt = $dt.ToLocalTime()
                if ($localDt.Hour -ge 22 -and $localDt.Hour -lt 8) {
                    $offHoursEventsList.Add($row) | Out-Null
                }
            } catch {}
        }
    }
}
$offHoursEvents = $offHoursEventsList.ToArray()

# Failed operations details
$failedByOp = @{}
$failedEventCount = 0
foreach ($row in $data) {
    if ((Get-SuccessValue -Row $row) -eq 'false') {
        $failedEventCount++
        $op = Get-OperationValue -Row $row
        $failedByOp[$op] = ($failedByOp[$op] + 1)
    }
}
$failedByOp = $failedByOp.GetEnumerator() | Sort-Object Value -Descending

# High-privilege operations
$highPrivOps = @('ExportReport', 'Search', 'EditDataset', 'Delete', 'DeleteDataset', 'DeleteReport', 'DeleteWorkspace', 'AdminAction')
$highPrivEvents = if ($TableName -eq 'AssignedLicensesDCR_CL') { @() } else { @($data | Where-Object { Test-OperationContains -Row $_ -Operations $highPrivOps }) }
$highPrivByUser = @{}
foreach ($row in $highPrivEvents) {
    $u = Format-UserDisplayValue -User (Get-UserValue -Row $row)
    $key = "$u | $(Get-OperationValue -Row $row)"
    $highPrivByUser[$key] = ($highPrivByUser[$key] + 1)
}
$highPrivByUser = $highPrivByUser.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 20

# Sensitive data events
$sensitiveOps = @('SensitivityLabeledFileOpened', 'SensitivityLabeledFileRenamed', 'IrmContent', 'AppliedSensitivityLabel', 'ChangedSensitivityLabel')
$sensitiveEvents = if ($TableName -eq 'AssignedLicensesDCR_CL') { @() } else { @($data | Where-Object { Test-OperationContains -Row $_ -Operations $sensitiveOps }) }
$sensitiveByOp = @{}
foreach ($row in $sensitiveEvents) {
    $op = Get-OperationValue -Row $row
    $sensitiveByOp[$op] = ($sensitiveByOp[$op] + 1)
}
$sensitiveByOp = $sensitiveByOp.GetEnumerator() | Sort-Object Value -Descending

# Service account activity (GUIDs starting with 0000-...)
$serviceAcctEvents = if ($TableName -eq 'AssignedLicensesDCR_CL') { @() } else { @($data | Where-Object { (Get-UserValue -Row $_) -match '^00000009-' }) }
$serviceAcctByOp = @{}
foreach ($row in $serviceAcctEvents) {
    $op = Get-OperationValue -Row $row
    $serviceAcctByOp[$op] = ($serviceAcctByOp[$op] + 1)
}
$serviceAcctByOp = $serviceAcctByOp.GetEnumerator() | Sort-Object Value -Descending

# IP velocity - single IP with multiple users
$ipUsers = @{}
if ($TableName -ne 'AssignedLicensesDCR_CL') {
    foreach ($row in $data) {
        $ip = Get-ClientIpValue -Row $row
        if (-not (Test-UsableIpValue -IP $ip)) { continue }
        $u = Format-UserDisplayValue -User (Get-UserValue -Row $row)
        if (-not $ipUsers.ContainsKey($ip)) { $ipUsers[$ip] = @{} }
        $ipUsers[$ip][$u] = 1
    }
}
$ipVelocity = @()
foreach ($ip in $ipUsers.Keys) {
    $count = $ipUsers[$ip].Count
    if ($count -gt 5) {
        $ipVelocity += [PSCustomObject]@{
            IP = $ip
            UserCount = $count
            Users = ($ipUsers[$ip].Keys -join ', ')
        }
    }
}
$ipVelocity = $ipVelocity | Sort-Object UserCount -Descending

# Suspicious IPs (non-RFC1918 accessing multiple workloads)
$ipWorkloads = @{}
if ($TableName -ne 'AssignedLicensesDCR_CL') {
    foreach ($row in $data) {
        $ip = Get-ClientIpValue -Row $row
        if (-not (Test-UsableIpValue -IP $ip)) { continue }
        # Skip RFC1918
        if ($ip -match '^(10\.|172\.(1[6-9]|2[0-9]|3[01])\.|192\.168\.)') { continue }
        $wl = Get-WorkloadValue -Row $row
        if (-not $ipWorkloads.ContainsKey($ip)) { $ipWorkloads[$ip] = @{} }
        $ipWorkloads[$ip][$wl] = 1
    }
}
$suspiciousIPs = @()
foreach ($ip in $ipWorkloads.Keys) {
    $wlCount = $ipWorkloads[$ip].Count
    if ($wlCount -gt 1) {
        $suspiciousIPs += [PSCustomObject]@{
            IP = $ip
            WorkloadCount = $wlCount
            Workloads = ($ipWorkloads[$ip].Keys -join ', ')
        }
    }
}
$suspiciousIPs = $suspiciousIPs | Sort-Object WorkloadCount -Descending

# Off-hours by user
$offHoursByUser = @{}
foreach ($row in $offHoursEvents) {
    $u = Format-UserDisplayValue -User (Get-UserValue -Row $row)
    $offHoursByUser[$u] = ($offHoursByUser[$u] + 1)
}
$offHoursByUser = $offHoursByUser.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 15

Write-Host "All statistics computed." -ForegroundColor Green

# ============================================================
# Build HTML Report
# ============================================================

Write-Host "Generating HTML report..." -ForegroundColor Cyan

# Helper: escape HTML
function EscapeHtml {
    param([string]$text)
    if (-not $text) { return '' }
    return $text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&#39;'
}

function ConvertTo-JsJsonLiteral {
    param([string]$Json)
    return [System.Uri]::EscapeDataString($Json)
}

# Build glossary (only ops present in data)
$glossaryOps = @{}
$opNames = $opMap.Keys | Sort-Object
foreach ($op in $opNames) {
    $expZh = $op  # default fallback
    $expJa = $op
    # Known glossary entries
    if ($TableName -eq 'AzureADUsersDCR_CL' -and $op -like 'Enabled Account*') {
        $expZh = '已启用用户账户（accountEnabled=true）；后缀表示该账号的部门、公司或职务分类。'
        $expJa = '有効なユーザーアカウント（accountEnabled=true）。末尾は部門、会社、または役職の分類を示します。'
    }
    elseif ($TableName -eq 'AzureADUsersDCR_CL' -and $op -like 'Disabled Account*') {
        $expZh = '已禁用用户账户（accountEnabled=false）；后缀表示该账号的部门、公司或职务分类。'
        $expJa = '無効なユーザーアカウント（accountEnabled=false）。末尾は部門、会社、または役職の分類を示します。'
    }
    else {
    switch ($op) {
        'ViewReport'        { $expZh = '查看 PowerBI 报表'; $expJa = 'PowerBI レポートを閲覧' }
        'GetWorkspaces'     { $expZh = '获取工作区列表'; $expJa = 'ワークスペース一覧を取得' }
        'RefreshDataset'    { $expZh = '刷新数据集'; $expJa = 'データセットを更新' }
        'ExportReport'      { $expZh = '导出报表（有风险）'; $expJa = 'レポートをエクスポート（リスクあり）' }
        'EditDataset'       { $expZh = '编辑数据集'; $expJa = 'データセットを編集' }
        'Search'            { $expZh = '执行搜索'; $expJa = '検索を実行' }
        'Import'            { $expZh = '导入内容'; $expJa = 'コンテンツをインポート' }
        'MessageSend'       { $expZh = '发送Teams消息'; $expJa = 'Teamsメッセージを送信' }
        'SensitivityLabeledFileOpened'  { $expZh = '打开敏感标签文件'; $expJa = '機密ラベル付きファイルを開く' }
        'SensitivityLabeledFileRenamed' { $expZh = '重命名敏感标签文件'; $expJa = '機密ラベル付きファイルの名前を変更' }
        'MessageReadReceiptReceived'    { $expZh = '收到已读回执'; $expJa = '既読確認を受信' }
        'GetSnapshots'      { $expZh = '获取快照'; $expJa = 'スナップショットを取得' }
        'RunEmailSubscription' { $expZh = '运行邮件订阅'; $expJa = 'メールサブスクリプションを実行' }
        'ApiEndpointCallEvent' { $expZh = 'API端点调用'; $expJa = 'APIエンドポイント呼び出し' }
        'UpdateSharingPermission' { $expZh = '更新共享权限'; $expJa = '共有権限を更新' }
        'ShareReport'       { $expZh = '共享报表'; $expJa = 'レポートを共有' }
        'CreateDataset'     { $expZh = '创建数据集'; $expJa = 'データセットを作成' }
        'CreateReport'      { $expZh = '创建报表'; $expJa = 'レポートを作成' }
        'Publish'           { $expZh = '发布内容'; $expJa = 'コンテンツを公開' }
        default             { $expZh = $op; $expJa = $op }
    }
    }
    $glossaryOps[$op] = @{ zh = $expZh; ja = $expJa; count = $opMap[$op] }
}

# Workload glossary
$wlGlossary = @{
    'PowerBI' = @{ zh = 'Power BI 报表平台'; ja = 'Power BI レポートプラットフォーム' }
    'MicrosoftTeams' = @{ zh = 'Microsoft Teams'; ja = 'Microsoft Teams' }
    'SecurityComplianceCenter' = @{ zh = '安全与合规中心'; ja = 'セキュリティコンプライアンスセンター' }
    'OneDrive' = @{ zh = 'OneDrive'; ja = 'OneDrive' }
    'SharePoint' = @{ zh = 'SharePoint'; ja = 'SharePoint' }
    'Exchange' = @{ zh = 'Exchange'; ja = 'Exchange' }
    'AzureActiveDirectory' = @{ zh = 'Azure Active Directory'; ja = 'Azure Active Directory' }
    'PowerPlatform' = @{ zh = 'Power Platform'; ja = 'Power Platform' }
}

# Chart colors (cycle through)
$chartColors = @('#58a6ff', '#3fb950', '#bc8cff', '#f0883e', '#f85149', '#39d2c0', '#58a6ff', '#3fb950', '#bc8cff', '#f0883e', '#f85149', '#39d2c0', '#58a6ff', '#3fb950', '#bc8cff')

# ---- Build JSON data structures ----

function ToJsonArray {
    param([hashtable]$map)
    $rows = @($map.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
        [PSCustomObject]@{ name = $_.Name; value = $_.Value }
    })
    return ConvertTo-Json -InputObject $rows -Compress -Depth 4
}

function ToSortedJsonArray {
    param([array]$items, [string]$nameKey = 'Name', [string]$valueKey = 'Value')
    $rows = @($items | Where-Object { $null -ne $_ -and $null -ne $_.$nameKey -and $null -ne $_.$valueKey } | ForEach-Object {
        [PSCustomObject]@{ name = $_.$nameKey; value = $_.$valueKey }
    })
    return ConvertTo-Json -InputObject $rows -Compress -Depth 4
}

function ToKeyValueJsonArray {
    param([array]$items, [string]$keyName = 'Name', [string]$valName = 'Value')
    $rows = @($items | Where-Object { $null -ne $_ -and $null -ne $_.$keyName -and $null -ne $_.$valName } | ForEach-Object {
        [PSCustomObject]@{ key = $_.$keyName; value = $_.$valName }
    })
    return ConvertTo-Json -InputObject $rows -Compress -Depth 4
}

$topUsersJson = ToSortedJsonArray $topUsers
$topUsersJsonJs = ConvertTo-JsJsonLiteral -Json $topUsersJson
$topOpsJson = ToSortedJsonArray $topOps
$topOpsJsonJs = ConvertTo-JsJsonLiteral -Json $topOpsJson
$operationGroupBreakdown = $opMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 25
$operationGroupBreakdownJson = ToSortedJsonArray $operationGroupBreakdown
$operationGroupBreakdownJsonJs = ConvertTo-JsJsonLiteral -Json $operationGroupBreakdownJson
$topIPsJson = ToSortedJsonArray $topIPs
$topIPsJsonJs = ConvertTo-JsJsonLiteral -Json $topIPsJson
$timelineJson = ToKeyValueJsonArray $timelineSorted 'Name' 'Value'
$timelineJsonJs = ConvertTo-JsJsonLiteral -Json $timelineJson
$workloadJson = ToJsonArray $workloadMap
$workloadJsonJs = ConvertTo-JsJsonLiteral -Json $workloadJson

# Failed events JSON
$failedOpsJson = ToSortedJsonArray -items $failedByOp -nameKey 'Name' -valueKey 'Value'
$failedOpsJsonJs = ConvertTo-JsJsonLiteral -Json $failedOpsJson

# High-priv events JSON
$highPrivJson = ToSortedJsonArray -items $highPrivByUser -nameKey 'Name' -valueKey 'Value'
$highPrivJsonJs = ConvertTo-JsJsonLiteral -Json $highPrivJson

# Sensitive events JSON
$sensitiveJson = ToSortedJsonArray -items $sensitiveByOp -nameKey 'Name' -valueKey 'Value'
$sensitiveJsonJs = ConvertTo-JsJsonLiteral -Json $sensitiveJson

# Service accounts JSON
$serviceAcctJson = ToSortedJsonArray -items $serviceAcctByOp -nameKey 'Name' -valueKey 'Value'
$serviceAcctJsonJs = ConvertTo-JsJsonLiteral -Json $serviceAcctJson

# Off-hours JSON
$offHoursJson = ToSortedJsonArray -items $offHoursByUser
$offHoursJsonJs = ConvertTo-JsJsonLiteral -Json $offHoursJson
$suspiciousIPsJson = ToKeyValueJsonArray -items $suspiciousIPs -keyName 'IP' -valName 'Workloads'
$suspiciousIPsJsonJs = ConvertTo-JsJsonLiteral -Json $suspiciousIPsJson
$ipVelocityJson = ToKeyValueJsonArray -items $ipVelocity -keyName 'IP' -valName 'UserCount'
$ipVelocityJsonJs = ConvertTo-JsJsonLiteral -Json $ipVelocityJson

# Glossary JSON
$glossaryJsonObject = @{}
foreach ($op in $glossaryOps.Keys | Sort-Object) {
    $glossaryJsonObject[$op] = [PSCustomObject]@{
        zh = $glossaryOps[$op].zh
        ja = $glossaryOps[$op].ja
        count = $glossaryOps[$op].count
    }
}
$glossaryJson = ConvertTo-Json -InputObject $glossaryJsonObject -Compress -Depth 5
$glossaryJsonJs = ConvertTo-JsJsonLiteral -Json $glossaryJson

# Workload glossary JSON
$wlGlossaryJsonObject = @{}
foreach ($wl in $wlGlossary.Keys | Sort-Object) {
    if ($workloadMap.ContainsKey($wl)) {
        $wlGlossaryJsonObject[$wl] = [PSCustomObject]@{
            zh = $wlGlossary[$wl].zh
            ja = $wlGlossary[$wl].ja
        }
    }
}
$wlGlossaryJson = ConvertTo-Json -InputObject $wlGlossaryJsonObject -Compress -Depth 5
$wlGlossaryJsonJs = ConvertTo-JsJsonLiteral -Json $wlGlossaryJson

# Build data table - first 1000 rows
$tableRows = ''
$previewRows = [Math]::Min(1000, $data.Count)
for ($i = 0; $i -lt $previewRows; $i++) {
    $row = $data[$i]
    $tg = if ($row.TimeGenerated) { $row.TimeGenerated } else { '' }
    $user = Format-UserDisplayValue -User (Get-UserValue -Row $row)
    $op = Get-OperationValue -Row $row
    $wl = Get-WorkloadValue -Row $row
    $ip = Get-ClientIpValue -Row $row
    $success = Get-SuccessValue -Row $row
    $tableRows += "<tr><td>$i</td><td>$(EscapeHtml $tg)</td><td class='op-cell' data-op='$(EscapeHtml $op)'>$(EscapeHtml $op)</td><td>$(EscapeHtml $user)</td><td>$(EscapeHtml $wl)</td><td>$(EscapeHtml $ip)</td><td class='status-$(EscapeHtml $success)'>$(EscapeHtml $success)</td></tr>`n"
}

# Determine risk count
$riskCount = 0
if ($failCount -gt 0) { $riskCount++ }
if ((Get-SafeCount $suspiciousIPs) -gt 0) { $riskCount++ }
if ((Get-SafeCount $offHoursEvents) -gt 0) { $riskCount++ }
if ((Get-SafeCount $highPrivEvents) -gt 0) { $riskCount++ }
if ((Get-SafeCount $sensitiveEvents) -gt 0) { $riskCount++ }
if ((Get-SafeCount $ipVelocity) -gt 0) { $riskCount++ }

# Determine if we should show risk section
$showRiskSection = 'true'
$riskSectionDisplay = 'block'
$groupBreakdownDisplay = if ($analysisProfile.UseCompositeOperationGroup) { 'block' } else { 'none' }
$highPrivCount = Get-SafeCount $highPrivEvents
$sensitiveCount = Get-SafeCount $sensitiveEvents
$serviceAcctCount = Get-SafeCount $serviceAcctEvents
$offHoursCount = Get-SafeCount $offHoursEvents
$suspiciousCount = Get-SafeCount $suspiciousIPs
$ipVelocityCount = Get-SafeCount $ipVelocity

$html = @"
<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>$reportTitle - $analysisDate</title>
<style>
:root {
  --bg-primary: #ffffff; --bg-secondary: #ffffff; --bg-tertiary: #f6f8fa;
  --border: #d0d7de; --text-primary: #111111; --text-secondary: #4b5563;
  --accent: #0969da; --accent-green: #116329; --accent-red: #cf222e;
  --accent-yellow: #9a6700; --accent-purple: #8250df; --accent-orange: #bc4c00;
  --accent-cyan: #0550ae;
}
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans SC', 'Noto Sans JP', Helvetica, Arial, sans-serif; background: var(--bg-primary); color: var(--text-primary); padding: 24px; line-height: 1.6; }
.navbar { display: flex; justify-content: space-between; align-items: center; padding: 12px 0; border-bottom: 1px solid var(--border); margin-bottom: 24px; }
.navbar h2 { font-size: 14px; color: var(--text-secondary); }
.lang-toggle { display: flex; gap: 8px; }
.lang-toggle button { background: var(--bg-tertiary); color: var(--text-primary); border: 1px solid var(--border); padding: 6px 14px; border-radius: 8px; cursor: pointer; font-size: 13px; }
.lang-toggle button.active { background: var(--accent); color: #fff; border-color: var(--accent); }
.lang-toggle button:hover:not(.active) { background: var(--border); }
.header { background: var(--bg-secondary); border-radius: 8px; padding: 24px; margin-bottom: 24px; border: 1px solid var(--border); }
.header h1 { font-size: 24px; margin-bottom: 8px; }
.header .subtitle { color: var(--text-secondary); font-size: 14px; margin-bottom: 16px; }
.meta-tags { display: flex; flex-wrap: wrap; gap: 8px; }
.meta-tag { background: var(--bg-tertiary); padding: 4px 12px; border-radius: 20px; font-size: 12px; color: var(--text-secondary); border: 1px solid var(--border); }
.summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px; }
.summary-card { background: var(--bg-secondary); border: 1px solid var(--border); border-radius: 8px; padding: 20px; }
.summary-card .label { font-size: 13px; color: var(--text-secondary); margin-bottom: 8px; }
.summary-card .value { font-size: 28px; font-weight: 700; }
.summary-card .value.green { color: var(--accent-green); }
.summary-card .value.red { color: var(--accent-red); }
.summary-card .value.blue { color: var(--accent); }
.summary-card .value.yellow { color: var(--accent-yellow); }
.summary-card .value.purple { color: var(--accent-purple); }
.section { background: var(--bg-secondary); border: 1px solid var(--border); border-radius: 8px; padding: 24px; margin-bottom: 24px; }
.section h2 { font-size: 18px; margin-bottom: 16px; padding-bottom: 12px; border-bottom: 1px solid var(--border); }
.bar-chart { margin-bottom: 8px; }
.bar-item { display: flex; align-items: center; margin-bottom: 4px; font-size: 13px; }
.bar-label { width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex-shrink: 0; padding-right: 12px; color: var(--text-secondary); cursor: help; }
.bar-container { flex: 1; background: var(--bg-tertiary); border-radius: 4px; height: 22px; position: relative; min-width: 60px; }
.bar-fill { height: 100%; border-radius: 4px; display: flex; align-items: center; padding-left: 8px; font-size: 11px; color: #ffffff; min-width: 30px; transition: width 0.3s ease; }
.bar-count { position: absolute; right: 8px; top: 50%; transform: translateY(-50%); font-size: 11px; color: var(--text-secondary); }
.timeline-bar { display: flex; align-items: center; margin-bottom: 3px; font-size: 12px; }
.timeline-label { width: 130px; color: var(--text-secondary); flex-shrink: 0; }
.timeline-container { flex: 1; background: var(--bg-tertiary); border-radius: 3px; height: 18px; }
.timeline-fill { height: 100%; background: var(--accent); border-radius: 3px; opacity: 0.7; }
.donut-container { display: flex; justify-content: center; align-items: center; gap: 32px; flex-wrap: wrap; }
.donut-svg { width: 180px; height: 180px; }
.donut-legend { display: flex; flex-direction: column; gap: 6px; }
.legend-item { display: flex; align-items: center; gap: 8px; font-size: 13px; }
.legend-color { width: 12px; height: 12px; border-radius: 3px; flex-shrink: 0; }
.seg-bar { height: 24px; border-radius: 6px; overflow: hidden; display: flex; margin-bottom: 12px; }
.seg-fill { height: 100%; display: flex; align-items: center; justify-content: center; font-size: 11px; color: #ffffff; }
.seg-legend { display: flex; gap: 16px; flex-wrap: wrap; }
.seg-legend-item { display: flex; align-items: center; gap: 6px; font-size: 13px; color: var(--text-secondary); }
table { width: 100%; border-collapse: collapse; font-size: 12px; }
thead { position: sticky; top: 0; z-index: 10; }
th { background: var(--bg-tertiary); padding: 10px 8px; text-align: left; border-bottom: 2px solid var(--border); cursor: pointer; user-select: none; color: var(--text-secondary); font-size: 11px; text-transform: uppercase; }
th:hover { color: var(--text-primary); }
td { padding: 8px; border-bottom: 1px solid var(--border); }
tr:hover td { background: var(--bg-tertiary); }
.table-wrapper { overflow-x: auto; }
.table-scroll { max-height: 500px; overflow-y: auto; }
.pagination { display: flex; justify-content: center; align-items: center; gap: 12px; margin-top: 12px; font-size: 13px; }
.pagination button { background: var(--bg-tertiary); color: var(--text-primary); border: 1px solid var(--border); padding: 4px 12px; border-radius: 4px; cursor: pointer; }
.pagination button:disabled { opacity: 0.4; cursor: default; }
.risk-table { width: 100%; margin-bottom: 16px; font-size: 13px; }
.risk-table th { background: var(--bg-tertiary); }
.risk-badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
.risk-high { background: rgba(248,81,73,0.15); color: var(--accent-red); }
.risk-medium { background: rgba(240,136,62,0.15); color: var(--accent-orange); }
.risk-low { background: rgba(210,153,34,0.15); color: var(--accent-yellow); }
.risk-subsection { margin-bottom: 24px; }
.risk-subsection h3 { font-size: 15px; margin-bottom: 12px; color: var(--accent-red); }
.tooltip-box { position: fixed; background: var(--bg-secondary); border: 1px solid var(--border); border-radius: 6px; padding: 10px 14px; font-size: 12px; z-index: 1000; pointer-events: none; max-width: 350px; box-shadow: 0 8px 24px rgba(0,0,0,0.15); display: none; }
.tooltip-box .tip-cn { color: var(--accent); margin-bottom: 4px; }
.tooltip-box .tip-jp { color: var(--accent-purple); }
.tooltip-box .tip-op { color: var(--text-secondary); font-size: 11px; margin-bottom: 6px; }
.glossary-section { background: var(--bg-secondary); border: 1px solid var(--border); border-radius: 8px; padding: 24px; margin-bottom: 24px; }
.glossary-toggle { background: var(--bg-tertiary); color: var(--text-primary); border: 1px solid var(--border); padding: 8px 20px; border-radius: 8px; cursor: pointer; font-size: 14px; margin-bottom: 0; }
.glossary-toggle:hover { background: var(--border); }
.glossary-content { display: none; margin-top: 16px; }
.glossary-content.visible { display: block; }
.status-true { color: var(--accent-green); }
.status-false { color: var(--accent-red); }
</style>
</head>
<body>
<div class="tooltip-box" id="tooltip">
  <div class="tip-op" id="tip-op"></div>
  <div class="tip-cn" id="tip-cn"></div>
  <div class="tip-jp" id="tip-jp"></div>
</div>

<div class="navbar">
  <h2>$TableName</h2>
  <div class="lang-toggle">
    <button class="active" onclick="switchLang('zh')">中文</button>
    <button onclick="switchLang('ja')">日本語</button>
  </div>
</div>

<div class="header">
  <h1>$reportTitle</h1>
  <div class="subtitle" data-i18n="subtitle">Log Analytics 日志分析报告</div>
  <div class="meta-tags">
    <span class="meta-tag">查询时间段: $analysisDate</span>
    <span class="meta-tag">Total Records: $totalEvents</span>
    <span class="meta-tag">Source: $sourceName</span>
  </div>
</div>

<div class="section" id="risk-section" style="display: $riskSectionDisplay;">
  <h2 data-i18n="riskAnalysis">风险分析</h2>
  <p style="color:var(--accent-red);margin-bottom:16px;font-size:14px;" data-i18n="riskIndicators">$riskCount 个风险指标已检出</p>
  <div id="risk-content"></div>
</div>

<div class="summary-grid">
  <div class="summary-card">
    <div class="label" data-i18n="totalEvents">总事件数</div>
    <div class="value blue">$totalEvents</div>
  </div>
  <div class="summary-card">
    <div class="label" data-i18n="uniqueUsers">唯一用户</div>
    <div class="value purple">$uniqueUsers</div>
  </div>
  <div class="summary-card">
    <div class="label" data-i18n="uniqueOps">唯一操作</div>
    <div class="value yellow">$uniqueOps</div>
  </div>
  <div class="summary-card">
    <div class="label" data-i18n="workloads">工作负载</div>
    <div class="value green">$($workloadMap.Count)</div>
  </div>
  <div class="summary-card">
    <div class="label" data-i18n="success">成功</div>
    <div class="value green">$successCount</div>
  </div>
  <div class="summary-card">
    <div class="label" data-i18n="failed">失败</div>
    <div class="value red">$failCount</div>
  </div>
</div>

<div class="section">
  <h2 data-i18n="activityTimeline">活动时间线</h2>
  <p style="color:var(--text-secondary);font-size:13px;margin-bottom:12px;" data-i18n="$timelineNoteKey">$timelineNoteZh</p>
  <div id="timeline-chart"></div>
</div>

<div class="section">
  <h2 data-i18n="workloadDist">工作负载分布</h2>
  <div id="donut-chart"></div>
</div>

<div class="section">
  <h2 data-i18n="topUsers">活跃用户排行</h2>
  <div id="users-chart" class="bar-chart"></div>
</div>

<div class="section">
  <h2 data-i18n="topOps">操作类型排行</h2>
  <div id="ops-chart" class="bar-chart"></div>
</div>

<div class="section" id="group-breakdown-section" style="display: $groupBreakdownDisplay;">
  <h2 data-i18n="groupBreakdown">Activity / Operation / Workload 分组排行</h2>
  <p style="color:var(--text-secondary);font-size:13px;margin-bottom:12px;">AuditGeneralDCR_CL 和 SharePointAuditDCR_CL 使用 Activity + Operation + Workload 组合分组。</p>
  <div id="group-breakdown-chart" class="bar-chart"></div>
</div>

<div class="section">
  <h2 data-i18n="topIPs">客户端 IP 排行</h2>
  <div id="ips-chart" class="bar-chart"></div>
</div>

<div class="section">
  <h2 data-i18n="successRatio">成功/失败比率</h2>
  <div id="success-ratio"></div>
  <p style="color:var(--text-secondary);font-size:13px;margin-top:12px;" data-i18n="$statusNoteKey">$statusNoteZh</p>
</div>

<div class="glossary-section">
  <button class="glossary-toggle" onclick="toggleGlossary()" data-i18n="showGlossary">显示术语表</button>
  <div class="glossary-content" id="glossary-content">
    <table class="risk-table">
      <thead><tr>
        <th>Operation (EN)</th>
        <th>说明 (CN)</th>
        <th>説明 (JP)</th>
        <th>Count</th>
      </tr></thead>
      <tbody id="glossary-body"></tbody>
    </table>
  </div>
</div>

<div class="section">
  <h2 data-i18n="detailedTable">详细数据</h2>
  <p style="color:var(--text-secondary);font-size:13px;margin-bottom:12px;" data-i18n="tablePreview">预览前 1000 行</p>
  <div class="table-wrapper">
    <div class="table-scroll">
      <table id="data-table">
        <thead>
          <tr>
            <th onclick="sortTable(0)">#</th>
            <th onclick="sortTable(1)" data-i18n="time">时间</th>
            <th onclick="sortTable(2)" data-i18n="operation">操作</th>
            <th onclick="sortTable(3)" data-i18n="user">用户</th>
            <th onclick="sortTable(4)" data-i18n="workload">工作负载</th>
            <th onclick="sortTable(5)" data-i18n="clientIP">IP</th>
            <th onclick="sortTable(6)" data-i18n="status">状态</th>
          </tr>
        </thead>
        <tbody id="table-body">
$tableRows
        </tbody>
      </table>
    </div>
  </div>
  <div class="pagination">
    <button id="prev-btn" onclick="prevPage()" disabled data-i18n="previous">上一页</button>
    <span id="page-info">1 / 1</span>
    <button id="next-btn" onclick="nextPage()" disabled data-i18n="next">下一页</button>
  </div>
</div>

<script>
// ===== i18n =====
let currentLang = 'zh';
const i18n = {
  zh: {
    "totalEvents":"总事件数","uniqueUsers":"唯一用户","uniqueOps":"唯一操作","workloads":"工作负载",
    "success":"成功","failed":"失败","activityTimeline":"$timelineTitleZh","workloadDist":"工作负载分布",
    "topUsers":"活跃用户排行","topOps":"操作类型排行","groupBreakdown":"Activity / Operation / Workload 分组排行","topIPs":"客户端 IP 排行",
    "successRatio":"成功/失败比率","riskAnalysis":"风险分析","detailedTable":"详细数据",
    "showGlossary":"显示术语表","hideGlossary":"隐藏术语表","metric":"指标","value":"值",
    "severity":"严重程度","unknown":"未知","previous":"上一页","next":"下一页",
    "subtitle":"Log Analytics 日志分析报告","tablePreview":"预览前 1000 行",
    "time":"时间","operation":"操作","user":"用户","workload":"工作负载","clientIP":"IP",
    "status":"状态","riskIndicators":"$riskCount 个风险指标已检出",
    "failedOps":"失败操作","suspiciousIPs":"可疑 IP","offHours":"非工作时间活动",
    "highPrivOps":"高权限操作","sensitiveData":"敏感数据事件","ipVelocity":"IP 多用户关联",
    "serviceAccounts":"服务账户活动","low":"低风险","medium":"中风险","high":"高风险",
    "offHoursUsers":"非工作时活跃用户","failedOpSummary":"失败操作汇总","count":"次数",
    "timelineNoteAzureAD":"AzureADUsersDCR_CL 是用户目录快照表；TimeGenerated 是本批数据写入 Log Analytics 的时间，不是用户真实活动时间。",
    "timelineNoteAssignedLicenses":"AssignedLicensesDCR_CL 是许可证分配快照表；TimeGenerated 是本批许可证状态写入 Log Analytics 的时间，不是用户真实活动时间。",
    "timelineNoteMailbox":"MailboxStatisticsDCR_CL 是邮箱容量快照表；TimeGenerated 是本批统计数据写入 Log Analytics 的时间，不是邮箱用户活动时间。",
    "clientIpNoDataGeneric":"无可用客户端 IP 数据。",
    "clientIpEmptyAzureAD":"此表不包含客户端 IP 字段；它是用户目录快照数据，不记录登录或访问来源 IP。",
    "clientIpEmptyAssignedLicenses":"此表不包含客户端 IP 字段；它是许可证分配快照数据，不记录登录或访问来源 IP。",
    "clientIpEmptyMailbox":"此表不包含客户端 IP 字段；它是邮箱容量统计快照，不记录客户端访问来源 IP。",
    "statusUnknownNote":"未知表示源日志没有提供 IsSuccess、ResultStatus、Status 或 Result 等可判断成功/失败的字段，或该事件类型本身没有成功/失败语义。",
    "statusNoteAssignedLicenses":"AssignedLicensesDCR_CL 使用 ProvisioningStatus 判断状态：Success 计为成功，其他非空状态计为失败/需关注，空值计为未知。",
    "statusNoteWQC":"WQCLogDCR_CL 不提供单条记录的失败状态；成功/失败比率按已采集规则记录展示，真实日志类型来自 OperationType。"
  },
  ja: {
    "totalEvents":"総イベント数","uniqueUsers":"ユニークユーザー","uniqueOps":"ユニーク操作","workloads":"ワークロード",
    "success":"成功","failed":"失敗","activityTimeline":"$timelineTitleJa","workloadDist":"ワークロード分布",
    "topUsers":"アクティブユーザーランキング","topOps":"操作タイプランキング","groupBreakdown":"Activity / Operation / Workload グループランキング","topIPs":"クライアント IP ランキング",
    "successRatio":"成功/失敗比率","riskAnalysis":"リスク分析","detailedTable":"詳細データ",
    "showGlossary":"用語集を表示","hideGlossary":"用語集を非表示","metric":"指標","value":"値",
    "severity":"重要度","unknown":"不明","previous":"前へ","next":"次へ",
    "subtitle":"Log Analytics ログ分析レポート","tablePreview":"最初の 1000 行をプレビュー",
    "time":"時間","operation":"操作","user":"ユーザー","workload":"ワークロード","clientIP":"IP",
    "status":"ステータス","riskIndicators":"$riskCount 件のリスク指標が検出されました",
    "failedOps":"失敗した操作","suspiciousIPs":"不審な IP","offHours":"時間外のアクティビティ",
    "highPrivOps":"高権限操作","sensitiveData":"機密データイベント","ipVelocity":"IP 複数ユーザー",
    "serviceAccounts":"サービスアカウント","low":"低リスク","medium":"中リスク","high":"高リスク",
    "offHoursUsers":"時間外アクティブユーザー","failedOpSummary":"失敗操作まとめ","count":"回数",
    "timelineNoteAzureAD":"AzureADUsersDCR_CL はユーザーディレクトリのスナップショットテーブルです。TimeGenerated は Log Analytics への取り込み時刻であり、実際のユーザー操作時刻ではありません。",
    "timelineNoteAssignedLicenses":"AssignedLicensesDCR_CL はライセンス割り当てのスナップショットテーブルです。TimeGenerated はライセンス状態が Log Analytics に取り込まれた時刻であり、実際のユーザー操作時刻ではありません。",
    "timelineNoteMailbox":"MailboxStatisticsDCR_CL はメールボックス容量統計のスナップショットテーブルです。TimeGenerated は統計データが Log Analytics に取り込まれた時刻であり、メールボックス利用者の操作時刻ではありません。",
    "clientIpNoDataGeneric":"利用可能なクライアント IP データはありません。",
    "clientIpEmptyAzureAD":"このテーブルにはクライアント IP フィールドがありません。ユーザーディレクトリのスナップショットであり、ログインやアクセス元 IP は記録されません。",
    "clientIpEmptyAssignedLicenses":"このテーブルにはクライアント IP フィールドがありません。ライセンス割り当てのスナップショットであり、ログインやアクセス元 IP は記録されません。",
    "clientIpEmptyMailbox":"このテーブルにはクライアント IP フィールドがありません。メールボックス容量統計のスナップショットであり、クライアントアクセス元 IP は記録されません。",
    "statusUnknownNote":"不明は、元ログに IsSuccess、ResultStatus、Status、Result など成功/失敗を判断できるフィールドがない、またはそのイベント種別自体に成功/失敗の意味がないことを示します。",
    "statusNoteAssignedLicenses":"AssignedLicensesDCR_CL は ProvisioningStatus で状態を判定します。Success は成功、その他の空でない状態は失敗または確認対象、空値は不明として扱います。",
    "statusNoteWQC":"WQCLogDCR_CL は各レコードの失敗状態を提供しません。成功/失敗比率は収集済みルールレコードとして表示し、実際のログ種別は OperationType から取得します。"
  }
};

function switchLang(lang) {
  currentLang = lang;
  document.querySelectorAll('[data-i18n]').forEach(el => {
    const key = el.getAttribute('data-i18n');
    if (i18n[lang] && i18n[lang][key]) {
      el.textContent = i18n[lang][key];
    }
  });
  document.querySelectorAll('.lang-toggle button').forEach(btn => {
    btn.classList.remove('active');
    if ((lang === 'zh' && btn.textContent === '中文') || (lang === 'ja' && btn.textContent === '日本語')) {
      btn.classList.add('active');
    }
  });
  refreshGlossary();
  buildRiskSection();
}

// ===== Glossary =====
const glossaryData = JSON.parse(decodeURIComponent('$glossaryJsonJs'));
const wlGlossaryData = JSON.parse(decodeURIComponent('$wlGlossaryJsonJs'));
let glossaryVisible = false;

function toggleGlossary() {
  glossaryVisible = !glossaryVisible;
  const el = document.getElementById('glossary-content');
  const btn = document.querySelector('.glossary-toggle');
  if (glossaryVisible) {
    el.classList.add('visible');
    btn.textContent = i18n[currentLang]['hideGlossary'] || '隐藏术语表';
    refreshGlossary();
  } else {
    el.classList.remove('visible');
    btn.textContent = i18n[currentLang]['showGlossary'] || '显示术语表';
  }
}

function refreshGlossary() {
  const tbody = document.getElementById('glossary-body');
  if (!tbody) return;
  let html = '';
  const keys = Object.keys(glossaryData).sort((a, b) => glossaryData[b].count - glossaryData[a].count);
  keys.forEach(op => {
    const g = glossaryData[op];
    const desc = currentLang === 'zh' ? g.zh : g.ja;
    html += '<tr><td>' + op + '</td><td>' + desc + '</td><td>' + (currentLang === 'zh' ? g.ja : g.zh) + '</td><td>' + g.count + '</td></tr>';
  });
  tbody.innerHTML = html;
}

// ===== Tooltip =====
const tooltip = document.getElementById('tooltip');
function showTooltip(e, op) {
  const g = glossaryData[op];
  if (!g) return;
  document.getElementById('tip-op').textContent = op;
  document.getElementById('tip-cn').textContent = g.zh;
  document.getElementById('tip-jp').textContent = g.ja;
  tooltip.style.display = 'block';
  const x = Math.min(e.pageX + 10, window.innerWidth - 360);
  const y = Math.min(e.pageY + 10, window.innerHeight - 100);
  tooltip.style.left = x + 'px';
  tooltip.style.top = y + 'px';
}
function hideTooltip() { tooltip.style.display = 'none'; }
document.addEventListener('mouseover', function(e) {
  const cell = e.target.closest('.op-cell');
  if (cell) showTooltip(e, cell.getAttribute('data-op'));
});
document.addEventListener('mouseout', function(e) {
  if (e.target.closest('.op-cell')) hideTooltip();
});

// ===== Render Charts =====
const chartColors = ['#58a6ff','#3fb950','#bc8cff','#f0883e','#f85149','#39d2c0'];

function renderBarChart(containerId, data, maxItems, emptyMessageKey) {
  const container = document.getElementById(containerId);
  const messageKey = emptyMessageKey || 'clientIpNoDataGeneric';
  const message = (i18n[currentLang] && i18n[currentLang][messageKey]) || 'No data';
  if (!container || !data || data.length === 0) { container.innerHTML = '<p style="color:var(--text-secondary)" data-i18n="' + messageKey + '">' + message + '</p>'; return; }
  const items = data.slice(0, maxItems);
  const maxVal = items.length > 0 ? items[0].value : 1;
  let html = '';
  items.forEach((item, i) => {
    const pct = Math.max((item.value / maxVal) * 100, 2);
    const color = chartColors[i % chartColors.length];
    const g = glossaryData[item.name];
    const tooltipAttr = g ? 'data-tooltip="true"' : '';
    html += '<div class="bar-item">';
    html += '<div class="bar-label op-cell" data-op="' + item.name + '" style="cursor:help">' + item.name + '</div>';
    html += '<div class="bar-container"><div class="bar-fill" style="width:' + pct + '%;background:' + color + '">' + item.value + '</div></div>';
    html += '</div>';
  });
  container.innerHTML = html;
}

function renderCompositeBreakdown(containerId, data) {
  const container = document.getElementById(containerId);
  if (!container || !data || data.length === 0) { container.innerHTML = '<p style="color:var(--text-secondary)">No data</p>'; return; }
  let html = '';
  data.slice(0, 25).forEach((item, i) => {
    const color = chartColors[i % chartColors.length];
    const maxVal = data[0] && data[0].value ? data[0].value : 1;
    const pct = Math.max((item.value / maxVal) * 100, 2);
    html += '<div class="bar-item">';
    html += '<div class="bar-label" style="width:320px">' + item.name + '</div>';
    html += '<div class="bar-container"><div class="bar-fill" style="width:' + pct + '%;background:' + color + '">' + item.value + '</div></div>';
    html += '</div>';
  });
  container.innerHTML = html;
}

function renderDonut(containerId, data) {
  const container = document.getElementById(containerId);
  if (!container || !data || data.length === 0) { container.innerHTML = '<p style="color:var(--text-secondary)">No data</p>'; return; }
  const total = data.reduce((s, d) => s + d.value, 0);
  const r = 70;
  const circumference = 2 * Math.PI * r;
  let offset = 0;
  let circles = '';
  let legend = '';
  data.forEach((item, i) => {
    const pct = item.value / total;
    const dash = pct * circumference;
    const gap = circumference - dash;
    const color = chartColors[i % chartColors.length];
    circles += '<circle cx="90" cy="90" r="' + r + '" fill="none" stroke="' + color + '" stroke-width="24" stroke-dasharray="' + dash + ' ' + gap + '" stroke-dashoffset="' + (-offset) + '" transform="rotate(-90 90 90)"/>';
    offset += dash;
    const wl = wlGlossaryData[item.name];
    const wlLabel = wl ? (currentLang === 'zh' ? wl.zh : wl.ja) : item.name;
    legend += '<div class="legend-item"><div class="legend-color" style="background:' + color + '"></div><span>' + wlLabel + ' (' + item.value + ')</span></div>';
  });
  let svgHtml = '<div class="donut-container">';
  svgHtml += '<svg class="donut-svg" viewBox="0 0 180 180">' + circles + '<text x="90" y="90" text-anchor="middle" dominant-baseline="central" fill="var(--text-primary)" font-size="18" font-weight="600">' + total + '</text></svg>';
  svgHtml += '<div class="donut-legend">' + legend + '</div></div>';
  container.innerHTML = svgHtml;
}

function renderTimeline(containerId, data) {
  const container = document.getElementById(containerId);
  if (!container || !data || data.length === 0) { container.innerHTML = '<p style="color:var(--text-secondary)">No data</p>'; return; }
  const maxVal = data.reduce((m, d) => Math.max(m, d.value), 0);
  let html = '';
  data.forEach(item => {
    const pct = maxVal > 0 ? (item.value / maxVal) * 100 : 0;
    const label = item.key.split(' ')[1] || item.key;
    html += '<div class="timeline-bar"><div class="timeline-label">' + label + '</div><div class="timeline-container"><div class="timeline-fill" style="width:' + pct + '%"></div></div></div>';
  });
  container.innerHTML = html;
}

function renderSuccessRatio(containerId, success, fail, unknown) {
  const container = document.getElementById(containerId);
  const total = success + fail + unknown;
  if (total === 0) { container.innerHTML = '<p style="color:var(--text-secondary)">No data</p>'; return; }
  const sPct = ((success / total) * 100).toFixed(1);
  const fPct = ((fail / total) * 100).toFixed(1);
  const uPct = ((unknown / total) * 100).toFixed(1);
  let html = '<div class="seg-bar">';
  if (success > 0) html += '<div class="seg-fill" style="width:' + sPct + '%;background:var(--accent-green)">' + sPct + '%</div>';
  if (fail > 0) html += '<div class="seg-fill" style="width:' + fPct + '%;background:var(--accent-red)">' + fPct + '%</div>';
  if (unknown > 0) html += '<div class="seg-fill" style="width:' + uPct + '%;background:var(--text-secondary)">' + uPct + '%</div>';
  html += '</div><div class="seg-legend">';
  html += '<div class="seg-legend-item"><div class="legend-color" style="background:var(--accent-green)"></div>成功 (' + success + ')</div>';
  html += '<div class="seg-legend-item"><div class="legend-color" style="background:var(--accent-red)"></div>失败 (' + fail + ')</div>';
  if (unknown > 0) html += '<div class="seg-legend-item"><div class="legend-color" style="background:var(--text-secondary)"></div>未知 (' + unknown + ')</div>';
  html += '</div>';
  container.innerHTML = html;
}

// ===== Table Pagination =====
const rowsPerPage = 50;
let currentPage = 1;
function initPagination() {
  const total = document.querySelectorAll('#table-body tr').length;
  const totalPages = Math.ceil(total / rowsPerPage) || 1;
  showPage(1, totalPages);
  window._totalPages = totalPages;
}
function showPage(page, total) {
  currentPage = page;
  const totalRows = document.querySelectorAll('#table-body tr').length;
  const start = (page - 1) * rowsPerPage;
  const rows = document.querySelectorAll('#table-body tr');
  rows.forEach((row, i) => { row.style.display = (i >= start && i < start + rowsPerPage) ? '' : 'none'; });
  document.getElementById('page-info').textContent = page + ' / ' + total;
  document.getElementById('prev-btn').disabled = (page <= 1);
  document.getElementById('next-btn').disabled = (page >= total);
}
function prevPage() { if (currentPage > 1) showPage(currentPage - 1, window._totalPages); }
function nextPage() { if (currentPage < window._totalPages) showPage(currentPage + 1, window._totalPages); }

// ===== Table Sort =====
let sortDir = {};
function sortTable(colIdx) {
  const tbody = document.getElementById('table-body');
  const rows = Array.from(tbody.querySelectorAll('tr'));
  sortDir[colIdx] = !sortDir[colIdx];
  const dir = sortDir[colIdx] ? 1 : -1;
  rows.sort((a, b) => {
    const aText = a.cells[colIdx]?.textContent || '';
    const bText = b.cells[colIdx]?.textContent || '';
    const aNum = parseFloat(aText);
    const bNum = parseFloat(bText);
    if (!isNaN(aNum) && !isNaN(bNum)) return (aNum - bNum) * dir;
    return aText.localeCompare(bText) * dir;
  });
  rows.forEach(r => tbody.appendChild(r));
}

// ===== Risk Section =====
const riskData = {
  failedOps: JSON.parse(decodeURIComponent('$failedOpsJsonJs')),
  highPriv: JSON.parse(decodeURIComponent('$highPrivJsonJs')),
  sensitive: JSON.parse(decodeURIComponent('$sensitiveJsonJs')),
  serviceAcct: JSON.parse(decodeURIComponent('$serviceAcctJsonJs')),
  offHoursUsers: JSON.parse(decodeURIComponent('$offHoursJsonJs')),
  suspiciousIPs: JSON.parse(decodeURIComponent('$suspiciousIPsJsonJs')),
  ipVelocity: JSON.parse(decodeURIComponent('$ipVelocityJsonJs')),
  failCount: $failCount,
  highPrivCount: $highPrivCount,
  sensitiveCount: $sensitiveCount,
  serviceAcctCount: $serviceAcctCount,
  offHoursCount: $offHoursCount,
  suspiciousCount: $suspiciousCount,
  ipVelocityCount: $ipVelocityCount
};

function buildRiskSection() {
  const container = document.getElementById('risk-content');
  if (!container) return;
  let html = '';

  // Failed operations
  if (riskData.failCount > 0) {
    html += '<div class="risk-subsection"><h3>' + (i18n[currentLang]['failedOps'] || '失败操作') + '</h3>';
    html += '<table class="risk-table"><thead><tr><th>Operation</th><th>' + (i18n[currentLang]['count'] || '次数') + '</th></tr></thead><tbody>';
    riskData.failedOps.slice(0, 10).forEach(item => {
      html += '<tr><td class="op-cell" data-op="' + item.name + '">' + item.name + '</td><td class="status-false">' + item.value + '</td></tr>';
    });
    html += '</tbody></table></div>';
  }

  // Off-hours activity
  if (riskData.offHoursCount > 0) {
    html += '<div class="risk-subsection"><h3>' + (i18n[currentLang]['offHours'] || '非工作时间活动') + '</h3>';
    html += '<table class="risk-table"><thead><tr><th>' + (i18n[currentLang]['user'] || '用户') + '</th><th>Events</th></tr></thead><tbody>';
    riskData.offHoursUsers.slice(0, 10).forEach(item => {
      const sev = item.value > 20 ? 'high' : item.value > 5 ? 'medium' : 'low';
      html += '<tr><td>' + item.name + '</td><td><span class="risk-badge risk-' + sev + '">' + item.value + '</span></td></tr>';
    });
    html += '</tbody></table></div>';
  }

  // High-priv operations
  if (riskData.highPrivCount > 0) {
    html += '<div class="risk-subsection"><h3>' + (i18n[currentLang]['highPrivOps'] || '高权限操作') + '</h3>';
    html += '<table class="risk-table"><thead><tr><th>User | Operation</th><th>Count</th></tr></thead><tbody>';
    riskData.highPriv.slice(0, 15).forEach(item => {
      html += '<tr><td class="op-cell" data-op="' + item.name + '">' + item.name + '</td><td>' + item.value + '</td></tr>';
    });
    html += '</tbody></table></div>';
  }

  // Sensitive data events
  if (riskData.sensitiveCount > 0) {
    html += '<div class="risk-subsection"><h3>' + (i18n[currentLang]['sensitiveData'] || '敏感数据事件') + '</h3>';
    html += '<table class="risk-table"><thead><tr><th>Operation</th><th>Count</th></tr></thead><tbody>';
    riskData.sensitive.forEach(item => {
      html += '<tr><td class="op-cell" data-op="' + item.name + '">' + item.name + '</td><td>' + item.value + '</td></tr>';
    });
    html += '</tbody></table></div>';
  }

  // Suspicious IPs
  if (riskData.suspiciousCount > 0) {
    html += '<div class="risk-subsection"><h3>' + (i18n[currentLang]['suspiciousIPs'] || '可疑 IP') + '</h3>';
    html += '<table class="risk-table"><thead><tr><th>IP</th><th>Workloads</th></tr></thead><tbody>';
    riskData.suspiciousIPs.slice(0, 10).forEach(item => {
      html += '<tr><td>' + item.key + '</td><td>' + item.value + '</td></tr>';
    });
    html += '</tbody></table></div>';
  }

  // IP velocity
  if (riskData.ipVelocityCount > 0) {
    html += '<div class="risk-subsection"><h3>' + (i18n[currentLang]['ipVelocity'] || 'IP 多用户关联') + '</h3>';
    html += '<table class="risk-table"><thead><tr><th>IP</th><th>User Count</th></tr></thead><tbody>';
    riskData.ipVelocity.slice(0, 10).forEach(item => {
      html += '<tr><td>' + item.key + '</td><td><span class="risk-badge risk-high">' + item.value + '</span></td></tr>';
    });
    html += '</tbody></table></div>';
  }

  // Service accounts
  if (riskData.serviceAcctCount > 0) {
    html += '<div class="risk-subsection"><h3>' + (i18n[currentLang]['serviceAccounts'] || '服务账户活动') + '</h3>';
    html += '<table class="risk-table"><thead><tr><th>Operation</th><th>Count</th></tr></thead><tbody>';
    riskData.serviceAcct.slice(0, 15).forEach(item => {
      html += '<tr><td class="op-cell" data-op="' + item.name + '">' + item.name + '</td><td>' + item.value + '</td></tr>';
    });
    html += '</tbody></table></div>';
  }

  if (html === '') {
    html = '<p style="color:var(--accent-green)">No risk indicators detected.</p>';
  }
  container.innerHTML = html;
}

// ===== Init =====
document.addEventListener('DOMContentLoaded', function() {
  renderTimeline('timeline-chart', JSON.parse(decodeURIComponent('$timelineJsonJs')));
  renderDonut('donut-chart', JSON.parse(decodeURIComponent('$workloadJsonJs')));
  renderBarChart('users-chart', JSON.parse(decodeURIComponent('$topUsersJsonJs')), 15);
  renderBarChart('ops-chart', JSON.parse(decodeURIComponent('$topOpsJsonJs')), 15);
  renderCompositeBreakdown('group-breakdown-chart', JSON.parse(decodeURIComponent('$operationGroupBreakdownJsonJs')));
  renderBarChart('ips-chart', JSON.parse(decodeURIComponent('$topIPsJsonJs')), 10, '$clientIpEmptyKey');
  renderSuccessRatio('success-ratio', $successCount, $failCount, $unknownCount);
  initPagination();
  buildRiskSection();
});
</script>
</body>
</html>
"@

$html | Out-File -FilePath $outputPath -Encoding UTF8
Write-Host "Report saved to: $outputPath" -ForegroundColor Green
