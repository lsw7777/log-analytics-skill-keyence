# Azure Log Bulk Analyzer
# Generates self-contained HTML report from exported CSV

param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\General_$(Get-Date -Format 'yyyyMMdd').csv",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\report_$(Get-Date -Format 'yyyyMMdd_HHmm').html",

    [Parameter(Mandatory = $false)]
    [string]$AnalysisDate = "$(Get-Date -Format 'yyyy-MM-dd')"
)

# ============================================================
# Table Schema Mapping - defines how to extract success/failure
# status from different DCR_CL tables
# ============================================================
$TableSchemas = @{
    # AuditGeneralDCR_CL: has IsSuccess field
    'AuditGeneralDCR_CL' = @{
        SuccessField = 'IsSuccess'
        SuccessValue = 'true'
        FailValue = 'false'
        UserField = 'UserUPN'
        UserFallback = 'UserId'
        OpField = 'Operation'
        WlField = 'Workload'
        IpField = 'ClientIP'
        TimeField = 'TimeGenerated'
        ShortName = 'General'
        DisplayName = 'Audit General'
        Description = 'Microsoft 365 通用审计日志，记录用户和管理员在Microsoft 365服务中的操作，包括Power BI、Teams、Exchange、SharePoint等'
        DescriptionJa = 'Microsoft 365 汎用監査ログ。Power BI、Teams、Exchange、SharePointなどのサービスにおけるユーザーおよび管理者の操作を記録'
        Fields = @{
            'TimeGenerated' = @{ zh = '日志生成时间'; ja = 'ログ生成時間'; en = 'Log generation time' }
            'UserUPN' = @{ zh = '用户主体名称'; ja = 'ユーザープリンシパル名'; en = 'User Principal Name' }
            'UserId' = @{ zh = '用户ID'; ja = 'ユーザーID'; en = 'User ID' }
            'Operation' = @{ zh = '操作类型'; ja = '操作タイプ'; en = 'Operation type' }
            'Workload' = @{ zh = '工作负载/服务'; ja = 'ワークロード/サービス'; en = 'Workload/Service' }
            'ClientIP' = @{ zh = '客户端IP地址'; ja = 'クライアントIPアドレス'; en = 'Client IP address' }
            'IsSuccess' = @{ zh = '是否成功'; ja = '成功かどうか'; en = 'Whether operation succeeded' }
            'RecordType' = @{ zh = '记录类型'; ja = 'レコードタイプ'; en = 'Record type' }
            'ObjectId' = @{ zh = '操作对象'; ja = '操作対象'; en = 'Operation target' }
        }
    }
    # SharePointAuditDCR_CL: no IsSuccess field, use ResultStatus
    'SharePointAuditDCR_CL' = @{
        SuccessField = 'ResultStatus'
        SuccessValue = '0'
        FailValue = '1'
        UserField = 'UserId'
        UserFallback = 'UserKey'
        OpField = 'Operation'
        WlField = 'Workload'
        IpField = 'ClientIP'
        TimeField = 'TimeGenerated'
        ShortName = 'SPAudit'
        DisplayName = 'SharePoint Audit'
        Description = 'SharePoint Online 审计日志，记录SharePoint站点和文档的访问、修改、共享等操作'
        DescriptionJa = 'SharePoint Online 監査ログ。SharePointサイトおよびドキュメントへのアクセス、変更、共有などの操作を記録'
        Fields = @{
            'TimeGenerated' = @{ zh = '日志生成时间'; ja = 'ログ生成時間'; en = 'Log generation time' }
            'UserId' = @{ zh = '用户ID'; ja = 'ユーザーID'; en = 'User ID' }
            'UserKey' = @{ zh = '用户密钥'; ja = 'ユーザーキー'; en = 'User Key' }
            'Operation' = @{ zh = '操作类型'; ja = '操作タイプ'; en = 'Operation type' }
            'Workload' = @{ zh = '工作负载'; ja = 'ワークロード'; en = 'Workload' }
            'ClientIP' = @{ zh = '客户端IP地址'; ja = 'クライアントIPアドレス'; en = 'Client IP address' }
            'ResultStatus' = @{ zh = '结果状态 (0=成功, 1=失败)'; ja = '結果ステータス (0=成功, 1=失敗)'; en = 'Result status (0=success, 1=failure)' }
            'SourceFileName' = @{ zh = '源文件名'; ja = 'ソースファイル名'; en = 'Source file name' }
            'SiteUrl' = @{ zh = '站点URL'; ja = 'サイトURL'; en = 'Site URL' }
        }
    }
    # MessageTraceDataDCR_CL: use Status field
    'MessageTraceDataDCR_CL' = @{
        SuccessField = 'Status'
        SuccessValue = 'Delivered'
        FailValue = 'Failed'
        UserField = 'SenderAddress'
        UserFallback = 'RecipientAddress'
        OpField = 'Status'
        WlField = ''
        IpField = 'FromIP'
        TimeField = 'TimeGenerated'
        ShortName = 'MsgTrace'
        DisplayName = 'Message Trace'
        Description = 'Exchange Online 邮件追踪数据，记录邮件的发送、接收、传递状态和路径信息'
        DescriptionJa = 'Exchange Online メール追跡データ。メールの送信、受信、配信状態と経路情報を記録'
        Fields = @{
            'TimeGenerated' = @{ zh = '日志生成时间'; ja = 'ログ生成時間'; en = 'Log generation time' }
            'SenderAddress' = @{ zh = '发件人地址'; ja = '送信者アドレス'; en = 'Sender address' }
            'RecipientAddress' = @{ zh = '收件人地址'; ja = '受信者アドレス'; en = 'Recipient address' }
            'Status' = @{ zh = '传递状态 (Delivered/Failed)'; ja = '配信状態 (Delivered/Failed)'; en = 'Delivery status' }
            'Subject' = @{ zh = '邮件主题'; ja = 'メール件名'; en = 'Email subject' }
            'MessageId' = @{ zh = '邮件ID'; ja = 'メールID'; en = 'Message ID' }
            'FromIP' = @{ zh = '发件人IP地址'; ja = '送信者IPアドレス'; en = 'Sender IP address' }
            'Received' = @{ zh = '接收时间'; ja = '受信時間'; en = 'Received time' }
            'TenantId' = @{ zh = '租户ID'; ja = 'テナントID'; en = 'Tenant ID' }
        }
    }
    # AssignedLicensesDCR_CL: no success/failure concept
    'AssignedLicensesDCR_CL' = @{
        SuccessField = ''
        SuccessValue = ''
        FailValue = ''
        UserField = 'UserPrincipalName'
        UserFallback = ''
        OpField = 'DisplayName'
        WlField = ''
        IpField = ''
        TimeField = 'TimeGenerated'
        ShortName = 'Licenses'
        DisplayName = 'Assigned Licenses'
        Description = 'Microsoft 365 许可证分配记录，记录用户被分配的许可证类型和分配时间'
        DescriptionJa = 'Microsoft 365 ライセンス割り当て記録。ユーザーに割り当てられたライセンスタイプと割り当て時間を記録'
        Fields = @{
            'TimeGenerated' = @{ zh = '日志生成时间'; ja = 'ログ生成時間'; en = 'Log generation time' }
            'UserPrincipalName' = @{ zh = '用户主体名称'; ja = 'ユーザープリンシパル名'; en = 'User Principal Name' }
            'DisplayName' = @{ zh = '许可证名称'; ja = 'ライセンス名'; en = 'License display name' }
            'SkuPartNumber' = @{ zh = '许可证SKU'; ja = 'ライセンスSKU'; en = 'License SKU' }
            'ServicePlanName' = @{ zh = '服务计划名称'; ja = 'サービスプラン名'; en = 'Service plan name' }
            'ProvisioningStatus' = @{ zh = '配置状态'; ja = 'プロビジョニング状態'; en = 'Provisioning status' }
            'AppliesTo' = @{ zh = '适用对象'; ja = '適用対象'; en = 'Applies to' }
        }
    }
    # AzureADUsersDCR_CL: no success/failure concept
    'AzureADUsersDCR_CL' = @{
        SuccessField = ''
        SuccessValue = ''
        FailValue = ''
        UserField = 'userPrincipalName'
        UserFallback = 'displayName'
        OpField = 'department'
        WlField = ''
        IpField = ''
        TimeField = 'TimeGenerated'
        ShortName = 'AADUsers'
        DisplayName = 'Azure AD Users'
        Description = 'Azure Active Directory 用户信息，记录用户属性、角色和状态变更'
        DescriptionJa = 'Azure Active Directory ユーザー情報。ユーザー属性、ロール、状態の変更を記録'
        Fields = @{
            'TimeGenerated' = @{ zh = '日志生成时间'; ja = 'ログ生成時間'; en = 'Log generation time' }
            'userPrincipalName' = @{ zh = '用户主体名称'; ja = 'ユーザープリンシパル名'; en = 'User Principal Name' }
            'displayName' = @{ zh = '显示名称'; ja = '表示名'; en = 'Display name' }
            'mail' = @{ zh = '邮箱地址'; ja = 'メールアドレス'; en = 'Mail address' }
            'department' = @{ zh = '部门'; ja = '部門'; en = 'Department' }
            'jobTitle' = @{ zh = '职位'; ja = '職位'; en = 'Job title' }
            'companyName' = @{ zh = '公司名称'; ja = '会社名'; en = 'Company name' }
            'officeLocation' = @{ zh = '办公位置'; ja = 'オフィス所在地'; en = 'Office location' }
            'employeeId' = @{ zh = '员工ID'; ja = '従業員ID'; en = 'Employee ID' }
            'accountEnabled' = @{ zh = '账户启用状态'; ja = 'アカウント有効状態'; en = 'Account enabled status' }
            'businessPhones' = @{ zh = '商务电话'; ja = 'ビジネス電話'; en = 'Business phones' }
        }
    }
    # MailboxStatisticsDCR_CL: no success/failure concept
    'MailboxStatisticsDCR_CL' = @{
        SuccessField = ''
        SuccessValue = ''
        FailValue = ''
        UserField = 'UserPrincipalName'
        UserFallback = 'DisplayName'
        OpField = 'RecipientTypeDetails'
        WlField = ''
        IpField = ''
        TimeField = 'TimeGenerated'
        ShortName = 'Mailbox'
        DisplayName = 'Mailbox Statistics'
        Description = 'Exchange 邮箱统计信息，记录邮箱大小、项目数量、最后访问时间等'
        DescriptionJa = 'Exchange メールボックス統計情報。メールボックスサイズ、アイテム数、最終アクセス時間などを記録'
        Fields = @{
            'TimeGenerated' = @{ zh = '日志生成时间'; ja = 'ログ生成時間'; en = 'Log generation time' }
            'UserPrincipalName' = @{ zh = '用户主体名称'; ja = 'ユーザープリンシパル名'; en = 'User Principal Name' }
            'DisplayName' = @{ zh = '显示名称'; ja = '表示名'; en = 'Display name' }
            'EmailAddress' = @{ zh = '邮箱地址'; ja = 'メールアドレス'; en = 'Email address' }
            'RecipientTypeDetails' = @{ zh = '收件人类型详情'; ja = '受信者タイプの詳細'; en = 'Recipient type details' }
            'TotalItemSize' = @{ zh = '总项目大小'; ja = '総アイテムサイズ'; en = 'Total item size' }
            'TotalDeletedItemSize' = @{ zh = '总删除项目大小'; ja = '総削除アイテムサイズ'; en = 'Total deleted item size' }
            'AvailableSpaceGB' = @{ zh = '可用空间(GB)'; ja = '利用可能領域(GB)'; en = 'Available space (GB)' }
            'QuotaLimitGB' = @{ zh = '配额限制(GB)'; ja = 'クォータ制限(GB)'; en = 'Quota limit (GB)' }
        }
    }
    # WQCLogDCR_CL: use OperationType field
    'WQCLogDCR_CL' = @{
        SuccessField = ''
        SuccessValue = ''
        FailValue = ''
        UserField = 'CurrentUsername'
        UserFallback = 'CurrentMail'
        OpField = 'OperationType'
        WlField = ''
        IpField = ''
        TimeField = 'TimeGenerated'
        ShortName = 'WQC'
        DisplayName = 'WQC Log'
        Description = 'WQC (Workplace Quality Control) 日志，记录质量检查操作和结果'
        DescriptionJa = 'WQC (Workplace Quality Control) ログ。品質チェック操作と結果を記録'
        Fields = @{
            'TimeGenerated' = @{ zh = '日志生成时间'; ja = 'ログ生成時間'; en = 'Log generation time' }
            'CurrentMail' = @{ zh = '当前用户邮箱'; ja = '現在のユーザーメール'; en = 'Current user email' }
            'CurrentUsername' = @{ zh = '当前用户名'; ja = '現在のユーザー名'; en = 'Current username' }
            'OperationType' = @{ zh = '操作类型'; ja = '操作タイプ'; en = 'Operation type' }
            'InboxRuleName' = @{ zh = '收件箱规则名'; ja = '受信トレイルール名'; en = 'Inbox rule name' }
            'ExceptIfFrom' = @{ zh = '排除发件人'; ja = '除外送信元'; en = 'Exclude from sender' }
            'ExceptIfSentTo' = @{ zh = '排除收件人'; ja = '除外送信先'; en = 'Exclude to recipient' }
            'ForwardtoMail' = @{ zh = '转发邮箱'; ja = '転送先メール'; en = 'Forward to email' }
            'ForwardtoUsername' = @{ zh = '转发用户名'; ja = '転送先ユーザー名'; en = 'Forward to username' }
            'WQCDate' = @{ zh = 'WQC日期'; ja = 'WQC日付'; en = 'WQC date' }
            'Id' = @{ zh = '记录ID'; ja = 'レコードID'; en = 'Record ID' }
        }
    }
}

# ============================================================
# Detect table type from CSV first row
# ============================================================
function Detect-TableType {
    param([array]$Data)
    if ($Data.Count -eq 0) { return 'AuditGeneralDCR_CL' }
    $firstRow = $Data[0]
    foreach ($schema in $TableSchemas.Keys) {
        $fields = $TableSchemas[$schema]
        $userField = $fields.UserField
        if ($firstRow.PSObject.Properties.Name -contains $userField) {
            return $schema
        }
    }
    return 'AuditGeneralDCR_CL'
}

Write-Host "Loading CSV data..." -ForegroundColor Cyan
$data = Import-Csv -Path $csvPath -Encoding UTF8
$totalEvents = $data.Count
Write-Host "Loaded $totalEvents records" -ForegroundColor Green

# Detect table type and get schema
$tableType = Detect-TableType -Data $data
$schema = $TableSchemas[$tableType]
Write-Host "Detected table: $tableType ($($schema.DisplayName))" -ForegroundColor Cyan

Write-Host "Computing statistics..." -ForegroundColor Cyan

# Helper: get field value from row using schema
function Get-FieldValue {
    param([object]$Row, [string]$FieldName)
    if (-not $FieldName) { return '' }
    $val = $Row.$FieldName
    if ($val -eq $null) { return '' }
    return $val.ToString()
}

# Helper: get user from row using schema
function Get-User {
    param([object]$Row)
    $u = Get-FieldValue -Row $Row -FieldName $schema.UserField
    if (-not $u -and $schema.UserFallback) { $u = Get-FieldValue -Row $Row -FieldName $schema.UserFallback }
    if (-not $u) { $u = 'Unknown' }
    return $u
}

# Unique users
$allUsers = @()
foreach ($row in $data) {
    $allUsers += Get-User -Row $row
}
$uniqueUsers = ($allUsers | Select-Object -Unique).Count

# Unique operations
$allOps = @($data | ForEach-Object { Get-FieldValue -Row $_ -FieldName $schema.OpField })
$uniqueOps = ($allOps | Select-Object -Unique).Count

# Workload distribution
$workloadMap = @{}
$hasWorkload = ($schema.WlField -ne '')
if ($hasWorkload) {
    foreach ($row in $data) {
        $wl = Get-FieldValue -Row $row -FieldName $schema.WlField
        if (-not $wl) { $wl = 'Unknown' }
        $workloadMap[$wl] = ($workloadMap[$wl] + 1)
    }
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
$hasIP = ($schema.IpField -ne '')
if ($hasIP) {
    foreach ($row in $data) {
        $ip = Get-FieldValue -Row $row -FieldName $schema.IpField
        if (-not $ip) { $ip = 'Unknown' }
        $ipMap[$ip] = ($ipMap[$ip] + 1)
    }
}
$topIPs = $ipMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10

# Success/Failure - using schema-based field mapping
$successCount = 0
$failCount = 0
$unknownCount = 0
$successField = $schema.SuccessField
$successValue = $schema.SuccessValue
$failValue = $schema.FailValue

if ($successField) {
    # Table has success/failure field
    foreach ($row in $data) {
        $s = Get-FieldValue -Row $row -FieldName $successField
        if ($successValue -and $s -eq $successValue) {
            $successCount++
        }
        elseif ($failValue -and $s -eq $failValue) {
            $failCount++
        }
        else {
            $unknownCount++
        }
    }
}
else {
    # Table has no success/failure concept - all unknown
    $unknownCount = $totalEvents
}

# Flag: does this table have success/failure status?
$hasStatus = ($successField -ne '')

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
$offHoursEvents = @()
foreach ($row in $data) {
    $tg = $row.TimeGenerated
    if ($tg -and $tg -ne '') {
        try {
            $dt = [DateTime]::Parse($tg)
            $localDt = $dt.ToLocalTime()
            if ($localDt.Hour -ge 0 -and $localDt.Hour -lt 7) {
                $offHoursEvents += $row
            }
        } catch {}
    }
}

# Failed operations details - using schema-based field mapping
$failedEvents = @()
if ($successField -and $failValue) {
    $failedEvents = @($data | Where-Object { (Get-FieldValue -Row $_ -FieldName $successField) -eq $failValue })
}
$failedByOp = @{}
foreach ($row in $failedEvents) {
    $op = Get-FieldValue -Row $row -FieldName $schema.OpField
    if (-not $op) { $op = 'Unknown' }
    $failedByOp[$op] = ($failedByOp[$op] + 1)
}
$failedByOp = $failedByOp.GetEnumerator() | Sort-Object Value -Descending

# High-privilege operations - using schema-based field mapping
$highPrivOps = @('ExportReport', 'Search', 'EditDataset', 'Delete', 'DeleteDataset', 'DeleteReport', 'DeleteWorkspace', 'AdminAction')
$highPrivEvents = @($data | Where-Object { $highPrivOps -contains (Get-FieldValue -Row $_ -FieldName $schema.OpField) })
$highPrivByUser = @{}
foreach ($row in $highPrivEvents) {
    $u = Get-User -Row $row
    $op = Get-FieldValue -Row $row -FieldName $schema.OpField
    $key = "$u | $op"
    $highPrivByUser[$key] = ($highPrivByUser[$key] + 1)
}
$highPrivByUser = $highPrivByUser.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 20

# Sensitive data events - using schema-based field mapping
$sensitiveOps = @('SensitivityLabeledFileOpened', 'SensitivityLabeledFileRenamed', 'IrmContent', 'AppliedSensitivityLabel', 'ChangedSensitivityLabel')
$sensitiveEvents = @($data | Where-Object { $sensitiveOps -contains (Get-FieldValue -Row $_ -FieldName $schema.OpField) })
$sensitiveByOp = @{}
foreach ($row in $sensitiveEvents) {
    $op = Get-FieldValue -Row $row -FieldName $schema.OpField
    $sensitiveByOp[$op] = ($sensitiveByOp[$op] + 1)
}
$sensitiveByOp = $sensitiveByOp.GetEnumerator() | Sort-Object Value -Descending

# Service account activity (GUIDs starting with 0000-...) - using schema-based field mapping
$serviceAcctEvents = @($data | Where-Object { (Get-FieldValue -Row $_ -FieldName $schema.UserFallback) -match '^00000009-' })
$serviceAcctByOp = @{}
foreach ($row in $serviceAcctEvents) {
    $op = Get-FieldValue -Row $row -FieldName $schema.OpField
    if (-not $op) { $op = 'Unknown' }
    $serviceAcctByOp[$op] = ($serviceAcctByOp[$op] + 1)
}
$serviceAcctByOp = $serviceAcctByOp.GetEnumerator() | Sort-Object Value -Descending

# IP velocity - single IP with multiple users - using schema-based field mapping
$ipUsers = @{}
foreach ($row in $data) {
    $ip = Get-FieldValue -Row $row -FieldName $schema.IpField
    if (-not $ip) { continue }
    $u = Get-User -Row $row
    if (-not $ipUsers.ContainsKey($ip)) { $ipUsers[$ip] = @{} }
    $ipUsers[$ip][$u] = 1
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

# Suspicious IPs (non-RFC1918 accessing multiple workloads) - using schema-based field mapping
$ipWorkloads = @{}
foreach ($row in $data) {
    $ip = Get-FieldValue -Row $row -FieldName $schema.IpField
    if (-not $ip) { continue }
    # Skip RFC1918
    if ($ip -match '^(10\.|172\.(1[6-9]|2[0-9]|3[01])\.|192\.168\.)' -or $ip -eq 'Unknown' -or $ip -eq '0.0.0.0') { continue }
    $wl = Get-FieldValue -Row $row -FieldName $schema.WlField
    if (-not $wl) { $wl = 'Unknown' }
    if (-not $ipWorkloads.ContainsKey($ip)) { $ipWorkloads[$ip] = @{} }
    $ipWorkloads[$ip][$wl] = 1
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

# Off-hours by user - using schema-based field mapping
$offHoursByUser = @{}
foreach ($row in $offHoursEvents) {
    $u = Get-User -Row $row
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

# Build glossary (only ops present in data)
$glossaryOps = @{}
$opNames = $opMap.Keys | Sort-Object
foreach ($op in $opNames) {
    $expZh = $op  # default fallback
    $expJa = $op
    # Known glossary entries
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
    $items = $map.GetEnumerator() | Sort-Object Value -Descending
    $json = "["
    $first = $true
    foreach ($item in $items) {
        if (-not $first) { $json += ',' }
        $name = (EscapeHtml $item.Name) -replace '"', '\"'
        $json += '{' + '"name":"' + $name + '","value":' + $item.Value + '}'
        $first = $false
    }
    $json += "]"
    return $json
}

function ToSortedJsonArray {
    param([array]$items, [string]$nameKey = 'Name', [string]$valueKey = 'Value')
    $json = "["
    $first = $true
    foreach ($item in $items) {
        if (-not $first) { $json += ',' }
        $name = (EscapeHtml $item.$nameKey) -replace '"', '\"'
        $json += '{' + '"name":"' + $name + '","value":' + $item.$valueKey + '}'
        $first = $false
    }
    $json += "]"
    return $json
}

function ToKeyValueJsonArray {
    param([array]$items, [string]$keyName = 'Name', [string]$valName = 'Value')
    $json = "["
    $first = $true
    foreach ($item in $items) {
        if (-not $first) { $json += ',' }
        $k = (EscapeHtml $item.$keyName) -replace '"', '\"'
        $v = $item.$valName
        if ($v -is [int] -or $v -is [double]) {
            $json += '{' + '"key":"' + $k + '","value":' + $v + '}'
        } else {
            $vs = (EscapeHtml $v) -replace '"', '\"'
            $json += '{' + '"key":"' + $k + '","value":"' + $vs + '"}'
        }
        $first = $false
    }
    $json += "]"
    return $json
}

$topUsersJson = ToSortedJsonArray $topUsers
$topOpsJson = ToSortedJsonArray $topOps
$topIPsJson = ToSortedJsonArray $topIPs
$timelineJson = ToKeyValueJsonArray $timelineSorted 'Name' 'Value'
$workloadJson = ToJsonArray $workloadMap

# Failed events JSON
$failedOpsJson = ToSortedJsonArray -items $failedByOp -nameKey 'Name' -valueKey 'Value'

# High-priv events JSON
$highPrivJson = ToSortedJsonArray -items $highPrivByUser -nameKey 'Name' -valueKey 'Value'

# Sensitive events JSON
$sensitiveJson = ToSortedJsonArray -items $sensitiveByOp -nameKey 'Name' -valueKey 'Value'

# Service accounts JSON
$serviceAcctJson = ToSortedJsonArray -items $serviceAcctByOp -nameKey 'Name' -valueKey 'Value'

# Off-hours JSON
$offHoursJson = ToSortedJsonArray -items $offHoursByUser
$suspiciousIPsJson = ToKeyValueJsonArray -items $suspiciousIPs -keyName 'IP' -valName 'Workloads'
$ipVelocityJson = ToKeyValueJsonArray -items $ipVelocity -keyName 'IP' -valName 'UserCount'

# Glossary JSON
$glossaryJson = "{"
$gfirst = $true
foreach ($op in $glossaryOps.Keys | Sort-Object) {
    if (-not $gfirst) { $glossaryJson += ',' }
    $opName = (EscapeHtml $op) -replace '"', '\"'
    $zhVal = $glossaryOps[$op].zh -replace '"', '\"'
    $jaVal = $glossaryOps[$op].ja -replace '"', '\"'
    $glossaryJson += '"' + $opName + '":{"zh":"' + $zhVal + '","ja":"' + $jaVal + '","count":' + $glossaryOps[$op].count + '}'
    $gfirst = $false
}
$glossaryJson += "}"

# Workload glossary JSON
$wlGlossaryJson = "{"
$wfirst = $true
foreach ($wl in $wlGlossary.Keys | Sort-Object) {
    if ($workloadMap.ContainsKey($wl)) {
        if (-not $wfirst) { $wlGlossaryJson += ',' }
        $wlName = (EscapeHtml $wl) -replace '"', '\"'
        $zhVal = $wlGlossary[$wl].zh -replace '"', '\"'
        $jaVal = $wlGlossary[$wl].ja -replace '"', '\"'
        $wlGlossaryJson += '"' + $wlName + '":{"zh":"' + $zhVal + '","ja":"' + $jaVal + '"}'
        $wfirst = $false
    }
}
$wlGlossaryJson += "}"

# Build data table - first 500 rows - using schema-based field mapping
$tableRows = ''
$previewRows = [Math]::Min(500, $data.Count)
for ($i = 0; $i -lt $previewRows; $i++) {
    $row = $data[$i]
    $tg = Get-FieldValue -Row $row -FieldName $schema.TimeField
    $user = Get-User -Row $row
    $op = Get-FieldValue -Row $row -FieldName $schema.OpField
    $wl = Get-FieldValue -Row $row -FieldName $schema.WlField
    $ip = Get-FieldValue -Row $row -FieldName $schema.IpField
    
    # Build row cells based on available fields
    $cells = "<td>$i</td><td>$(EscapeHtml $tg)</td><td class='op-cell' data-op='$(EscapeHtml $op)'>$(EscapeHtml $op)</td><td>$(EscapeHtml $user)</td>"
    if ($hasWorkload) { $cells += "<td>$(EscapeHtml $wl)</td>" }
    if ($hasIP) { $cells += "<td>$(EscapeHtml $ip)</td>" }
    if ($hasStatus) {
        $success = Get-FieldValue -Row $row -FieldName $successField
        $cells += "<td class='status-$(EscapeHtml $success)'>$(EscapeHtml $success)</td>"
    }
    $tableRows += "<tr>$cells</tr>`n"
}

# ============================================================
# AI Analysis - Generate insights from log data
# ============================================================
Write-Host "Generating AI analysis..." -ForegroundColor Cyan

$aiInsights = @()

# Insight 1: Activity pattern analysis
if ($timelineSorted.Count -gt 0) {
    $peakHour = $timelineSorted | Sort-Object Value -Descending | Select-Object -First 1
    $aiInsights += @{
        type = 'activity'
        severity = 'info'
        title_zh = '活动高峰时段'
        title_ja = 'アクティビティピーク時間'
        content_zh = "日志活动高峰时段为 $($peakHour.Name)，共 $($peakHour.Value) 条记录。建议关注该时段的异常活动。"
        content_ja = "ログアクティビティのピーク時間は $($peakHour.Name) で、$($peakHour.Value) 件の記録があります。この時間帯の異常アクティビティに注意してください。"
    }
}

# Insight 2: Top user analysis
if ($topUsers.Count -gt 0) {
    $topUser = $topUsers | Select-Object -First 1
    $aiInsights += @{
        type = 'user'
        severity = 'info'
        title_zh = '最活跃用户'
        title_ja = '最もアクティブなユーザー'
        content_zh = "用户 $($topUser.Name) 是最活跃用户，共执行 $($topUser.Value) 次操作。建议定期审查高活跃用户权限。"
        content_ja = "ユーザー $($topUser.Name) が最もアクティブで、$($topUser.Value) 回の操作を実行しました。高アクティブユーザーの権限を定期的にレビューしてください。"
    }
}

# Insight 3: Failure rate analysis
if ($totalEvents -gt 0 -and $successField) {
    $failRate = [math]::Round(($failCount / $totalEvents) * 100, 2)
    if ($failRate -gt 10) {
        $aiInsights += @{
            type = 'failure'
            severity = 'high'
            title_zh = '高失败率警告'
            title_ja = '高失敗率警告'
            content_zh = "操作失败率为 $failRate% ($failCount/$totalEvents)，超过10%阈值。建议立即调查失败原因。"
            content_ja = "操作失敗率は $failRate% ($failCount/$totalEvents) で、10%の閾値を超えています。失敗理由を直ちに調査してください。"
        }
    }
}

# Insight 4: Workload concentration
if ($workloadMap.Count -gt 0) {
    $topWl = $workloadMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
    $wlPct = [math]::Round(($topWl.Value / $totalEvents) * 100, 1)
    $aiInsights += @{
        type = 'workload'
        severity = 'info'
        title_zh = '工作负载集中度'
        title_ja = 'ワークロード集中度'
        content_zh = "$($topWl.Name) 是主要工作负载，占总活动的 $wlPct%。建议重点关注该服务的安全配置。"
        content_ja = "$($topWl.Name) が主要なワークロードで、全アクティビティの $wlPct% を占めています。このサービスのセキュリティ設定に重点的に注意してください。"
    }
}

# Insight 5: Off-hours activity
if ($offHoursEvents.Count -gt 0) {
    $offHoursPct = [math]::Round(($offHoursEvents.Count / $totalEvents) * 100, 1)
    $aiInsights += @{
        type = 'offhours'
        severity = 'medium'
        title_zh = '非工作时间活动'
        title_ja = '時間外アクティビティ'
        content_zh = "检测到 $($offHoursEvents.Count) 条非工作时间活动记录 (00:00-07:00)，占总活动的 $offHoursPct%。建议审查这些操作是否合规。"
        content_ja = "時間外アクティビティ (00:00-07:00) が $($offHoursEvents.Count) 件検出され、全アクティビティの $offHoursPct% を占めています。これらの操作がコンプライアンスに準拠しているかレビューしてください。"
    }
}

# Insight 6: Suspicious IP detection
if ($suspiciousIPs.Count -gt 0) {
    $aiInsights += @{
        type = 'security'
        severity = 'high'
        title_zh = '可疑IP检测'
        title_ja = '不審なIP検出'
        content_zh = "发现 $($suspiciousIPs.Count) 个可疑IP地址，这些IP访问了多个工作负载。建议立即调查这些IP的来源和活动。"
        content_ja = "$($suspiciousIPs.Count) 件の不審なIPアドレスが発見されました。これらのIPは複数のワークロードにアクセスしています。IPのソースとアクティビティを直ちに調査してください。"
    }
}

# Convert AI insights to JSON
$aiInsightsJson = "["
$aiFirst = $true
foreach ($insight in $aiInsights) {
    if (-not $aiFirst) { $aiInsightsJson += ',' }
    $aiInsightsJson += '{'
    $aiInsightsJson += '"type":"' + $insight.type + '",'
    $aiInsightsJson += '"severity":"' + $insight.severity + '",'
    $aiInsightsJson += '"title_zh":"' + ($insight.title_zh -replace '"', '\"') + '",'
    $aiInsightsJson += '"title_ja":"' + ($insight.title_ja -replace '"', '\"') + '",'
    $aiInsightsJson += '"content_zh":"' + ($insight.content_zh -replace '"', '\"') + '",'
    $aiInsightsJson += '"content_ja":"' + ($insight.content_ja -replace '"', '\"') + '"'
    $aiInsightsJson += '}'
    $aiFirst = $false
}
$aiInsightsJson += "]"

# Build field dictionary JSON
$fieldsJson = "{"
$fFirst = $true
foreach ($fieldName in $schema.Fields.Keys | Sort-Object) {
    if (-not $fFirst) { $fieldsJson += ',' }
    $f = $schema.Fields[$fieldName]
    $fieldsJson += '"' + $fieldName + '":{'
    $fieldsJson += '"zh":"' + ($f.zh -replace '"', '\"') + '",'
    $fieldsJson += '"ja":"' + ($f.ja -replace '"', '\"') + '",'
    $fieldsJson += '"en":"' + ($f.en -replace '"', '\"') + '"'
    $fieldsJson += '}'
    $fFirst = $false
}
$fieldsJson += "}"

# Build Azure Log Analytics query URL
$workspaceId = '703a5771-97fc-4bf3-a585-f607d18c4479'
$azurePortalUrl = "https://portal.azure.cn/#@$($tenantId)/resource/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.OperationalInsights/workspaces/$workspaceId/logs"
$queryUrl = "$azurePortalUrl?query=$tableType"

# Determine risk count
$riskCount = 0
if ($failCount -gt 0) { $riskCount++ }
if ($suspiciousIPs.Count -gt 0) { $riskCount++ }
if ($offHoursEvents.Count -gt 0) { $riskCount++ }
if ($highPrivEvents.Count -gt 0) { $riskCount++ }
if ($sensitiveEvents.Count -gt 0) { $riskCount++ }
if ($ipVelocity.Count -gt 0) { $riskCount++ }

# Determine if we should show risk section
$showRiskSection = ($riskCount -gt 0).ToString().ToLower()

$html = @"
<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Azure Audit Report - $analysisDate</title>
<style>
:root {
  --bg-primary: #0d1117; --bg-secondary: #161b22; --bg-tertiary: #21262d;
  --border: #30363d; --text-primary: #e6edf3; --text-secondary: #8b949e;
  --accent: #58a6ff; --accent-green: #3fb950; --accent-red: #f85149;
  --accent-yellow: #d29922; --accent-purple: #bc8cff; --accent-orange: #f0883e;
  --accent-cyan: #39d2c0;
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
.bar-fill { height: 100%; border-radius: 4px; display: flex; align-items: center; padding-left: 8px; font-size: 11px; color: rgba(255,255,255,0.9); min-width: 30px; transition: width 0.3s ease; }
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
.seg-fill { height: 100%; display: flex; align-items: center; justify-content: center; font-size: 11px; color: rgba(255,255,255,0.9); }
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
.tooltip-box { position: fixed; background: var(--bg-tertiary); border: 1px solid var(--border); border-radius: 6px; padding: 10px 14px; font-size: 12px; z-index: 1000; pointer-events: none; max-width: 350px; box-shadow: 0 8px 24px rgba(0,0,0,0.4); display: none; }
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
  <h2>Azure Audit Report</h2>
  <div class="lang-toggle">
    <button class="active" onclick="switchLang('zh')">中文</button>
    <button onclick="switchLang('ja')">日本語</button>
  </div>
</div>

<div class="header">
  <h1>$($schema.DisplayName) Report</h1>
  <div class="subtitle" data-i18n="subtitle">$($schema.Description)</div>
  <div class="meta-tags">
    <span class="meta-tag">查询时间段: $analysisDate</span>
    <span class="meta-tag">Total Records: $totalEvents</span>
    <span class="meta-tag">Source Table: $tableType</span>
    <span class="meta-tag"><a href="$queryUrl" target="_blank" style="color:var(--accent);text-decoration:underline;">在 Azure 门户中查看日志</a></span>
  </div>
</div>

<!-- Table Info Section -->
<div class="section" id="table-info-section">
  <h2 data-i18n="tableInfo">日志表信息</h2>
  <table class="risk-table">
    <tbody>
      <tr><th style="width:150px;">日志表名称</th><td>$tableType</td></tr>
      <tr><th>说明 (CN)</th><td>$($schema.Description)</td></tr>
      <tr><th>説明 (JP)</th><td>$($schema.DescriptionJa)</td></tr>
    </tbody>
  </table>
</div>

<!-- Field Dictionary Section -->
<div class="section" id="field-dict-section">
  <h2 data-i18n="fieldDict">字段字典</h2>
  <p style="color:var(--text-secondary);font-size:13px;margin-bottom:12px;" data-i18n="fieldDictDesc">日志表中各字段的含义</p>
  <table class="risk-table">
    <thead><tr>
      <th>字段名 (EN)</th>
      <th>说明 (CN)</th>
      <th>説明 (JP)</th>
      <th>Description</th>
    </tr></thead>
    <tbody id="field-dict-body"></tbody>
  </table>
</div>

<!-- AI Analysis Section -->
<div class="section" id="ai-analysis-section">
  <h2 data-i18n="aiAnalysis">AI 智能分析</h2>
  <p style="color:var(--text-secondary);font-size:13px;margin-bottom:12px;" data-i18n="aiAnalysisDesc">基于日志数据的自动分析结论</p>
  <div id="ai-analysis-content"></div>
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
$(if($hasStatus){@"
  <div class="summary-card">
    <div class="label" data-i18n="success">成功</div>
    <div class="value green">$successCount</div>
  </div>
  <div class="summary-card">
    <div class="label" data-i18n="failed">失败</div>
    <div class="value red">$failCount</div>
  </div>
"@})
</div>

<div class="section">
  <h2 data-i18n="activityTimeline">活动时间线</h2>
  <div id="timeline-chart"></div>
</div>

<div class="section" id="workload-section" style="display: $(if($hasWorkload){'block'}else{'none'});">
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

<div class="section" id="ip-section" style="display: $(if($hasIP){'block'}else{'none'});">
  <h2 data-i18n="topIPs">客户端 IP 排行</h2>
  <div id="ips-chart" class="bar-chart"></div>
</div>

<div class="section" id="success-ratio-section" style="display: $(if($hasStatus){'block'}else{'none'});">
  <h2 data-i18n="successRatio">成功/失败比率</h2>
  <div id="success-ratio"></div>
</div>

<div class="section" id="risk-section" style="display: $(if($showRiskSection -eq 'true'){'block'}else{'none'});">
  <h2 data-i18n="riskAnalysis">风险分析</h2>
  <p style="color:var(--accent-red);margin-bottom:16px;font-size:14px;" data-i18n="riskIndicators">$riskCount 个风险指标已检出</p>
  <div id="risk-content"></div>
</div>

<div class="section">
  <h2 data-i18n="detailedTable">详细数据</h2>
  <p style="color:var(--text-secondary);font-size:13px;margin-bottom:12px;" data-i18n="tablePreview">预览前 500 行</p>
  <div class="table-wrapper">
    <div class="table-scroll">
      <table id="data-table">
        <thead>
          <tr>
            <th onclick="sortTable(0)">#</th>
            <th onclick="sortTable(1)" data-i18n="time">时间</th>
            <th onclick="sortTable(2)" data-i18n="operation">操作</th>
            <th onclick="sortTable(3)" data-i18n="user">用户</th>
$(if($hasWorkload){'            <th onclick="sortTable(4)" data-i18n="workload">工作负载</th>'})
$(if($hasIP){'            <th onclick="sortTable(5)" data-i18n="clientIP">IP</th>'})
$(if($hasStatus){'            <th onclick="sortTable(6)" data-i18n="status">状态</th>'})
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
    "success":"成功","failed":"失败","activityTimeline":"活动时间线","workloadDist":"工作负载分布",
    "topUsers":"活跃用户排行","topOps":"操作类型排行","topIPs":"客户端 IP 排行",
    "successRatio":"成功/失败比率","riskAnalysis":"风险分析","detailedTable":"详细数据",
    "showGlossary":"显示术语表","hideGlossary":"隐藏术语表","metric":"指标","value":"值",
    "severity":"严重程度","unknown":"未知","previous":"上一页","next":"下一页",
    "subtitle":"Office365 审计日志分析报告","tablePreview":"预览前 500 行",
    "time":"时间","operation":"操作","user":"用户","workload":"工作负载","clientIP":"IP",
    "status":"状态","riskIndicators":"$riskCount 个风险指标已检出",
    "failedOps":"失败操作","suspiciousIPs":"可疑 IP","offHours":"非工作时间活动",
    "highPrivOps":"高权限操作","sensitiveData":"敏感数据事件","ipVelocity":"IP 多用户关联",
    "serviceAccounts":"服务账户活动","low":"低风险","medium":"中风险","high":"高风险",
    "offHoursUsers":"非工作时活跃用户","failedOpSummary":"失败操作汇总","count":"次数",
    "tableInfo":"日志表信息","fieldDict":"字段字典","fieldDictDesc":"日志表中各字段的含义",
    "aiAnalysis":"AI 智能分析","aiAnalysisDesc":"基于日志数据的自动分析结论"
  },
  ja: {
    "totalEvents":"総イベント数","uniqueUsers":"ユニークユーザー","uniqueOps":"ユニーク操作","workloads":"ワークロード",
    "success":"成功","failed":"失敗","activityTimeline":"アクティビティタイムライン","workloadDist":"ワークロード分布",
    "topUsers":"アクティブユーザーランキング","topOps":"操作タイプランキング","topIPs":"クライアント IP ランキング",
    "successRatio":"成功/失敗比率","riskAnalysis":"リスク分析","detailedTable":"詳細データ",
    "showGlossary":"用語集を表示","hideGlossary":"用語集を非表示","metric":"指標","value":"値",
    "severity":"重要度","unknown":"不明","previous":"前へ","next":"次へ",
    "subtitle":"Office365 監査ログ分析レポート","tablePreview":"最初の 500 行をプレビュー",
    "time":"時間","operation":"操作","user":"ユーザー","workload":"ワークロード","clientIP":"IP",
    "status":"ステータス","riskIndicators":"$riskCount 件のリスク指標が検出されました",
    "failedOps":"失敗した操作","suspiciousIPs":"不審な IP","offHours":"時間外のアクティビティ",
    "highPrivOps":"高権限操作","sensitiveData":"機密データイベント","ipVelocity":"IP 複数ユーザー",
    "serviceAccounts":"サービスアカウント","low":"低リスク","medium":"中リスク","high":"高リスク",
    "offHoursUsers":"時間外アクティブユーザー","failedOpSummary":"失敗操作まとめ","count":"回数",
    "tableInfo":"ログテーブル情報","fieldDict":"フィールド辞書","fieldDictDesc":"ログテーブルの各フィールドの意味",
    "aiAnalysis":"AI 分析","aiAnalysisDesc":"ログデータに基づく自動分析結果"
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
const glossaryData = $glossaryJson;
const wlGlossaryData = $wlGlossaryJson;
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

function renderBarChart(containerId, data, maxItems) {
  const container = document.getElementById(containerId);
  if (!container || !data || data.length === 0) { container.innerHTML = '<p style="color:var(--text-secondary)">No data</p>'; return; }
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

function renderDonut(containerId, data) {
  const container = document.getElementById(containerId);
  if (!container || !data || data.length === 0) return;
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
  if (!container || !data || data.length === 0) return;
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
  if (total === 0) return;
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
  failedOps: $failedOpsJson,
  highPriv: $highPrivJson,
  sensitive: $sensitiveJson,
  serviceAcct: $serviceAcctJson,
  offHoursUsers: $offHoursJson,
  suspiciousIPs: $suspiciousIPsJson,
  ipVelocity: $ipVelocityJson,
  failCount: $failCount,
  highPrivCount: $($highPrivEvents.Count),
  sensitiveCount: $($sensitiveEvents.Count),
  serviceAcctCount: $($serviceAcctEvents.Count),
  offHoursCount: $($offHoursEvents.Count),
  suspiciousCount: $($suspiciousIPs.Count),
  ipVelocityCount: $($ipVelocity.Count)
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

// ===== Field Dictionary =====
const fieldsData = $fieldsJson;

function renderFieldDictionary() {
  const tbody = document.getElementById('field-dict-body');
  if (!tbody) return;
  let html = '';
  const keys = Object.keys(fieldsData).sort();
  keys.forEach(field => {
    const f = fieldsData[field];
    html += '<tr><td><code>' + field + '</code></td><td>' + f.zh + '</td><td>' + f.ja + '</td><td>' + f.en + '</td></tr>';
  });
  tbody.innerHTML = html;
}

// ===== AI Analysis =====
const aiData = $aiInsightsJson;

function renderAIAnalysis() {
  const container = document.getElementById('ai-analysis-content');
  if (!container || !aiData || aiData.length === 0) {
    if (container) container.innerHTML = '<p style="color:var(--text-secondary)">No AI analysis available.</p>';
    return;
  }
  let html = '';
  aiData.forEach(insight => {
    const severityClass = insight.severity === 'high' ? 'risk-high' : insight.severity === 'medium' ? 'risk-medium' : 'risk-low';
    const title = currentLang === 'zh' ? insight.title_zh : insight.title_ja;
    const content = currentLang === 'zh' ? insight.content_zh : insight.content_ja;
    html += '<div style="margin-bottom:16px;padding:16px;background:var(--bg-tertiary);border-radius:8px;border-left:4px solid var(--accent-' + (insight.severity === 'high' ? 'red' : insight.severity === 'medium' ? 'orange' : 'green') + ')">';
    html += '<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">';
    html += '<span class="risk-badge ' + severityClass + '">' + title + '</span>';
    html += '</div>';
    html += '<p style="color:var(--text-secondary);font-size:13px;line-height:1.6;">' + content + '</p>';
    html += '</div>';
  });
  container.innerHTML = html;
}

// ===== Init =====
document.addEventListener('DOMContentLoaded', function() {
  renderTimeline('timeline-chart', $timelineJson);
  renderDonut('donut-chart', $workloadJson);
  renderBarChart('users-chart', $topUsersJson, 15);
  renderBarChart('ops-chart', $topOpsJson, 15);
  renderBarChart('ips-chart', $topIPsJson, 10);
  $(if($hasStatus){
    "renderSuccessRatio('success-ratio', $successCount, $failCount, $unknownCount);"
  }else{
    "// No success/failure status for this table - hiding ratio chart`n  document.getElementById('success-ratio').parentElement.style.display = 'none';"
  })
  initPagination();
  buildRiskSection();
  renderFieldDictionary();
  renderAIAnalysis();
});
</script>
</body>
</html>
"@

# Use UTF8 with BOM to prevent garbled filenames on Windows
[System.IO.File]::WriteAllText($outputPath, $html, [System.Text.UTF8Encoding]::new($true))
Write-Host "Report saved to: $outputPath" -ForegroundColor Green
