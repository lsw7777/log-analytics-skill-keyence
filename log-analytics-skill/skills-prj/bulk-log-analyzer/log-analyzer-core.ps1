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
    [PSCustomObject]@{ Name = 'SigninLogs'; Description = 'Microsoft Entra Sign-in Logs' },
    [PSCustomObject]@{ Name = 'WQCLogDCR_CL'; Description = 'WQC Log' }
)

function Get-SupportedLogTables {
    return $SupportedLogTables
}

function Resolve-LogTableSelection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Selection
    )

    $trimmed = $Selection.Trim()
    $index = 0
    if ([int]::TryParse($trimmed, [ref]$index)) {
        if ($index -ge 1 -and $index -le $SupportedLogTables.Count) {
            return $SupportedLogTables[$index - 1].Name
        }
    }

    foreach ($table in $SupportedLogTables) {
        if ($table.Name -eq $trimmed) {
            return $table.Name
        }
    }

    throw "Invalid table selection: $Selection"
}

function Select-LogTableInteractive {
    Write-Host 'Select a log table to process:' -ForegroundColor Cyan
    for ($i = 0; $i -lt $SupportedLogTables.Count; $i++) {
        $table = $SupportedLogTables[$i]
        Write-Host ("  {0}. {1} ({2})" -f ($i + 1), $table.Name, $table.Description) -ForegroundColor Cyan
    }

    $selection = Read-Host 'Enter number'
    return Resolve-LogTableSelection -Selection $selection
}

function Get-DefaultLogTimeRange {
    param(
        [datetime]$Now = (Get-Date)
    )

    $endTime = $Now
    $startTime = $endTime.AddHours(-3)

    return [PSCustomObject]@{
        StartTime = $startTime
        EndTime = $endTime
        AnalysisDateStr = "$($startTime.ToString('yyyyMMddHHmm'))_$($endTime.ToString('yyyyMMddHHmm'))"
        AnalysisDateDisplay = "$($startTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    }
}

function Get-RelativeLogTimeRange {
    param(
        [datetime]$Now = (Get-Date),
        [int]$Days = 0
    )

    if ($Days -lt 0 -or $Days -gt 31) {
        throw 'Days must be between 0 and 31. Use 0 for the last 3 hours.'
    }

    $endTime = $Now
    $startTime = if ($Days -eq 0) { $endTime.AddHours(-3) } else { $endTime.AddDays(-$Days) }

    return [PSCustomObject]@{
        StartTime = $startTime
        EndTime = $endTime
        AnalysisDateStr = "$($startTime.ToString('yyyyMMddHHmm'))_$($endTime.ToString('yyyyMMddHHmm'))"
        AnalysisDateDisplay = "$($startTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    }
}

function Assert-LogTimeRangeWithinLimit {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$StartTime,

        [Parameter(Mandatory = $true)]
        [datetime]$EndTime,

        [int]$MaxDays = 31
    )

    if ($EndTime -le $StartTime) {
        throw 'End time must be greater than start time.'
    }

    if (($EndTime - $StartTime).TotalDays -gt $MaxDays) {
        throw "Time range cannot exceed $MaxDays days."
    }
}

function Get-LogTimeRangeFromDates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StartDate,

        [Parameter(Mandatory = $true)]
        [string]$EndDate
    )

    $startDay = [DateTime]::Parse($StartDate).Date
    $endDay = [DateTime]::Parse($EndDate).Date
    if ($endDay -lt $startDay) {
        throw "End date must be greater than or equal to start date."
    }

    Assert-LogTimeRangeWithinLimit -StartTime $startDay -EndTime $endDay.AddDays(1)

    $analysisDateStr = $startDay.ToString('yyyyMMdd')
    $analysisDateDisplay = $startDay.ToString('yyyy-MM-dd')
    if ($endDay -ne $startDay) {
        $analysisDateStr = "$($startDay.ToString('yyyyMMdd'))_$($endDay.ToString('yyyyMMdd'))"
        $analysisDateDisplay = "$($startDay.ToString('yyyy-MM-dd')) to $($endDay.ToString('yyyy-MM-dd'))"
    }

    return [PSCustomObject]@{
        StartTime = $startDay
        EndTime = $endDay.AddDays(1)
        AnalysisDateStr = $analysisDateStr
        AnalysisDateDisplay = $analysisDateDisplay
    }
}

function Select-LogTimeRangeInteractive {
    param(
        [datetime]$Now = (Get-Date)
    )

    Write-Host 'Select log time range:' -ForegroundColor Cyan
    Write-Host '  0. Last 3 hours (test range)' -ForegroundColor Cyan
    Write-Host '  1. Last 1 day' -ForegroundColor Cyan
    Write-Host '  n. Last n days (max 31)' -ForegroundColor Cyan
    $selection = Read-Host 'Enter 0, 1, or n'

    if ([string]::IsNullOrWhiteSpace($selection)) {
        $selection = '0'
    }

    $days = 0
    if (-not [int]::TryParse($selection.Trim(), [ref]$days)) {
        throw "Invalid time range selection: $selection"
    }
    if ($days -lt 0 -or $days -gt 31) {
        throw 'Time range number must be between 0 and 31.'
    }

    return Get-RelativeLogTimeRange -Now $Now -Days $days
}

function ConvertTo-RelativeLogPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    $baseFull = [System.IO.Path]::GetFullPath($BasePath)
    $targetFull = [System.IO.Path]::GetFullPath($TargetPath)
    if (-not $baseFull.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $baseFull += [System.IO.Path]::DirectorySeparatorChar
    }

    $baseUri = [System.Uri]::new($baseFull)
    $targetUri = [System.Uri]::new($targetFull)
    $relative = [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString()) -replace '/', [System.IO.Path]::DirectorySeparatorChar

    if (-not $relative.StartsWith('.')) {
        $relative = ".\$relative"
    }

    return $relative
}

function Get-LogAnalyzerRepositoryRoot {
    $coreDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    return Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $coreDir))
}

function Get-TableAnalysisProfile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TableName
    )

    $common = [PSCustomObject]@{
        UserFields = @('UserUPN', 'UserId', 'UserPrincipalName', 'UPN', 'User', 'UserName', 'Mail', 'EmailAddress', 'DisplayName', 'Identity', 'Account', 'AccountName', 'SenderAddress', 'RecipientAddress')
        OperationFields = @('Operation', 'Activity', 'Action', 'EventName', 'OperationName', 'ActivityDisplayName', 'EventType', 'ActionType', 'Command', 'Cmdlet', 'Status', 'RecordType')
        WorkloadFields = @('Workload', 'Service', 'SourceSystem', 'RecordType', 'App', 'Application', 'ApplicationName', 'Product', 'Source', 'Category')
        ClientIpFields = @('ClientIP', 'ClientIp', 'ClientIPAddress', 'ClientIpAddress', 'IPAddress', 'SourceIP', 'SourceIp', 'SenderIP', 'SenderIp', 'RemoteIP', 'RemoteIp', 'IP', 'Client_IP_Address')
        SuccessFields = @('IsSuccess', 'ResultStatus', 'Status', 'Result', 'DeliveryStatus', 'EventStatus', 'ActionStatus', 'OperationStatus')
        DefaultUser = 'System'
        DefaultOperation = ($TableName -replace 'DCR_CL$', '')
        DefaultWorkload = ($TableName -replace 'DCR_CL$', '')
        DefaultClientIp = 'N/A'
        DefaultSuccess = 'unknown'
        GroupFields = @()
        UseCompositeOperationGroup = $false
    }

    switch ($TableName) {
        'AADManagedIdentitySignInLogs' {
            $common.UserFields = @('ServicePrincipalName', 'Identity', 'ManagedIdentityName', 'AppDisplayName', 'ResourceIdentity', 'ServicePrincipalId', 'AppId') + $common.UserFields
            $common.OperationFields = @('ResultType', 'ResultDescription', 'Status', 'ConditionalAccessStatus', 'OperationName', 'ActivityDisplayName') + $common.OperationFields
            $common.WorkloadFields = @('ResourceDisplayName', 'ResourceServicePrincipalId', 'ServicePrincipalName', 'AppDisplayName', 'Category') + $common.WorkloadFields
            $common.ClientIpFields = @('IPAddress', 'IpAddress', 'ClientIP', 'ClientIpAddress') + $common.ClientIpFields
            $common.SuccessFields = @('ResultType', 'Status', 'ResultDescription') + $common.SuccessFields
            $common.DefaultOperation = 'Managed Identity Sign-in'
            $common.DefaultWorkload = 'Microsoft Entra Managed Identity'
            $common.DefaultSuccess = 'unknown'
        }
        'AADServicePrincipalSignInLogs' {
            $common.UserFields = @('ServicePrincipalName', 'AppDisplayName', 'ServicePrincipalId', 'AppId', 'ResourceDisplayName') + $common.UserFields
            $common.OperationFields = @('ResultType', 'ResultDescription', 'Status', 'ConditionalAccessStatus', 'OperationName', 'ActivityDisplayName') + $common.OperationFields
            $common.WorkloadFields = @('ResourceDisplayName', 'ResourceServicePrincipalId', 'AppDisplayName', 'Category') + $common.WorkloadFields
            $common.ClientIpFields = @('IPAddress', 'IpAddress', 'ClientIP', 'ClientIpAddress') + $common.ClientIpFields
            $common.SuccessFields = @('ResultType', 'Status', 'ResultDescription') + $common.SuccessFields
            $common.DefaultOperation = 'Service Principal Sign-in'
            $common.DefaultWorkload = 'Microsoft Entra Service Principal'
            $common.DefaultSuccess = 'unknown'
        }
        'AssignedLicensesDCR_CL' {
            $common.UserFields = @('UserPrincipalName', 'UserUPN', 'UPN', 'Mail', 'EmailAddress', 'DisplayName', 'UserId', 'ObjectId', 'Id') + $common.UserFields
            $common.OperationFields = @('ProvisioningStatus', 'ServicePlanName', 'SkuPartNumber', 'LicenseName', 'AssignedLicenses', 'Licenses', 'AccountEnabled', 'UserType') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'ServicePlanName', 'SkuPartNumber', 'SourceSystem') + $common.WorkloadFields
            $common.DefaultOperation = 'Assigned License Record'
            $common.DefaultWorkload = 'Microsoft Entra ID Licensing'
            $common.SuccessFields = @('ProvisioningStatus') + $common.SuccessFields
            $common.DefaultSuccess = 'unknown'
        }
        'AuditGeneralDCR_CL' {
            $common.UserFields = @('UserUPN', 'UserId', 'Actor', 'ActorId', 'ActorUserPrincipalName') + $common.UserFields
            $common.OperationFields = @('Activity', 'Operation') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'RecordType') + $common.WorkloadFields
            $common.GroupFields = @('Activity', 'Operation', 'Workload')
            $common.UseCompositeOperationGroup = $true
        }
        'AuditLogs' {
            $common.UserFields = @('InitiatedByUserPrincipalName', 'ActorUserPrincipalName', 'UserPrincipalName', 'Identity', 'Actor', 'InitiatedBy') + $common.UserFields
            $common.OperationFields = @('OperationName', 'ActivityDisplayName', 'Activity', 'Operation', 'Category', 'Result') + $common.OperationFields
            $common.WorkloadFields = @('Category', 'LoggedByService', 'Service', 'ResourceProvider', 'AADOperationType') + $common.WorkloadFields
            $common.ClientIpFields = @('IPAddress', 'IpAddress', 'ClientIP', 'ClientIpAddress') + $common.ClientIpFields
            $common.SuccessFields = @('Result', 'ResultType', 'Status', 'ActivityStatus') + $common.SuccessFields
            $common.DefaultOperation = 'Audit Log Event'
            $common.DefaultWorkload = 'Microsoft Entra Audit'
            $common.DefaultSuccess = 'unknown'
        }
        'AzureADUsersDCR_CL' {
            $common.UserFields = @('userPrincipalName', 'mail', 'displayName', 'UserPrincipalName', 'UserUPN', 'UPN', 'Mail', 'EmailAddress', 'DisplayName', 'Id', 'ObjectId') + $common.UserFields
            $common.OperationFields = @('department', 'jobTitle', 'companyName', 'Department', 'JobTitle', 'CompanyName') + $common.OperationFields
            $common.WorkloadFields = @('department', 'companyName', 'jobTitle', 'Workload', 'SourceSystem', 'UserType', 'Department') + $common.WorkloadFields
            $common.DefaultOperation = 'Azure AD User Record'
            $common.DefaultWorkload = 'AzureActiveDirectory'
            $common.DefaultSuccess = 'true'
        }
        'MailboxStatisticsDCR_CL' {
            $common.UserFields = @('UserPrincipalName', 'MailboxOwnerUPN', 'PrimarySmtpAddress', 'DisplayName', 'Identity', 'Mail', 'EmailAddress') + $common.UserFields
            $common.OperationFields = @('RecipientTypeDetails', 'MailboxType', 'IsArchiveMailbox', 'StorageLimitStatus', 'LastLogonTime') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'RecipientTypeDetails', 'MailboxType', 'SourceSystem') + $common.WorkloadFields
            $common.DefaultOperation = 'Mailbox Statistics Record'
            $common.DefaultWorkload = 'Exchange'
            $common.DefaultSuccess = 'true'
        }
        'MessageTraceDataDCR_CL' {
            $common.UserFields = @('SenderAddress', 'RecipientAddress', 'Sender', 'Recipient', 'FromIP', 'MessageId', 'Subject') + $common.UserFields
            $common.OperationFields = @('Status', 'EventType', 'Directionality', 'MessageTraceEvent', 'MessageTraceType', 'Subject') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'Service', 'Directionality', 'SourceSystem') + $common.WorkloadFields
            $common.ClientIpFields = @('FromIP', 'ToIP', 'ClientIP', 'SenderIP', 'RecipientIP') + $common.ClientIpFields
            $common.DefaultOperation = 'Message Trace Event'
            $common.DefaultWorkload = 'Exchange'
        }
        'SigninLogs' {
            $common.UserFields = @('UserPrincipalName', 'UserDisplayName', 'Identity', 'UserId', 'User') + $common.UserFields
            $common.OperationFields = @('AppDisplayName', 'ResultType', 'ResultDescription', 'Status', 'ConditionalAccessStatus') + $common.OperationFields
            $common.WorkloadFields = @('AppDisplayName', 'ResourceDisplayName', 'ClientAppUsed', 'Category') + $common.WorkloadFields
            $common.ClientIpFields = @('IPAddress', 'IpAddress', 'ClientIP', 'ClientIpAddress') + $common.ClientIpFields
            $common.SuccessFields = @('ResultType', 'Status', 'ResultDescription') + $common.SuccessFields
            $common.DefaultOperation = 'User Sign-in'
            $common.DefaultWorkload = 'Microsoft Entra Sign-in'
            $common.DefaultSuccess = 'unknown'
        }
        'SharePointAuditDCR_CL' {
            $common.UserFields = @('UserUPN', 'UserId', 'UserPrincipalName', 'Actor', 'ActorId') + $common.UserFields
            $common.OperationFields = @('Activity', 'Operation') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'SiteUrl', 'SourceSystem') + $common.WorkloadFields
            $common.GroupFields = @('Activity', 'Operation', 'Workload')
            $common.UseCompositeOperationGroup = $true
        }
        'WQCLogDCR_CL' {
            $common.UserFields = @('CurrentMail', 'CurrentUsername', 'ForwardtoMail', 'ForwardtoUsername', 'UserPrincipalName', 'UserUPN', 'UserId', 'User', 'Requester', 'Submitter', 'Owner') + $common.UserFields
            $common.OperationFields = @('OperationType', 'InboxRuleName', 'Result', 'Action', 'Operation', 'Activity', 'QueryType', 'Command', 'Status') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'Service', 'Category', 'SourceSystem') + $common.WorkloadFields
            $common.DefaultOperation = 'WQC Log Event'
            $common.DefaultWorkload = 'WQC'
            $common.DefaultSuccess = 'true'
        }
    }

    return $common
}

function Get-FieldValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Row,

        [Parameter(Mandatory = $true)]
        [string[]]$Names,

        [string]$Default = 'Unknown'
    )

    foreach ($name in $Names) {
        if ($Row.PSObject.Properties.Name -contains $name) {
            $value = [string]$Row.$name
            if ($value -and $value.Trim() -ne '') {
                return $value.Trim()
            }
        }
    }

    return $Default
}

function Get-UserValue {
    param([object]$Row, [string]$TableName)
    $profile = Get-TableAnalysisProfile -TableName $TableName
    return Get-FieldValue -Row $Row -Names $profile.UserFields -Default $profile.DefaultUser
}

function Get-OperationValue {
    param([object]$Row, [string]$TableName)
    $profile = Get-TableAnalysisProfile -TableName $TableName
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
    return Get-FieldValue -Row $Row -Names $profile.OperationFields -Default $profile.DefaultOperation
}

function Get-WorkloadValue {
    param([object]$Row, [string]$TableName)
    $profile = Get-TableAnalysisProfile -TableName $TableName
    return Get-FieldValue -Row $Row -Names $profile.WorkloadFields -Default $profile.DefaultWorkload
}

function Get-ClientIpValue {
    param([object]$Row, [string]$TableName)
    $profile = Get-TableAnalysisProfile -TableName $TableName
    return Get-FieldValue -Row $Row -Names $profile.ClientIpFields -Default $profile.DefaultClientIp
}

function Get-SuccessValue {
    param([object]$Row, [string]$TableName)

    $profile = Get-TableAnalysisProfile -TableName $TableName
    $value = (Get-FieldValue -Row $Row -Names $profile.SuccessFields -Default $profile.DefaultSuccess).ToLowerInvariant()
    if ($TableName -eq 'AssignedLicensesDCR_CL') {
        if ([string]::IsNullOrWhiteSpace($value) -or $value -eq 'unknown') { return 'unknown' }
        if ($value -eq 'success') { return 'true' }
        return 'false'
    }
    if ($value -match '^(true|success|succeeded|delivered|expanded|completed|complete|ok|pass|passed|0)$') { return 'true' }
    if ($value -match '^(false|fail|failed|failure|undelivered|blocked|rejected|denied|error|timeout|quarantined|1)$') { return 'false' }
    if ($value -match '^\d+$' -and [int]$value -ne 0) { return 'false' }
    if ($profile.DefaultSuccess -eq 'true') { return 'true' }
    return 'unknown'
}

function Get-LogCacheTimeKey {
    param([datetime]$Time)
    return $Time.ToUniversalTime().ToString('o')
}

function Test-LogCacheMetadataMatches {
    param(
        [object]$Meta,
        [string]$TableName,
        [datetime]$StartTime,
        [datetime]$EndTime
    )

    if ($null -eq $Meta) { return $false }
    if ($Meta.TableName -ne $TableName) { return $false }
    if (-not $Meta.StartTimeUtc -or -not $Meta.EndTimeUtc) { return $false }
    if ($Meta.StartTimeUtc -ne (Get-LogCacheTimeKey -Time $StartTime)) { return $false }
    if ($Meta.EndTimeUtc -ne (Get-LogCacheTimeKey -Time $EndTime)) { return $false }
    return $true
}

function Get-LogCsvRecordCount {
    param([string]$CsvPath)

    if (-not (Test-Path $CsvPath)) { return 0 }
    if ((Get-Item -Path $CsvPath).Length -eq 0) { return 0 }

    try {
        return @(Import-Csv -Path $CsvPath -Encoding UTF8).Count
    } catch {
        return 0
    }
}

function Test-LogCachePayloadValid {
    param(
        [string]$CacheCsv,
        [object]$RecordCount
    )

    if (-not (Test-Path $CacheCsv)) { return $false }
    if ((Get-Item -Path $CacheCsv).Length -eq 0) { return $false }
    if ($null -eq $RecordCount) { return $false }
    if ([int]$RecordCount -le 0) { return $false }
    return $true
}

function Clear-LogCache {
    param([string]$CacheDir)

    if (-not (Test-Path $CacheDir)) {
        New-Item -ItemType Directory -Path $CacheDir -Force | Out-Null
        return 0
    }

    $items = @(Get-ChildItem -Path $CacheDir -Force -ErrorAction SilentlyContinue)
    foreach ($item in $items) {
        Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }

    return $items.Count
}

function Get-TrustedLocationPaths {
    $repoRoot = Get-LogAnalyzerRepositoryRoot
    return @(
        (Join-Path $repoRoot 'TrustedLocation_KJ.txt'),
        (Join-Path $repoRoot 'TrustedLocation_IDC_Ali.txt')
    )
}

function Get-TrustedIpRules {
    param([string[]]$Paths = (Get-TrustedLocationPaths))

    $rules = [System.Collections.Generic.List[object]]::new()
    foreach ($path in $Paths) {
        if (-not (Test-Path $path)) { continue }
        foreach ($line in Get-Content -Path $path -Encoding UTF8) {
            $value = ($line -replace '#.*$', '').Trim()
            if (-not $value) { continue }
            $parts = $value.Split('/')
            $ip = $parts[0].Trim()
            $prefix = if ($parts.Count -gt 1) { [int]$parts[1] } else { 32 }
            if ($prefix -lt 0 -or $prefix -gt 32) { continue }
            $parsed = $null
            if (-not [System.Net.IPAddress]::TryParse($ip, [ref]$parsed)) { continue }
            $bytes = $parsed.GetAddressBytes()
            if ($bytes.Count -ne 4) { continue }
            [array]::Reverse($bytes)
            $addr = [BitConverter]::ToUInt32($bytes, 0)
            $mask = if ($prefix -eq 0) { [uint32]0 } else { [uint32]([uint32]::MaxValue -shl (32 - $prefix)) }
            $rules.Add([PSCustomObject]@{
                Source = Split-Path -Leaf $path
                Cidr = $value
                Address = $addr
                Mask = $mask
            }) | Out-Null
        }
    }

    return $rules.ToArray()
}

function Test-IpInTrustedRules {
    param(
        [string]$IP,
        [object[]]$Rules
    )

    $normalized = Get-NormalizedIpValue -IP $IP
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $false }
    $parsed = $null
    if (-not [System.Net.IPAddress]::TryParse($normalized, [ref]$parsed)) { return $false }
    $bytes = $parsed.GetAddressBytes()
    if ($bytes.Count -ne 4) { return $false }
    [array]::Reverse($bytes)
    $addr = [BitConverter]::ToUInt32($bytes, 0)
    foreach ($rule in @($Rules)) {
        if (($addr -band $rule.Mask) -eq ($rule.Address -band $rule.Mask)) {
            return $true
        }
    }
    return $false
}

function Get-NormalizedIpValue {
    param([string]$IP)

    if ([string]::IsNullOrWhiteSpace($IP)) { return '' }
    $normalized = $IP.Trim()
    if ($normalized -in @('Unknown', 'N/A', '-', '0.0.0.0', '::', '::1', '127.0.0.1', '255.255.255.255')) { return '' }

    $parsed = $null
    if ([System.Net.IPAddress]::TryParse($normalized, [ref]$parsed)) {
        if ($parsed.GetAddressBytes().Count -eq 4) { return $parsed.ToString() }
        return ''
    }

    $ipv4Match = [regex]::Match($normalized, '(?<!\d)(\d{1,3}(?:\.\d{1,3}){3})(?!\d)')
    if ($ipv4Match.Success) {
        $candidate = $ipv4Match.Groups[1].Value
        if ([System.Net.IPAddress]::TryParse($candidate, [ref]$parsed) -and $parsed.GetAddressBytes().Count -eq 4) {
            return $parsed.ToString()
        }
    }

    return ''
}

function Test-PrivateOrInvalidIp {
    param([string]$IP)

    $normalized = Get-NormalizedIpValue -IP $IP
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $true }
    if ($normalized -in @('0.0.0.0', '255.255.255.255')) { return $true }
    if ($normalized -match '^(127\.|169\.254\.)') { return $true }
    if ($normalized -match '^(10\.|172\.(1[6-9]|2[0-9]|3[01])\.|192\.168\.)') { return $true }
    return $false
}

function Test-DeleteOrDisableOperation {
    param(
        [string]$Operation,
        [string]$TableName = ''
    )

    if ([string]::IsNullOrWhiteSpace($Operation)) { return $false }
    if ($TableName -eq 'AzureADUsersDCR_CL') { return $false }
    $normalized = $Operation.ToLowerInvariant()
    return ($normalized -match '(^|[^a-z])(delete|deleted|remove|removed|disable|disabled|deactivate|deactivated)([^a-z]|$)')
}

function Test-LogOffHours {
    param(
        [string]$TimeGenerated,
        [int]$StartHour = 21,
        [int]$EndHour = 8
    )

    if ([string]::IsNullOrWhiteSpace($TimeGenerated)) { return $false }
    try {
        $localDt = ([DateTime]::Parse($TimeGenerated)).ToLocalTime()
        return ($localDt.Hour -ge $StartHour -or $localDt.Hour -lt $EndHour)
    } catch {
        return $false
    }
}

function ConvertTo-KqlStringLiteral {
    param([string]$Value)
    return '"' + (($Value -replace '\\', '\\') -replace '"', '\"') + '"'
}

function Get-TrustedIpKqlDynamicLiteral {
    $rules = @(Get-TrustedIpRules)
    $ips = @()
    foreach ($rule in $rules) {
        if ($rule.Cidr -match '^([^/]+)/32$') {
            $ips += $Matches[1]
        }
    }
    $ips = @($ips | Sort-Object -Unique)
    if ($ips.Count -eq 0) { return 'dynamic([])' }
    return 'dynamic([' + (($ips | ForEach-Object { ConvertTo-KqlStringLiteral $_ }) -join ',') + '])'
}

function Get-LogRiskFilterKql {
    param([string]$TableName)

    $trustedIps = Get-TrustedIpKqlDynamicLiteral

    if ($TableName -eq 'AssignedLicensesDCR_CL' -or $TableName -eq 'AzureADUsersDCR_CL') {
        return ''
    }

    if ($TableName -eq 'MailboxStatisticsDCR_CL') {
        return @"
| extend __availableText = tostring(coalesce(column_ifexists("AvailableSpaceGB", ""), column_ifexists("AvailableSpaceInGB", ""), column_ifexists("AvailableSpace", "")))
| extend __quotaText = tostring(coalesce(column_ifexists("QuotaLimitGB", ""), column_ifexists("QuotaGB", ""), column_ifexists("StorageQuotaGB", ""), column_ifexists("ProhibitSendReceiveQuotaGB", "")))
| extend __available = todouble(extract(@"-?\d+(\.\d+)?", 0, __availableText)), __quota = todouble(extract(@"-?\d+(\.\d+)?", 0, __quotaText))
| extend __mailboxType = tolower(tostring(coalesce(column_ifexists("RecipientTypeDetails", ""), column_ifexists("MailboxType", ""), column_ifexists("RecipientType", ""))))
| where (__quota > 0 and __available / __quota < 0.05) or __mailboxType contains "shared"
| project-away __availableText, __quotaText, __available, __quota, __mailboxType
"@
    }

    $appAllowList = 'dynamic(["Windows Sign In","Microsoft Edge","Sangfor SASE VPN","Microsoft Office"])'

    return @"
| extend __op = tostring(coalesce(column_ifexists("OperationName", ""), column_ifexists("ActivityDisplayName", ""), column_ifexists("Operation", ""), column_ifexists("Activity", ""), column_ifexists("Action", ""), column_ifexists("EventName", ""), column_ifexists("Status", ""), column_ifexists("ResultType", "")))
| extend __status = tolower(tostring(coalesce(column_ifexists("IsSuccess", ""), column_ifexists("ResultStatus", ""), column_ifexists("Result", ""), column_ifexists("DeliveryStatus", ""), column_ifexists("Status", ""), column_ifexists("ResultType", ""))))
| extend __ipRaw = tostring(coalesce(column_ifexists("ClientIP", ""), column_ifexists("ClientIp", ""), column_ifexists("ClientIPAddress", ""), column_ifexists("IPAddress", ""), column_ifexists("IpAddress", ""), column_ifexists("SourceIP", ""), column_ifexists("FromIP", ""), column_ifexists("SenderIP", ""), column_ifexists("RemoteIP", ""), column_ifexists("IP", "")))
| extend __ip = extract(@"(?<!\d)(\d{1,3}(?:\.\d{1,3}){3})(?!\d)", 1, __ipRaw)
| extend __app = tostring(coalesce(column_ifexists("AppDisplayName", ""), column_ifexists("Application", ""), column_ifexists("ApplicationDisplayName", ""), column_ifexists("ClientAppUsed", "")))
| extend __permissionText = strcat(__op, " ", tostring(coalesce(column_ifexists("TargetResources", ""), column_ifexists("ModifiedProperties", ""), column_ifexists("ResultDescription", ""), column_ifexists("Subject", ""))))
| extend __localHour = datetime_part("hour", TimeGenerated + 8h)
| extend __isFailed = (__status in ("false","fail","failed","failure","undelivered","blocked","rejected","denied","error","timeout","quarantined","1") or (__status matches regex @"^\d+$" and toint(__status) != 0))
| extend __isDeleteDisable = tolower(__op) matches regex @"(^|[^a-z])(delete|deleted|remove|removed|disable|disabled|deactivate|deactivated)([^a-z]|$)"
| extend __isOffHours = (__localHour >= 21 or __localHour < 8)
| extend __isPublicUntrustedIp = isnotempty(__ip) and not(__ip startswith "10.") and not(__ip matches regex @"^172\.(1[6-9]|2[0-9]|3[01])\.") and not(__ip startswith "192.168.") and not(__ip startswith "127.") and not(__ip startswith "169.254.") and __ip != "0.0.0.0" and __ip != "255.255.255.255" and not(__ip in ($trustedIps))
| extend __isSigninSuspiciousSuccess = ("$TableName" == "SigninLogs" and __status in ("true","success","succeeded","0") and __isPublicUntrustedIp and not(__app in ($appAllowList)))
| extend __isMessageTraceInteresting = ("$TableName" == "MessageTraceDataDCR_CL" and __permissionText matches regex @"(?i)fail|failed|failure|blocked|quarantined|rejected|undeliver|error|timeout|Power\s*BI|PBI|skyguard")
| extend __isIdentityPermissionChange = ("$TableName" in ("AuditLogs","AADManagedIdentitySignInLogs","AADServicePrincipalSignInLogs") and __permissionText matches regex @"(?i)permission|consent|credential|secret|certificate|app role|approle|service principal|managed identity")
| where __isFailed or __isDeleteDisable or __isOffHours or __isPublicUntrustedIp or __isSigninSuspiciousSuccess or __isMessageTraceInteresting or __isIdentityPermissionChange
| project-away __op, __status, __ipRaw, __ip, __app, __permissionText, __localHour, __isFailed, __isDeleteDisable, __isOffHours, __isPublicUntrustedIp, __isSigninSuspiciousSuccess, __isMessageTraceInteresting, __isIdentityPermissionChange
"@
}

function New-AssignedLicensesOptimizedQuery {
    param(
        [string]$TableName,
        [string]$StartUtc,
        [string]$EndUtc
    )

    return @"
let __base =
$TableName
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
| extend UserPrincipalName = tostring(coalesce(column_ifexists("UserPrincipalName", ""), column_ifexists("UserUPN", ""), column_ifexists("UPN", ""), column_ifexists("Mail", ""), column_ifexists("EmailAddress", ""), column_ifexists("DisplayName", ""), column_ifexists("UserId", "")))
| extend SkuPartNumber = tostring(coalesce(column_ifexists("SkuPartNumber", ""), column_ifexists("LicenseName", ""), column_ifexists("SkuDisplayName", ""), column_ifexists("ServicePlanName", ""), column_ifexists("AssignedLicenses", ""), column_ifexists("Licenses", ""), "Unknown License"))
| extend ServicePlanName = tostring(coalesce(column_ifexists("ServicePlanName", ""), column_ifexists("LicenseName", ""), column_ifexists("SkuPartNumber", "")))
| extend LicenseName = tostring(coalesce(column_ifexists("LicenseName", ""), column_ifexists("SkuPartNumber", ""), column_ifexists("ServicePlanName", "")))
| extend ProvisioningStatus = tostring(coalesce(column_ifexists("ProvisioningStatus", ""), column_ifexists("Status", "")))
| extend TotalLicenses = todouble(tostring(coalesce(column_ifexists("TotalLicenses", ""), column_ifexists("TotalUnits", ""), column_ifexists("PrepaidUnitsEnabled", ""), column_ifexists("SkuPrepaidUnitsEnabled", ""), column_ifexists("EnabledUnits", ""), column_ifexists("Enabled", ""))));
let __summary =
__base
| summarize TimeGenerated=max(TimeGenerated), UsedUsers=dcount(UserPrincipalName), TotalLicenses=max(TotalLicenses) by SkuPartNumber
| extend UserPrincipalName="", ServicePlanName=SkuPartNumber, LicenseName=SkuPartNumber, ProvisioningStatus="Success", __RecordKind="LicenseSummary";
let __failures =
__base
| where isnotempty(ProvisioningStatus) and tolower(ProvisioningStatus) != "success"
| extend UsedUsers=long(null), __RecordKind="LicenseFailure";
__summary
| union isfuzzy=true __failures
| project TimeGenerated, UserPrincipalName, SkuPartNumber, ServicePlanName, LicenseName, ProvisioningStatus, TotalLicenses, UsedUsers, __RecordKind
| sort by TimeGenerated desc
"@
}

function Get-LogQueryExecutionMode {
    param(
        [datetime]$QueryStartTime,
        [datetime]$QueryEndTime,
        [int]$Hours = 1
    )

    if ($QueryStartTime -and $QueryEndTime) {
        return [PSCustomObject]@{
            UseTimespan = $false
            Timespan = $null
            StartTime = $QueryStartTime
            EndTime = $QueryEndTime
        }
    }

    return [PSCustomObject]@{
        UseTimespan = $true
        Timespan = New-TimeSpan -Hours $Hours
        StartTime = (Get-Date).AddHours(-$Hours)
        EndTime = Get-Date
    }
}

function New-LogAnalyzerScheduleConfig {
    param(
        [string]$RunAt = '01:00',
        [string[]]$Tables = @()
    )

    if (-not $Tables -or $Tables.Count -eq 0) {
        $Tables = @($SupportedLogTables | ForEach-Object { $_.Name })
    }

    foreach ($table in $Tables) {
        Resolve-LogTableSelection -Selection $table | Out-Null
    }

    if ($RunAt -notmatch '^([01]\d|2[0-3]):[0-5]\d$') {
        throw "RunAt must use HH:mm format, got: $RunAt"
    }

    return [PSCustomObject]@{
        RunAt = $RunAt
        Tables = [string[]]$Tables
    }
}

function Get-LogAnalyzerScheduleConfigPath {
    param([string]$RootDir)
    return Join-Path $RootDir 'schedule-config.json'
}

function Get-LogAnalyzerStatusPath {
    param([string]$RootDir)
    return Join-Path $RootDir 'schedule-status.json'
}

function Get-LogAnalyzerBatchCommand {
    param(
        [string]$RootDir,
        [string]$ConfigPath
    )

    $scriptPath = Join-Path $RootDir 'scheduled-run.ps1'
    return "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ConfigPath `"$ConfigPath`""
}

function Get-LogAnalyzerTrayCommand {
    param(
        [string]$RootDir,
        [string]$ConfigPath
    )

    $scriptPath = Join-Path $RootDir 'tray.ps1'
    return "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" -ConfigPath `"$ConfigPath`""
}

function Get-LogAnalyzerNextRunTime {
    param(
        [string]$RunAt = '01:00',
        [datetime]$Now = (Get-Date)
    )

    $parts = $RunAt.Split(':')
    $next = $Now.Date.AddHours([int]$parts[0]).AddMinutes([int]$parts[1])
    if ($next -le $Now) { $next = $next.AddDays(1) }
    return $next
}

function Get-OperationGroupValue {
    param([object]$Row, [string]$TableName)

    $profile = Get-TableAnalysisProfile -TableName $TableName
    if (-not $profile.UseCompositeOperationGroup) {
        return Get-OperationValue -Row $Row -TableName $TableName
    }

    $parts = @()
    foreach ($field in $profile.GroupFields) {
        $value = Get-FieldValue -Row $Row -Names @($field) -Default ''
        if ($value) { $parts += $value }
    }

    if ($parts.Count -eq 0) {
        return Get-OperationValue -Row $Row -TableName $TableName
    }

    return ($parts -join ' | ')
}

function Get-LogArtifactPaths {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempDir,

        [Parameter(Mandatory = $true)]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [string]$AnalysisDateStr,

        [datetime]$Now = (Get-Date)
    )

    $timestamp = $Now.ToString('HHmm')

    $repoRoot = Get-LogAnalyzerRepositoryRoot
    $htmlTarget = Join-Path $repoRoot "final_report_$($TableName)_$($AnalysisDateStr)_$timestamp.html"
    $htmlRelative = ConvertTo-RelativeLogPath -BasePath $repoRoot -TargetPath $htmlTarget

    return [PSCustomObject]@{
        CsvFile = Join-Path $TempDir "$($TableName)_$AnalysisDateStr.csv"
        HtmlFile = $htmlRelative
        HtmlFilePath = $htmlTarget
        CacheCsv = Join-Path (Join-Path $TempDir 'cache') "$($TableName)_$AnalysisDateStr.csv"
        CacheMeta = Join-Path (Join-Path $TempDir 'cache') "$($TableName)_$AnalysisDateStr.meta.json"
    }
}

function Get-MergedLogArtifactPaths {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempDir,

        [Parameter(Mandatory = $true)]
        [string]$AnalysisDateStr,

        [datetime]$Now = (Get-Date)
    )

    $timestamp = $Now.ToString('HHmm')
    $repoRoot = Get-LogAnalyzerRepositoryRoot
    $htmlTarget = Join-Path $repoRoot "final_report_merged_$($AnalysisDateStr)_$timestamp.html"
    $htmlRelative = ConvertTo-RelativeLogPath -BasePath $repoRoot -TargetPath $htmlTarget

    return [PSCustomObject]@{
        HtmlFile = $htmlRelative
        HtmlFilePath = $htmlTarget
    }
}

function New-LogTableQuery {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [datetime]$StartTime,

        [Parameter(Mandatory = $true)]
        [datetime]$EndTime,

        [switch]$RiskOnly
    )

    $validTableNames = @($SupportedLogTables | ForEach-Object { $_.Name })
    if ($validTableNames -notcontains $TableName) {
        throw "Unsupported log table: $TableName"
    }

    $start = $StartTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $end = $EndTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    if ($RiskOnly -and $TableName -eq 'AssignedLicensesDCR_CL') {
        return New-AssignedLicensesOptimizedQuery -TableName $TableName -StartUtc $start -EndUtc $end
    }
    $query = "$TableName | where TimeGenerated >= datetime($start) and TimeGenerated < datetime($end)"
    if ($RiskOnly) {
        $filter = Get-LogRiskFilterKql -TableName $TableName
        if ($filter) { $query = "$query`n$filter" }
    }
    return "$query`n| sort by TimeGenerated desc"
}
