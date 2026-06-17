$SupportedLogTables = @(
    [PSCustomObject]@{ Name = 'AADManagedIdentitySignInLogs'; Description = 'AAD Managed Identity Sign-in Logs' },
    [PSCustomObject]@{ Name = 'AADServicePrincipalSignInLogs'; Description = 'AAD Service Principal Sign-in Logs' },
    [PSCustomObject]@{ Name = 'AssignedLicensesDCR_CL'; Description = 'Assigned Licenses' },
    [PSCustomObject]@{ Name = 'AuditLogs'; Description = 'Microsoft Entra Audit Logs' },
    [PSCustomObject]@{ Name = 'DCRLogErrors'; Description = 'DCR Log Errors' },
    [PSCustomObject]@{ Name = 'IntuneAuditLogsDCR_CL'; Description = 'Intune Audit Logs' },
    [PSCustomObject]@{ Name = 'MailboxStatisticsDCR_CL'; Description = 'Mailbox Statistics' },
    [PSCustomObject]@{ Name = 'SigninLogs'; Description = 'Microsoft Entra Sign-in Logs' }
)

function Get-SupportedLogTables {
    return $SupportedLogTables
}

function Resolve-LogTableSelnn {
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

    if ($Days -lt 0 -or $Days -gt 90) {
        throw 'Days must be between 0 and 90. Use 0 for the last 3 hours.'
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

        [int]$MaxDays = 90
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
    Write-Host '  n. Last n days (max 90)' -ForegroundColor Cyan
    $selection = Read-Host 'Enter 0, 1, or n'

    if ([string]::IsNullOrWhiteSpace($selection)) {
        $selection = '0'
    }

    $days = 0
    if (-not [int]::TryParse($selection.Trim(), [ref]$days)) {
        throw "Invalid time range selection: $selection"
    }
    if ($days -lt 0 -or $days -gt 90) {
        throw 'Time range number must be between 0 and 90.'
    }

    $roundedNow = [datetime]::new($Now.Year, $Now.Month, $Now.Day, $Now.Hour, 0, 0)
    return Get-RelativeLogTimeRange -Now $roundedNow -Days $days
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
    $currentDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

    while ($currentDir) {
        $hasSkill = Test-Path (Join-Path $currentDir 'SKILL.md')
        $hasScripts = Test-Path (Join-Path $currentDir 'scripts')
        if ($hasSkill -and $hasScripts) {
            return $currentDir
        }

        $parentDir = Split-Path -Parent $currentDir
        if (-not $parentDir -or $parentDir -eq $currentDir) {
            break
        }
        $currentDir = $parentDir
    }

    return Split-Path -Parent $PSScriptRoot
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
            $common.OperationFields = @('ResultSignature', 'ResultType', 'ResultDescription', 'Status', 'ConditionalAccessStatus', 'OperationName', 'ActivityDisplayName') + $common.OperationFields
            $common.WorkloadFields = @('ResourceDisplayName', 'ResourceServicePrincipalId', 'ServicePrincipalName', 'AppDisplayName', 'Category') + $common.WorkloadFields
            $common.ClientIpFields = @('IPAddress', 'IpAddress', 'ClientIP', 'ClientIpAddress') + $common.ClientIpFields
            $common.SuccessFields = @('ResultSignature', 'ResultType', 'Status', 'ResultDescription') + $common.SuccessFields
            $common.DefaultOperation = 'Managed Identity Sign-in'
            $common.DefaultWorkload = 'Microsoft Entra Managed Identity'
            $common.DefaultSuccess = 'unknown'
        }
        'AADServicePrincipalSignInLogs' {
            $common.UserFields = @('ServicePrincipalName', 'AppDisplayName', 'ServicePrincipalId', 'AppId', 'ResourceDisplayName') + $common.UserFields
            $common.OperationFields = @('ServicePrincipalName', 'OperationName', 'ActivityDisplayName', 'ResultSignature', 'ResultType', 'ResultDescription', 'Status', 'ConditionalAccessStatus') + $common.OperationFields
            $common.WorkloadFields = @('ResourceDisplayName', 'ResourceServicePrincipalId', 'AppDisplayName', 'Category') + $common.WorkloadFields
            $common.ClientIpFields = @('IPAddress', 'IpAddress', 'ClientIP', 'ClientIpAddress') + $common.ClientIpFields
            $common.SuccessFields = @('ResultSignature', 'ResultType', 'Status', 'ResultDescription') + $common.SuccessFields
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
            $common.UserFields = @('Actor', 'InitiatedByUserPrincipalName', 'ActorUserPrincipalName', 'UserPrincipalName', 'Identity', 'InitiatedBy') + $common.UserFields
            $common.OperationFields = @('OperationName', 'ActivityDisplayName', 'Activity', 'Operation', 'Category', 'Result') + $common.OperationFields
            $common.WorkloadFields = @('Category', 'LoggedByService', 'Service', 'ResourceProvider', 'AADOperationType') + $common.WorkloadFields
            $common.ClientIpFields = @('IPAddress', 'IpAddress', 'ClientIP', 'ClientIpAddress') + $common.ClientIpFields
            $common.SuccessFields = @('Result', 'ResultType', 'Status', 'ActivityStatus') + $common.SuccessFields
            $common.DefaultOperation = 'Audit Log Event'
            $common.DefaultWorkload = 'Microsoft Entra Audit'
            $common.DefaultSuccess = 'unknown'
        }
        'DCRLogErrors' {
            $common.UserFields = @('InputStreamId', 'StreamName', 'DataCollectionRuleId') + $common.UserFields
            $common.OperationFields = @('OperationName', 'Operation', 'Category') + $common.OperationFields
            $common.WorkloadFields = @('InputStreamId', 'DataCollectionRuleId', 'SourceSystem') + $common.WorkloadFields
            $common.SuccessFields = @('Status', 'Result', 'Level') + $common.SuccessFields
            $common.DefaultUser = 'DCR'
            $common.DefaultOperation = 'DCR Log Error'
            $common.DefaultWorkload = 'Data Collection Rule'
            $common.DefaultSuccess = 'false'
        }
        'IntuneAuditLogsDCR_CL' {
            $common.UserFields = @('ActorUPN', 'ActorUPN_s', 'ActorUserPrincipalName', 'ActorUserPrincipalName_s', 'Actor', 'Actor_s', 'InitiatedByUserPrincipalName', 'InitiatedByUserPrincipalName_s', 'UserPrincipalName', 'UserPrincipalName_s', 'UPN', 'UPN_s', 'UserId', 'UserId_s', 'Identity', 'Identity_s') + $common.UserFields
            $common.OperationFields = @('OperationName', 'OperationName_s', 'ActivityDisplayName', 'ActivityDisplayName_s', 'Activity', 'Activity_s', 'Operation', 'Operation_s', 'Action', 'Action_s', 'AuditEventType', 'AuditEventType_s') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'Category', 'Service', 'SourceSystem') + $common.WorkloadFields
            $common.ClientIpFields = @('IPAddress', 'IpAddress', 'ClientIP', 'ClientIpAddress') + $common.ClientIpFields
            $common.SuccessFields = @('Result', 'Result_s', 'ResultStatus', 'ResultStatus_s', 'Status', 'Status_s', 'ActivityResult', 'ActivityResult_s', 'OperationStatus', 'OperationStatus_s') + $common.SuccessFields
            $common.DefaultUser = 'Intune'
            $common.DefaultOperation = 'Intune Audit Event'
            $common.DefaultWorkload = 'Intune'
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
            $common.OperationFields = @('AppDisplayName', 'ServicePrincipalName', 'ResultSignature', 'ResultType', 'ResultDescription', 'Status', 'ConditionalAccessStatus') + $common.OperationFields
            $common.WorkloadFields = @('AppDisplayName', 'ServicePrincipalName', 'ResourceDisplayName', 'ClientAppUsed', 'Category') + $common.WorkloadFields
            $common.ClientIpFields = @('IPAddress', 'IpAddress', 'ClientIP', 'ClientIpAddress') + $common.ClientIpFields
            $common.SuccessFields = @('ResultSignature', 'ResultType', 'Status', 'ResultDescription') + $common.SuccessFields
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

function Join-DisplayIdentityValue {
    param(
        [string]$DisplayName,
        [string]$Identity
    )

    $display = if ($DisplayName) { $DisplayName.Trim() } else { '' }
    $id = if ($Identity) { $Identity.Trim() } else { '' }
    if (-not $display) { return $id }
    if (-not $id) { return $display }
    if ($display -eq $id -or $display.Contains($id)) { return $display }
    if ($display -match '@' -and $id -notmatch '@') { return "$id / $display" }
    return "$display / $id"
}

function Get-InitiatedByUserIdentity {
    param([object]$Row)

    if ($Row.PSObject.Properties.Name -notcontains 'InitiatedBy') { return '' }
    $raw = [string]$Row.InitiatedBy
    if ([string]::IsNullOrWhiteSpace($raw)) { return '' }
    try {
        $json = $raw | ConvertFrom-Json
        if ($json.user) {
            return Join-DisplayIdentityValue -DisplayName ([string]$json.user.displayName) -Identity ([string]$json.user.userPrincipalName)
        }
    } catch {
    }
    return ''
}

function Get-DisplayIdentityValue {
    param(
        [object]$Row,
        [string[]]$DisplayFields,
        [string[]]$IdentityFields
    )

    $display = Get-FieldValue -Row $Row -Names $DisplayFields -Default ''
    $identity = Get-FieldValue -Row $Row -Names $IdentityFields -Default ''
    return Join-DisplayIdentityValue -DisplayName $display -Identity $identity
}

function Get-UserValue {
    param([object]$Row, [string]$TableName)
    $profile = Get-TableAnalysisProfile -TableName $TableName

    switch ($TableName) {
        'AuditLogs' {
            $actor = Get-FieldValue -Row $Row -Names @('Actor') -Default ''
            if ($actor -and $actor -match '\s/\s|@') {
                if ($actor -match '\s/\s') { return $actor }
            }
            $fromInitiatedBy = Get-InitiatedByUserIdentity -Row $Row
            if ($fromInitiatedBy) { return $fromInitiatedBy }
            $combined = Get-DisplayIdentityValue -Row $Row -DisplayFields @('ActorDisplayName', 'InitiatedByUserDisplayName', 'UserDisplayName', 'DisplayName', 'Identity') -IdentityFields @('ActorUserPrincipalName', 'InitiatedByUserPrincipalName', 'UserPrincipalName', 'UserUPN', 'UPN', 'UserId')
            if ($combined) { return $combined }
        }
        'SigninLogs' {
            $combined = Get-DisplayIdentityValue -Row $Row -DisplayFields @('UserDisplayName', 'DisplayName') -IdentityFields @('UserPrincipalName', 'Identity', 'User', 'UserId')
            if ($combined) { return $combined }
        }
        'IntuneAuditLogsDCR_CL' {
            $combined = Get-DisplayIdentityValue -Row $Row -DisplayFields @('ActorDisplayName', 'ActorDisplayName_s', 'DisplayName', 'DisplayName_s') -IdentityFields @('ActorUPN', 'ActorUPN_s', 'ActorUserPrincipalName', 'ActorUserPrincipalName_s', 'InitiatedByUserPrincipalName', 'InitiatedByUserPrincipalName_s', 'UserPrincipalName', 'UserPrincipalName_s', 'UPN', 'UPN_s', 'UserId', 'UserId_s', 'Actor', 'Actor_s')
            if ($combined) { return $combined }
        }
        'AuditGeneralDCR_CL' {
            $combined = Get-DisplayIdentityValue -Row $Row -DisplayFields @('ActorDisplayName', 'UserDisplayName', 'DisplayName', 'Actor') -IdentityFields @('ActorUserPrincipalName', 'UserPrincipalName', 'UserUPN', 'UPN', 'UserId')
            if ($combined) { return $combined }
        }
    }

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
    if ($TableName -in @('AADManagedIdentitySignInLogs', 'AADServicePrincipalSignInLogs', 'SigninLogs') -and -not [string]::IsNullOrWhiteSpace($value) -and $value -ne 'unknown') { return 'false' }
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
        (Join-Path $repoRoot 'scripts\config\TrustedLocation_KJ.txt'),
        (Join-Path $repoRoot 'scripts\config\TrustedLocation_IDC_Ali.txt')
    )
}

function Get-MicrosoftServiceTagsCachePath {
    $repoRoot = Get-LogAnalyzerRepositoryRoot
    $cacheDir = Join-Path $repoRoot 'scripts\cache'
    return Join-Path $cacheDir 'microsoft-service-tags.json'
}

function Update-MicrosoftServiceTagsCache {
    param(
        [string]$CachePath = (Get-MicrosoftServiceTagsCachePath)
    )

    $cacheDir = Split-Path -Parent $CachePath
    if (-not (Test-Path $cacheDir)) {
        New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null
    }

    $downloadPages = @(
        'https://www.microsoft.com/en-us/download/confirmation.aspx?id=57062',
        'https://www.microsoft.com/en-us/download/confirmation.aspx?id=56519'
    )

    foreach ($page in $downloadPages) {
        try {
            $html = Invoke-WebRequest -Uri $page -UseBasicParsing -TimeoutSec 30
            $match = [regex]::Match($html.Content, 'https://download\.microsoft\.com/download/[^"]+ServiceTags_[^"]+\.json')
            if ($match.Success) {
                Invoke-WebRequest -Uri $match.Value -OutFile $CachePath -UseBasicParsing -TimeoutSec 60
                return [PSCustomObject]@{ Success = $true; Path = $CachePath; Source = $match.Value; Error = '' }
            }
        } catch {
            $lastError = $_.Exception.Message
        }
    }

    return [PSCustomObject]@{ Success = $false; Path = $CachePath; Source = ''; Error = $lastError }
}

function Get-MicrosoftServiceTagCidrs {
    param(
        [string]$CachePath = (Get-MicrosoftServiceTagsCachePath),
        [switch]$Refresh,
        [string]$TagNamePattern = '^(AzureActiveDirectory(\.|$)|AzureFrontDoor\.(Frontend|FirstParty|MicrosoftSecurity|Backend)|MicrosoftCloudAppSecurity(\.|$)|MicrosoftDefenderForEndpoint(\.|$)|PowerBI(\.|$)|AzurePortal|AzureResourceManager(\.|$)|AzureTrafficManager)$'
    )

    if (-not $Refresh -and $script:MicrosoftServiceTagCidrsCache) {
        return $script:MicrosoftServiceTagCidrsCache
    }

    if ($Refresh) {
        Update-MicrosoftServiceTagsCache -CachePath $CachePath | Out-Null
    }

    if (-not (Test-Path $CachePath)) {
        Update-MicrosoftServiceTagsCache -CachePath $CachePath | Out-Null
    }

    if (-not (Test-Path $CachePath)) { return @() }

    try {
        $json = Get-Content -Path $CachePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $cidrs = [System.Collections.Generic.List[string]]::new()
        foreach ($entry in @($json.values)) {
            $name = [string]$entry.name
            if ($TagNamePattern -and $name -notmatch $TagNamePattern) { continue }
            foreach ($prefix in @($entry.properties.addressPrefixes)) {
                if ($prefix -match '^\d{1,3}(\.\d{1,3}){3}/\d{1,2}$') {
                    $cidrs.Add($prefix) | Out-Null
                }
            }
        }
        $result = @($cidrs.ToArray() | Sort-Object -Unique)
        if (-not $Refresh) {
            $script:MicrosoftServiceTagCidrsCache = $result
        }
        return $result
    } catch {
        return @()
    }
}

function Get-TrustedIpRules {
    param(
        [string[]]$Paths = (Get-TrustedLocationPaths),
        [switch]$IncludeMicrosoft = $true
    )

    $useDefaultCache = (-not $PSBoundParameters.ContainsKey('Paths')) -and $IncludeMicrosoft
    if ($useDefaultCache -and $script:TrustedIpRulesCache) {
        return $script:TrustedIpRulesCache
    }

    $rules = [System.Collections.Generic.List[object]]::new()
    $cidrValues = [System.Collections.Generic.List[object]]::new()
    foreach ($path in $Paths) {
        if (-not (Test-Path $path)) { continue }
        foreach ($line in Get-Content -Path $path -Encoding UTF8) {
            $value = ($line -replace '#.*$', '').Trim()
            if (-not $value) { continue }
            $cidrValues.Add([PSCustomObject]@{ Cidr = $value; Source = Split-Path -Leaf $path }) | Out-Null
        }
    }

    if ($IncludeMicrosoft) {
        foreach ($cidr in Get-MicrosoftServiceTagCidrs) {
            $cidrValues.Add([PSCustomObject]@{ Cidr = $cidr; Source = 'MicrosoftServiceTags' }) | Out-Null
        }
    }

    foreach ($item in $cidrValues) {
            $value = $item.Cidr
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
                Source = $item.Source
                Cidr = $value
                Address = $addr
                Mask = $mask
            }) | Out-Null
    }

    $result = @($rules.ToArray())
    if ($useDefaultCache) {
        $script:TrustedIpRulesCache = $result
    }
    return $result
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

function ConvertTo-KqlStringLiteral {
    param([string]$Value)
    return '"' + (($Value -replace '\\', '\\') -replace '"', '\"') + '"'
}

function Get-TrustedIpKqlDynamicLiteral {
    if ($script:TrustedIpKqlDynamicLiteralCache) {
        return $script:TrustedIpKqlDynamicLiteralCache
    }

    $rules = @(Get-TrustedIpRules)
    $cidrs = @()
    foreach ($rule in $rules) {
        if ($rule.Cidr -match '^\d{1,3}(\.\d{1,3}){3}/\d{1,2}$') {
            $cidrs += $rule.Cidr
        }
    }
    $cidrs = @($cidrs | Sort-Object -Unique)
    if ($cidrs.Count -eq 0) {
        $script:TrustedIpKqlDynamicLiteralCache = 'dynamic([])'
        return $script:TrustedIpKqlDynamicLiteralCache
    }
    $script:TrustedIpKqlDynamicLiteralCache = 'dynamic([' + (($cidrs | ForEach-Object { ConvertTo-KqlStringLiteral $_ }) -join ',') + '])'
    return $script:TrustedIpKqlDynamicLiteralCache
}

function Get-LogRiskFilterKql {
    param([string]$TableName)

    if ($TableName -eq 'AssignedLicensesDCR_CL' -or $TableName -eq 'AzureADUsersDCR_CL') {
        return ''
    }

    if ($TableName -eq 'MailboxStatisticsDCR_CL') {
        return @"
| extend __availableText = strcat(" ", tostring(column_ifexists("AvailableSpaceGB", "")), " ", tostring(column_ifexists("AvailableSpaceGB_d", "")), " ", tostring(column_ifexists("AvailableSpaceGB_r", "")), " ", tostring(column_ifexists("AvailableSpaceGB_s", "")), " ", tostring(column_ifexists("AvailableSpaceInGB", "")), " ", tostring(column_ifexists("AvailableSpaceInGB_d", "")), " ", tostring(column_ifexists("AvailableSpaceInGB_r", "")), " ", tostring(column_ifexists("AvailableSpaceInGB_s", "")), " ", tostring(column_ifexists("AvailableSpace", "")), " ", tostring(column_ifexists("AvailableSpace_d", "")), " ", tostring(column_ifexists("AvailableSpace_r", "")), " ", tostring(column_ifexists("AvailableSpace_s", "")))
| extend __quotaText = strcat(" ", tostring(column_ifexists("QuotaLimitGB", "")), " ", tostring(column_ifexists("QuotaLimitGB_d", "")), " ", tostring(column_ifexists("QuotaLimitGB_r", "")), " ", tostring(column_ifexists("QuotaLimitGB_s", "")), " ", tostring(column_ifexists("QuotaGB", "")), " ", tostring(column_ifexists("QuotaGB_d", "")), " ", tostring(column_ifexists("QuotaGB_r", "")), " ", tostring(column_ifexists("QuotaGB_s", "")), " ", tostring(column_ifexists("StorageQuotaGB", "")), " ", tostring(column_ifexists("StorageQuotaGB_d", "")), " ", tostring(column_ifexists("StorageQuotaGB_r", "")), " ", tostring(column_ifexists("StorageQuotaGB_s", "")), " ", tostring(column_ifexists("ProhibitSendReceiveQuotaGB", "")), " ", tostring(column_ifexists("ProhibitSendReceiveQuotaGB_d", "")), " ", tostring(column_ifexists("ProhibitSendReceiveQuotaGB_r", "")), " ", tostring(column_ifexists("ProhibitSendReceiveQuotaGB_s", "")))
| extend __available = todouble(extract(@"-?\d+(\.\d+)?", 0, __availableText)), __quota = todouble(extract(@"-?\d+(\.\d+)?", 0, __quotaText))
| extend __mailboxTypeText = strcat(" ", tostring(column_ifexists("RecipientTypeDetails", "")), " ", tostring(column_ifexists("RecipientTypeDetails_s", "")), " ", tostring(column_ifexists("RecipientTypeDetail", "")), " ", tostring(column_ifexists("RecipientTypeDetail_s", "")), " ", tostring(column_ifexists("MailboxRecipientType", "")), " ", tostring(column_ifexists("MailboxRecipientType_s", "")), " ", tostring(column_ifexists("MailboxType", "")), " ", tostring(column_ifexists("MailboxType_s", "")), " ", tostring(column_ifexists("RecipientType", "")), " ", tostring(column_ifexists("RecipientType_s", "")), " ", tostring(column_ifexists("IsSharedMailbox", "")), " ", tostring(column_ifexists("IsSharedMailbox_s", "")), " ", tostring(column_ifexists("IsSharedMailBox", "")), " ", tostring(column_ifexists("IsSharedMailBox_s", "")), " ", tostring(column_ifexists("IsShared", "")), " ", tostring(column_ifexists("IsShared_s", "")), " ", tostring(column_ifexists("SharedMailbox", "")), " ", tostring(column_ifexists("SharedMailbox_s", "")), " ", tostring(column_ifexists("SharedMailBox", "")), " ", tostring(column_ifexists("SharedMailBox_s", "")))
| extend __isLowSpace = (__quota > 0 and __available < __quota * 0.05)
| extend __isSharedMailbox = tolower(__mailboxTypeText) contains "shared"
| where __isLowSpace or __isSharedMailbox
| project-away __availableText, __quotaText, __available, __quota, __mailboxTypeText, __isLowSpace, __isSharedMailbox
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
| extend __isFailed = (__status in ("false","fail","failed","failure","undelivered","blocked","rejected","denied","error","timeout","quarantined","1") or (__status matches regex @"^\d+$" and toint(__status) != 0))
| extend __isSuccess = (__status in ("true","success","succeeded","completed","complete","ok","pass","passed","0"))
| extend __isDeleteDisable = tolower(__op) matches regex @"(^|[^a-z])(delete|deleted|remove|removed|disable|disabled|deactivate|deactivated)([^a-z]|$)"
| extend __isSigninSuspiciousSuccess = ("$TableName" == "SigninLogs" and __status in ("true","success","succeeded","0") and not(__app in ($appAllowList)))
| extend __isMessageTraceCritical = ("$TableName" == "MessageTraceDataDCR_CL" and __permissionText matches regex @"(?i)\b(fail(ed|ure)?|blocked|quarantined|reject(ed)?|undeliver(ed|able)?|error|timeout|bounced)\b")
| extend __isMessageTraceBusiness = ("$TableName" == "MessageTraceDataDCR_CL" and __permissionText matches regex @"(?i)\b(Power\s*BI|PBI|SkyGuard)\b" and __permissionText matches regex @"(?i)\b(fail(ed|ure)?|blocked|quarantined|reject(ed)?|undeliver(ed|able)?|error|timeout|bounced)\b")
| extend __isMessageTraceInteresting = (__isMessageTraceCritical or __isMessageTraceBusiness)
| extend __isIdentityPermissionChange = ("$TableName" == "AuditLogs" and __isSuccess and __permissionText matches regex @"(?i)permission|consent|credential|secret|certificate|app role|approle|service principal|managed identity")
| where __isFailed or __isDeleteDisable or __isSigninSuspiciousSuccess or __isMessageTraceInteresting or __isIdentityPermissionChange
| project-away __op, __status, __ipRaw, __ip, __app, __permissionText, __isFailed, __isSuccess, __isDeleteDisable, __isSigninSuspiciousSuccess, __isMessageTraceCritical, __isMessageTraceBusiness, __isMessageTraceInteresting, __isIdentityPermissionChange
"@
}

function New-AssignedLicensesOptimizedQuery {
    param(
        [string]$TableName,
        [string]$StartUtc,
        [string]$EndUtc
    )

    return @"
$TableName
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
| extend UserPrincipalName = tostring(coalesce(column_ifexists("UserPrincipalName", ""), column_ifexists("UserUPN", ""), column_ifexists("UPN", ""), column_ifexists("Mail", ""), column_ifexists("EmailAddress", ""), column_ifexists("DisplayName", ""), column_ifexists("UserId", "")))
| extend SkuPartNumber = tostring(coalesce(column_ifexists("SkuPartNumber", ""), column_ifexists("LicenseName", ""), column_ifexists("SkuDisplayName", ""), column_ifexists("ServicePlanName", ""), column_ifexists("AssignedLicenses", ""), column_ifexists("Licenses", ""), "Unknown License"))
| extend ServicePlanName = tostring(coalesce(column_ifexists("ServicePlanName", ""), column_ifexists("LicenseName", ""), column_ifexists("SkuPartNumber", "")))
| extend LicenseName = tostring(coalesce(column_ifexists("LicenseName", ""), column_ifexists("SkuPartNumber", ""), column_ifexists("ServicePlanName", "")))
| extend ProvisioningStatus = tostring(coalesce(column_ifexists("ProvisioningStatus", ""), column_ifexists("Status", ""), "Success"))
| extend TotalLicenses = todouble(tostring(coalesce(column_ifexists("TotalLicenses", ""), column_ifexists("TotalUnits", ""), column_ifexists("PrepaidUnitsEnabled", ""), column_ifexists("SkuPrepaidUnitsEnabled", ""), column_ifexists("EnabledUnits", ""), column_ifexists("Enabled", ""))))
| summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count(), UsedUsers=dcount(UserPrincipalName), TotalLicenses=max(TotalLicenses) by SkuPartNumber, ServicePlanName, LicenseName, ProvisioningStatus
| extend UserPrincipalName="", __RecordKind="LicenseSummary"
| project TimeGenerated, FirstTime, LastTime, EventCount, UserPrincipalName, SkuPartNumber, ServicePlanName, LicenseName, ProvisioningStatus, TotalLicenses, UsedUsers, __RecordKind
"@
}

function New-AadIdentitySigninOptimizedQuery {
    param(
        [string]$TableName,
        [string]$StartUtc,
        [string]$EndUtc
    )

    $failedThresholdClause = '| where __RecordKind != "AggregatedFailedSignin" or EventCount > 10'
    return @"
$TableName
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
| extend __ServicePrincipalNameRaw = tostring(column_ifexists("ServicePrincipalName", ""))
| extend __ManagedIdentityNameRaw = tostring(column_ifexists("ManagedIdentityName", ""))
| extend __IdentityRaw = tostring(column_ifexists("Identity", ""))
| extend __AppDisplayNameRaw = tostring(column_ifexists("AppDisplayName", ""))
| extend __ServicePrincipalIdRaw = tostring(column_ifexists("ServicePrincipalId", ""))
| extend __AppIdRaw = tostring(column_ifexists("AppId", ""))
| extend ServicePrincipalName = iff(isnotempty(__ServicePrincipalNameRaw), __ServicePrincipalNameRaw, iff(isnotempty(__ManagedIdentityNameRaw), __ManagedIdentityNameRaw, iff(isnotempty(__IdentityRaw), __IdentityRaw, iff(isnotempty(__AppDisplayNameRaw), __AppDisplayNameRaw, iff(isnotempty(__ServicePrincipalIdRaw), __ServicePrincipalIdRaw, iff(isnotempty(__AppIdRaw), __AppIdRaw, "Unknown"))))))
| extend UserPrincipalName = ServicePrincipalName
| extend __ResourceDisplayNameRaw = tostring(column_ifexists("ResourceDisplayName", ""))
| extend __ResourceIdentityRaw = tostring(column_ifexists("ResourceIdentity", ""))
| extend __ResourceServicePrincipalIdRaw = tostring(column_ifexists("ResourceServicePrincipalId", ""))
| extend ResourceDisplayName = iff(isnotempty(__ResourceDisplayNameRaw), __ResourceDisplayNameRaw, iff(isnotempty(__ResourceIdentityRaw), __ResourceIdentityRaw, __ResourceServicePrincipalIdRaw))
| extend AppDisplayName = ServicePrincipalName
| extend __IPAddressRaw = tostring(column_ifexists("IPAddress", ""))
| extend __IpAddressRaw = tostring(column_ifexists("IpAddress", ""))
| extend __ClientIPRaw = tostring(column_ifexists("ClientIP", ""))
| extend __ClientIpAddressRaw = tostring(column_ifexists("ClientIpAddress", ""))
| extend IPAddress = iff(isnotempty(__IPAddressRaw), __IPAddressRaw, iff(isnotempty(__IpAddressRaw), __IpAddressRaw, iff(isnotempty(__ClientIPRaw), __ClientIPRaw, __ClientIpAddressRaw)))
| extend __ResultSignatureRaw = tostring(column_ifexists("ResultSignature", ""))
| extend __ResultRaw = tostring(column_ifexists("Result", ""))
| extend __ResultTypeRaw = tostring(column_ifexists("ResultType", ""))
| extend __StatusRaw = tostring(column_ifexists("Status", ""))
| extend __ResultDescriptionRaw = tostring(column_ifexists("ResultDescription", ""))
| extend __FailureReasonRaw = tostring(column_ifexists("FailureReason", ""))
| extend __ConditionalAccessStatusRaw = tostring(column_ifexists("ConditionalAccessStatus", ""))
| extend ResultSignature = iff(isnotempty(__ResultSignatureRaw), __ResultSignatureRaw, __ResultRaw)
| extend ResultType = iff(isnotempty(ResultSignature), ResultSignature, iff(isnotempty(__ResultTypeRaw), __ResultTypeRaw, iff(isnotempty(__StatusRaw), __StatusRaw, __ResultDescriptionRaw)))
| extend ResultDescription = iff(isnotempty(__ResultDescriptionRaw), __ResultDescriptionRaw, iff(isnotempty(__FailureReasonRaw), __FailureReasonRaw, iff(isnotempty(__StatusRaw), __StatusRaw, __ConditionalAccessStatusRaw)))
| extend __status = tolower(ResultType)
| extend __isSuccess = (__status in ("true","success","succeeded","completed","complete","ok","pass","passed","0"))
| extend __resultDescriptionLower = tolower(ResultDescription)
| extend __isFailed = (isnotempty(__status) and not(__isSuccess)) or __resultDescriptionLower contains "fail" or __resultDescriptionLower contains "failure" or __resultDescriptionLower contains "denied" or __resultDescriptionLower contains "error" or __resultDescriptionLower contains "timeout"
| extend __ip = extract(@"(\d{1,3}(?:\.\d{1,3}){3})", 1, IPAddress)
| where __isFailed or __isSuccess
| extend __RecordKind=iff(__isFailed, "AggregatedFailedSignin", "AggregatedSuspiciousSigninSuccess")
| summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by UserPrincipalName, ServicePrincipalName, ResourceDisplayName, AppDisplayName, IPAddress, ResultType, ResultDescription, __RecordKind
$failedThresholdClause
| extend OperationName=ServicePrincipalName, Status=ResultType
| project TimeGenerated, FirstTime, LastTime, EventCount, UserPrincipalName, ServicePrincipalName, ResourceDisplayName, AppDisplayName, OperationName, IPAddress, ResultType, ResultDescription, Status, __RecordKind
"@
}

function New-SigninLogsOptimizedQuery {
    param(
        [string]$StartUtc,
        [string]$EndUtc
    )

    return @"
SigninLogs
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
| extend UserDisplayName = tostring(column_ifexists("UserDisplayName", ""))
| extend UserPrincipalName = tostring(column_ifexists("UserPrincipalName", ""))
| extend AppDisplayName = tostring(column_ifexists("AppDisplayName", ""))
| extend IPAddress = tostring(column_ifexists("IPAddress", ""))
| extend ResultType = tostring(column_ifexists("ResultType", ""))
| extend ResultDescription = tostring(column_ifexists("ResultDescription", ""))
| extend UserPrincipalName = iff(isnotempty(UserPrincipalName), UserPrincipalName, iff(isnotempty(UserDisplayName), UserDisplayName, "Unknown"))
| extend AppDisplayName = iff(isnotempty(AppDisplayName), AppDisplayName, "Unknown")
| extend __status = tolower(ResultType)
| extend __isSuccess = (__status in ("true","success","succeeded","completed","complete","ok","pass","passed","0"))
| extend __resultDescriptionLower = tolower(ResultDescription)
| extend __isFailed = (isnotempty(__status) and not(__isSuccess) and __status != "unknown") or __resultDescriptionLower contains "fail" or __resultDescriptionLower contains "failure" or __resultDescriptionLower contains "denied" or __resultDescriptionLower contains "error" or __resultDescriptionLower contains "timeout"
| extend __ip = extract(@"(\d{1,3}(?:\.\d{1,3}){3})", 1, IPAddress)
| extend __appLower = tolower(AppDisplayName)
| extend __isAllowedApp = __appLower in ("windows sign in", "microsoft edge", "sangfor sase vpn", "microsoft office")
| extend __isSigninSuspiciousSuccess = (__isSuccess and not(__isAllowedApp))
| where __isFailed or __isSigninSuspiciousSuccess
| extend __RecordKind=iff(__isFailed, "AggregatedFailedSignin", "AggregatedSuspiciousSigninSuccess")
| summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by UserPrincipalName, UserDisplayName, AppDisplayName, IPAddress, ResultType, ResultDescription, __RecordKind
| extend OperationName=AppDisplayName, Status=ResultType
| project TimeGenerated, FirstTime, LastTime, EventCount, UserPrincipalName, UserDisplayName, AppDisplayName, OperationName, IPAddress, ResultType, ResultDescription, Status, __RecordKind
"@
}

function New-AuditActivityOptimizedQuery {
    param(
        [string]$TableName,
        [string]$StartUtc,
        [string]$EndUtc
    )

    return @"
$TableName
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
| extend UserDisplayName = tostring(coalesce(column_ifexists("UserDisplayName", ""), column_ifexists("ActorDisplayName", ""), column_ifexists("DisplayName", ""), ""))
| extend UserUPN = tostring(coalesce(column_ifexists("UserUPN", ""), column_ifexists("UserPrincipalName", ""), column_ifexists("ActorUserPrincipalName", ""), column_ifexists("Actor", ""), column_ifexists("UserId", ""), "Unknown"))
| extend ActorDisplayName = UserDisplayName
| extend ActorUserPrincipalName = UserUPN
| extend UserId = tostring(coalesce(column_ifexists("UserId", ""), column_ifexists("ActorId", ""), column_ifexists("UserUPN", ""), column_ifexists("UserPrincipalName", ""), "Unknown"))
| extend Activity = tostring(coalesce(column_ifexists("Activity", ""), column_ifexists("Operation", ""), column_ifexists("OperationName", ""), column_ifexists("EventName", ""), "Unknown"))
| extend Operation = tostring(coalesce(column_ifexists("Operation", ""), column_ifexists("Activity", ""), column_ifexists("OperationName", ""), column_ifexists("EventName", ""), "Unknown"))
| extend Workload = tostring(coalesce(column_ifexists("Workload", ""), column_ifexists("RecordType", ""), column_ifexists("Service", ""), column_ifexists("SourceSystem", ""), "Unknown"))
| extend ClientIP = tostring(coalesce(column_ifexists("ClientIP", ""), column_ifexists("ClientIp", ""), column_ifexists("ClientIPAddress", ""), column_ifexists("IPAddress", ""), column_ifexists("SourceIP", ""), ""))
| extend IsSuccess = tostring(coalesce(column_ifexists("IsSuccess", ""), column_ifexists("ResultStatus", ""), column_ifexists("Result", ""), column_ifexists("Status", ""), ""))
| extend ResultStatus = IsSuccess
| extend ResultDescription = tostring(coalesce(column_ifexists("ResultDescription", ""), column_ifexists("ResultReason", ""), column_ifexists("Status", ""), ""))
| extend __status = tolower(IsSuccess)
| extend __op = strcat(Activity, " ", Operation)
| extend __isFailed = (__status in ("false","fail","failed","failure","denied","error","timeout","1") or (__status matches regex @"^\d+$" and toint(__status) != 0))
| extend __isDeleteDisable = tolower(__op) matches regex @"(^|[^a-z])(delete|deleted|remove|removed|disable|disabled|deactivate|deactivated)([^a-z]|$)"
| where __isFailed or __isDeleteDisable
| extend __RecordKind = iff(__isDeleteDisable, "AggregatedDeleteDisable", "AggregatedFailedOperation")
| summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by UserUPN, UserDisplayName, ActorDisplayName, ActorUserPrincipalName, UserId, Activity, Operation, Workload, ClientIP, IsSuccess, ResultStatus, ResultDescription, __RecordKind
| project TimeGenerated, FirstTime, LastTime, EventCount, UserUPN, UserDisplayName, ActorDisplayName, ActorUserPrincipalName, UserId, Activity, Operation, Workload, ClientIP, IsSuccess, ResultStatus, ResultDescription, __RecordKind
"@
}

function New-AuditLogsOptimizedQuery {
    param(
        [string]$StartUtc,
        [string]$EndUtc
    )

    return @"
AuditLogs
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
| where tostring(Result) =~ "success"
| where OperationName in (
    "Add app role assignment to service principal",
    "Add app role assignment to user",
    "Add app role assignment to group",
    "Add delegated permission grant",
    "Add application",
    "Update application",
    "Consent to application",
    "Add owner to application",
    "Remove app role assignment from service principal",
    "Remove delegated permission grant",
    "Add service principal",
    "Update service principal",
    "Delete application",
    "Delete service principal"
)
| project TimeGenerated, 
    OperationName, 
    Actor = tostring(InitiatedBy.user.userPrincipalName),
    Target = tostring(TargetResources),
    Result,
    CorrelationId
| order by TimeGenerated desc
"@
}

function New-MessageTraceOptimizedQuery {
    param(
        [string]$StartUtc,
        [string]$EndUtc
    )

    return @"
MessageTraceDataDCR_CL
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
| extend SenderAddress = tostring(coalesce(column_ifexists("SenderAddress", ""), column_ifexists("Sender", ""), column_ifexists("From", ""), "Unknown"))
| extend RecipientAddress = tostring(coalesce(column_ifexists("RecipientAddress", ""), column_ifexists("Recipient", ""), column_ifexists("To", ""), "Unknown"))
| extend Status = tostring(coalesce(column_ifexists("Status", ""), column_ifexists("DeliveryStatus", ""), column_ifexists("EventType", ""), column_ifexists("Action", ""), column_ifexists("Result", ""), ""))
| extend Subject = tostring(coalesce(column_ifexists("Subject", ""), ""))
| extend FromIP = tostring(coalesce(column_ifexists("FromIP", ""), column_ifexists("SenderIP", ""), column_ifexists("ClientIP", ""), ""))
| extend __text = strcat(Status, " ", Subject, " ", SenderAddress, " ", RecipientAddress)
| extend __isCritical = (__text matches regex @"(?i)\b(fail(ed|ure)?|blocked|quarantined|reject(ed)?|undeliver(ed|able)?|error|timeout|bounced)\b")
| extend __isBusinessCritical = (__text matches regex @"(?i)\b(Power\s*BI|PBI|SkyGuard)\b" and __isCritical)
| where __isCritical or __isBusinessCritical
| summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by SenderAddress, RecipientAddress, Status, Subject, FromIP
| extend __RecordKind="AggregatedMessageTraceRisk"
| project TimeGenerated, FirstTime, LastTime, EventCount, SenderAddress, RecipientAddress, Status, Subject, FromIP, __RecordKind
"@
}

function New-AzureAdUsersOptimizedQuery {
    param(
        [string]$TableName,
        [string]$StartUtc,
        [string]$EndUtc
    )

    return @"
$TableName
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
| extend __UserPrincipalName = tostring(coalesce(column_ifexists("userPrincipalName", ""), column_ifexists("UserPrincipalName", ""), column_ifexists("UserUPN", ""), column_ifexists("UPN", ""), column_ifexists("mail", ""), column_ifexists("Mail", ""), column_ifexists("displayName", ""), column_ifexists("DisplayName", ""), column_ifexists("Id", ""), column_ifexists("ObjectId", ""), "Unknown"))
| summarize arg_max(TimeGenerated, *) by __UserPrincipalName
| extend userPrincipalName = __UserPrincipalName
| extend mail = tostring(coalesce(column_ifexists("mail", ""), column_ifexists("Mail", ""), column_ifexists("EmailAddress", ""), ""))
| extend displayName = tostring(coalesce(column_ifexists("displayName", ""), column_ifexists("DisplayName", ""), column_ifexists("UserDisplayName", ""), ""))
| extend accountEnabled = tostring(coalesce(column_ifexists("accountEnabled", ""), column_ifexists("AccountEnabled", ""), ""))
| extend department = tostring(coalesce(column_ifexists("department", ""), column_ifexists("Department", ""), ""))
| extend disabledDateTime = tostring(coalesce(column_ifexists("disabledDateTime", ""), column_ifexists("DisabledDateTime", ""), column_ifexists("accountDisabledDateTime", ""), column_ifexists("AccountDisabledDateTime", ""), ""))
| extend FirstTime=TimeGenerated, LastTime=TimeGenerated, EventCount=1, __RecordKind="LatestUserSnapshot"
| project TimeGenerated, FirstTime, LastTime, EventCount, userPrincipalName, mail, displayName, accountEnabled, department, disabledDateTime, __RecordKind
"@
}

function New-DcrLogErrorsOptimizedQuery {
    param(
        [string]$StartUtc,
        [string]$EndUtc
    )

    return @"
DCRLogErrors
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
| summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by InputStreamId, OperationName, Message
| extend Status="Failed", __RecordKind="AggregatedDcrLogError"
| project TimeGenerated, FirstTime, LastTime, EventCount, InputStreamId, OperationName, Message, Status, __RecordKind
"@
}

function New-IntuneAuditLogsOptimizedQuery {
    param(
        [string]$StartUtc,
        [string]$EndUtc
    )

    return @"
IntuneAuditLogsDCR_CL
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
| extend ActorDisplayName = tostring(coalesce(column_ifexists("InitiatorDisplayName", ""), column_ifexists("InitiatorDisplayName_s", ""), column_ifexists("ActorDisplayName", ""), column_ifexists("ActorDisplayName_s", ""), column_ifexists("DisplayName", ""), column_ifexists("DisplayName_s", ""), column_ifexists("InitiatedByUserDisplayName", ""), column_ifexists("InitiatedByUserDisplayName_s", ""), column_ifexists("UserDisplayName", ""), column_ifexists("UserDisplayName_s", ""), ""))
| extend ActorUserPrincipalName = tostring(coalesce(column_ifexists("InitiatorUserPrincipalName", ""), column_ifexists("InitiatorUserPrincipalName_s", ""), column_ifexists("ActorInitiator", ""), column_ifexists("ActorUPN", ""), column_ifexists("ActorUPN_s", ""), column_ifexists("ActorUserPrincipalName", ""), column_ifexists("ActorUserPrincipalName_s", ""), column_ifexists("InitiatedByUserPrincipalName", ""), column_ifexists("InitiatedByUserPrincipalName_s", ""), column_ifexists("UserPrincipalName", ""), column_ifexists("UserPrincipalName_s", ""), column_ifexists("UPN", ""), column_ifexists("UPN_s", ""), column_ifexists("Actor", ""), column_ifexists("Actor_s", ""), column_ifexists("UserId", ""), column_ifexists("UserId_s", ""), column_ifexists("Identity", ""), column_ifexists("Identity_s", ""), ""))
| extend Actor = case(isnotempty(ActorDisplayName) and isnotempty(ActorUserPrincipalName) and ActorDisplayName != ActorUserPrincipalName, strcat(ActorDisplayName, " / ", ActorUserPrincipalName), isnotempty(ActorDisplayName), ActorDisplayName, ActorUserPrincipalName)
| extend OperationName = tostring(coalesce(column_ifexists("OperationName", ""), column_ifexists("OperationName_s", ""), column_ifexists("ActivityDisplayName", ""), column_ifexists("ActivityDisplayName_s", ""), column_ifexists("Activity", ""), column_ifexists("Activity_s", ""), column_ifexists("Operation", ""), column_ifexists("Operation_s", ""), column_ifexists("Action", ""), column_ifexists("Action_s", ""), column_ifexists("AuditEventType", ""), column_ifexists("AuditEventType_s", ""), "Intune Audit Event"))
| extend TargetDeviceName = tostring(coalesce(column_ifexists("TargetDeviceName", ""), column_ifexists("TargetDeviceName_s", ""), column_ifexists("DeviceName", ""), column_ifexists("DeviceName_s", ""), column_ifexists("ManagedDeviceName", ""), column_ifexists("ManagedDeviceName_s", ""), ""))
| extend Result = tostring(coalesce(column_ifexists("Result", ""), column_ifexists("Result_s", ""), column_ifexists("ResultStatus", ""), column_ifexists("ResultStatus_s", ""), column_ifexists("Status", ""), column_ifexists("Status_s", ""), column_ifexists("ActivityResult", ""), column_ifexists("ActivityResult_s", ""), column_ifexists("OperationStatus", ""), column_ifexists("OperationStatus_s", ""), ""))
| extend ResultDescription = tostring(coalesce(column_ifexists("ResultDescription", ""), column_ifexists("ResultDescription_s", ""), column_ifexists("FailureReason", ""), column_ifexists("FailureReason_s", ""), column_ifexists("Message", ""), column_ifexists("Message_s", ""), column_ifexists("ErrorMessage", ""), column_ifexists("ErrorMessage_s", ""), ""))
| extend __status = tolower(Result)
| extend __isFailed = (__status in ("false","fail","failed","failure","denied","error","timeout","1") or (__status matches regex @"^\d+$" and toint(__status) != 0) or (tolower(ResultDescription) matches regex @"\b(fail|failed|failure|denied|error|timeout)\b"))
| extend __isDeleteDisable = tolower(OperationName) matches regex @"(^|[^a-z])(delete|deleted|remove|removed|disable|disabled|deactivate|deactivated)([^a-z]|$)"
| extend __RecordKind = case(__isDeleteDisable, "AggregatedDeleteDisable", __isFailed, "AggregatedIntuneAuditRisk", "AggregatedIntuneAuditRecord")
| summarize TimeGenerated=max(TimeGenerated), FirstTime=min(TimeGenerated), LastTime=max(TimeGenerated), EventCount=count() by Actor, ActorDisplayName, ActorUserPrincipalName, OperationName, TargetDeviceName, Result, ResultDescription, __RecordKind
| project TimeGenerated, FirstTime, LastTime, EventCount, Actor, ActorDisplayName, ActorUserPrincipalName, OperationName, TargetDeviceName, Result, ResultDescription, __RecordKind
"@
}

function New-MailboxStatisticsOptimizedQuery {
    param(
        [string]$StartUtc,
        [string]$EndUtc
    )

    # 返回所有邮箱统计数据，不在查询端过滤，由报告端进行风险分析
    return @"
MailboxStatisticsDCR_CL
| where TimeGenerated >= datetime(2026-06-07T06:00:00.0000000Z) and TimeGenerated < datetime(2026-06-17T06:00:00.0000000Z)
| extend RecipientTypeDetails = tostring(column_ifexists("RecipientTypeDetails", ""))
| extend AvailableSpaceGB = todouble(column_ifexists("AvailableSpaceGB", real(null)))
| extend QuotaLimitGB = todouble(column_ifexists("QuotaLimitGB", real(null)))
| where RecipientTypeDetails contains "SharedMailbox"
    or (isnotnull(QuotaLimitGB) and QuotaLimitGB > 0 and isnotnull(AvailableSpaceGB) and AvailableSpaceGB < QuotaLimitGB * 0.05)
| summarize arg_max(TimeGenerated, *) by RecipientTypeDetails
| project TimeGenerated, 
    RecipientTypeDetails, 
    AvailableSpaceGB, 
    QuotaLimitGB, 
    UsagePercent = iff(QuotaLimitGB > 0 and isnotnull(AvailableSpaceGB), round((1 - AvailableSpaceGB / QuotaLimitGB) * 100, 2), real(null))
| order by TimeGenerated desc
"@
}

function New-RiskOptimizedLogTableQuery {
    param(
        [string]$TableName,
        [string]$StartUtc,
        [string]$EndUtc
    )

    switch ($TableName) {
        'AADManagedIdentitySignInLogs' { return New-AadIdentitySigninOptimizedQuery -TableName $TableName -StartUtc $StartUtc -EndUtc $EndUtc }
        'AADServicePrincipalSignInLogs' { return New-AadIdentitySigninOptimizedQuery -TableName $TableName -StartUtc $StartUtc -EndUtc $EndUtc }
        'AssignedLicensesDCR_CL' { return New-AssignedLicensesOptimizedQuery -TableName $TableName -StartUtc $StartUtc -EndUtc $EndUtc }
        'AuditLogs' { return New-AuditLogsOptimizedQuery -StartUtc $StartUtc -EndUtc $EndUtc }
        'DCRLogErrors' { return New-DcrLogErrorsOptimizedQuery -StartUtc $StartUtc -EndUtc $EndUtc }
        'IntuneAuditLogsDCR_CL' { return New-IntuneAuditLogsOptimizedQuery -StartUtc $StartUtc -EndUtc $EndUtc }
        'MailboxStatisticsDCR_CL' { return New-MailboxStatisticsOptimizedQuery -StartUtc $StartUtc -EndUtc $EndUtc }
        'MessageTraceDataDCR_CL' { return New-MessageTraceOptimizedQuery -StartUtc $StartUtc -EndUtc $EndUtc }
        'SigninLogs' { return New-SigninLogsOptimizedQuery -StartUtc $StartUtc -EndUtc $EndUtc }
        default {
            $query = New-SafeLogTableQuery -TableName $TableName -StartUtc $StartUtc -EndUtc $EndUtc
            $riskFilter = Get-LogRiskFilterKql -TableName $TableName
            if (-not [string]::IsNullOrWhiteSpace($riskFilter)) {
                $query = @"
$TableName
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
$riskFilter
| sort by TimeGenerated desc
| take 20000
"@
            }
            return $query
        }
    }
}

function New-SafeLogTableQuery {
    param(
        [string]$TableName,
        [string]$StartUtc,
        [string]$EndUtc,
        [int]$MaxRows = 20000
    )

    return @"
$TableName
| where TimeGenerated >= datetime($StartUtc) and TimeGenerated < datetime($EndUtc)
| sort by TimeGenerated desc
| take $MaxRows
"@
}

function New-TableTotalCountQuery {
    <#
    .SYNOPSIS
        Generate a KQL query to get the total record count for a table within a time range
    .DESCRIPTION
        Generates a count query with explicit where clause for time range filtering.
        Example: TableName | where TimeGenerated >= datetime(...) and TimeGenerated < datetime(...) | count
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [datetime]$StartTime,

        [Parameter(Mandatory = $true)]
        [datetime]$EndTime
    )

    $start = $StartTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $end = $EndTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    return "$TableName | where TimeGenerated >= datetime($start) and TimeGenerated < datetime($end) | count"
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
    if ($RiskOnly) {
        return New-RiskOptimizedLogTableQuery -TableName $TableName -StartUtc $start -EndUtc $end
    }

    return New-SafeLogTableQuery -TableName $TableName -StartUtc $start -EndUtc $end
}
