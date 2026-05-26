$SupportedLogTables = @(
    [PSCustomObject]@{ Name = 'AssignedLicensesDCR_CL'; Description = 'Assigned Licenses' },
    [PSCustomObject]@{ Name = 'AuditGeneralDCR_CL'; Description = 'Audit General' },
    [PSCustomObject]@{ Name = 'AzureADUsersDCR_CL'; Description = 'Azure AD Users' },
    [PSCustomObject]@{ Name = 'MailboxStatisticsDCR_CL'; Description = 'Mailbox Statistics' },
    [PSCustomObject]@{ Name = 'MessageTraceDataDCR_CL'; Description = 'Message Trace Data' },
    [PSCustomObject]@{ Name = 'SharePointAuditDCR_CL'; Description = 'SharePoint Audit' },
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

    $startTime = $Now.Date.AddDays(-1)
    $endTime = $Now.Date

    return [PSCustomObject]@{
        StartTime = $startTime
        EndTime = $endTime
        AnalysisDateStr = $startTime.ToString('yyyyMMdd')
        AnalysisDateDisplay = $startTime.ToString('yyyy-MM-dd')
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
    Write-Host '  1. Yesterday (default)' -ForegroundColor Cyan
    Write-Host '  2. Single day' -ForegroundColor Cyan
    Write-Host '  3. Start date and end date' -ForegroundColor Cyan
    $selection = Read-Host 'Enter number or press Enter for 1'

    if ([string]::IsNullOrWhiteSpace($selection) -or $selection.Trim() -eq '1') {
        return Get-DefaultLogTimeRange -Now $Now
    }

    if ($selection.Trim() -eq '2') {
        $date = Read-Host 'Enter date (yyyy-MM-dd)'
        return Get-LogTimeRangeFromDates -StartDate $date -EndDate $date
    }

    if ($selection.Trim() -eq '3') {
        $startDate = Read-Host 'Enter start date (yyyy-MM-dd)'
        $endDate = Read-Host 'Enter end date (yyyy-MM-dd)'
        return Get-LogTimeRangeFromDates -StartDate $startDate -EndDate $endDate
    }

    throw "Invalid time range selection: $selection"
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
        'AssignedLicensesDCR_CL' {
            $common.UserFields = @('UserPrincipalName', 'UserUPN', 'UPN', 'Mail', 'EmailAddress', 'DisplayName', 'UserId', 'ObjectId', 'Id') + $common.UserFields
            $common.OperationFields = @('SkuPartNumber', 'LicenseName', 'AssignedLicenses', 'Licenses', 'AccountEnabled', 'UserType') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'ServicePlanName', 'SkuPartNumber', 'SourceSystem') + $common.WorkloadFields
            $common.DefaultOperation = 'Assigned License Record'
            $common.DefaultWorkload = 'Microsoft Entra ID Licensing'
            $common.DefaultSuccess = 'true'
        }
        'AuditGeneralDCR_CL' {
            $common.UserFields = @('UserUPN', 'UserId', 'Actor', 'ActorId', 'ActorUserPrincipalName') + $common.UserFields
            $common.OperationFields = @('Activity', 'Operation') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'RecordType') + $common.WorkloadFields
            $common.GroupFields = @('Activity', 'Operation', 'Workload')
            $common.UseCompositeOperationGroup = $true
        }
        'AzureADUsersDCR_CL' {
            $common.UserFields = @('UserPrincipalName', 'UserUPN', 'UPN', 'Mail', 'EmailAddress', 'DisplayName', 'Id', 'ObjectId') + $common.UserFields
            $common.OperationFields = @('AccountEnabled', 'UserType', 'JobTitle', 'Department', 'CreatedDateTime') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'SourceSystem', 'UserType', 'Department') + $common.WorkloadFields
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
        'SharePointAuditDCR_CL' {
            $common.UserFields = @('UserUPN', 'UserId', 'UserPrincipalName', 'Actor', 'ActorId') + $common.UserFields
            $common.OperationFields = @('Activity', 'Operation') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'SiteUrl', 'SourceSystem') + $common.WorkloadFields
            $common.GroupFields = @('Activity', 'Operation', 'Workload')
            $common.UseCompositeOperationGroup = $true
        }
        'WQCLogDCR_CL' {
            $common.UserFields = @('UserPrincipalName', 'UserUPN', 'UserId', 'User', 'Requester', 'Submitter', 'Owner') + $common.UserFields
            $common.OperationFields = @('Result', 'Action', 'Operation', 'Activity', 'QueryType', 'Command', 'Status') + $common.OperationFields
            $common.WorkloadFields = @('Workload', 'Service', 'Category', 'SourceSystem') + $common.WorkloadFields
            $common.DefaultOperation = 'WQC Log Event'
            $common.DefaultWorkload = 'WQC'
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
    if ($value -match '^(true|success|succeeded|delivered|expanded|completed|complete|ok|pass|passed|0)$') { return 'true' }
    if ($value -match '^(false|fail|failed|failure|undelivered|blocked|rejected|denied|error|timeout|quarantined|1)$') { return 'false' }
    return 'unknown'
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

function New-LogTableQuery {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TableName,

        [Parameter(Mandatory = $true)]
        [datetime]$StartTime,

        [Parameter(Mandatory = $true)]
        [datetime]$EndTime
    )

    $validTableNames = @($SupportedLogTables | ForEach-Object { $_.Name })
    if ($validTableNames -notcontains $TableName) {
        throw "Unsupported log table: $TableName"
    }

    $start = $StartTime.ToString('yyyy-MM-ddTHH:mm:ss')
    $end = $EndTime.ToString('yyyy-MM-ddTHH:mm:ss')
    return "$TableName | where TimeGenerated >= datetime($start) and TimeGenerated < datetime($end) | sort by TimeGenerated desc"
}
