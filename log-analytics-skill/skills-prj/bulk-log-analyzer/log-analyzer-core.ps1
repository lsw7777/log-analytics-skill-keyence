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

    $timestamp = $Now.ToString('yyyyMMdd_HHmm')

    return [PSCustomObject]@{
        CsvFile = Join-Path $TempDir "$($TableName)_$AnalysisDateStr.csv"
        HtmlFile = Join-Path $TempDir "$($TableName)_$timestamp.html"
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
