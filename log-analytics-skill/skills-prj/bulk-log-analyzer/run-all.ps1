# ============================================================
# Log Analytics One-Click Script
# ============================================================

param(
    [Parameter(Mandatory = $false)]
    [string]$TableName = "AuditGeneralDCR_CL",

    [Parameter(Mandatory = $false)]
    [int]$Hours = 24,

    [Parameter(Mandatory = $false)]
    [switch]$ForceLogin,

    [Parameter(Mandatory = $false)]
    [switch]$UseCache = $true,

    [Parameter(Mandatory = $false)]
    [switch]$ForceRefresh,

    [Parameter(Mandatory = $false)]
    [int]$CacheTTL = 24,

    [Parameter(Mandatory = $false)]
    [string]$AnalysisDate = "",

    [Parameter(Mandatory = $false)]
    [switch]$UseYesterday = $true,

    [Parameter(Mandatory = $false)]
    [ValidateSet("yesterday", "today", "last7days", "last30days", "custom")]
    [string]$DateRange = "yesterday",

    [Parameter(Mandatory = $false)]
    [int]$DateOffset = -1,

    [Parameter(Mandatory = $false)]
    [string]$CustomStart = "",

    [Parameter(Mandatory = $false)]
    [string]$CustomEnd = ""
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$TempDir = "$env:USERPROFILE\AppData\Local\Temp\opencode"
$CacheDir = "$TempDir\cache"
$DateStr = Get-Date -Format "yyyyMMdd_HHmm"

# Derive short name from table (e.g., AuditGeneralDCR_CL -> General)
$ShortName = if ($TableName -match '^(.+)DCR_CL$') {
    $matches[1] -replace 'AuditGeneral', 'General' -replace 'SharePointAudit', 'SPAudit' -replace 'MessageTraceData', 'MsgTrace' -replace 'AssignedLicenses', 'Licenses' -replace 'AzureADUsers', 'AADUsers' -replace 'MailboxStatistics', 'Mailbox'
} else {
    $TableName -replace '_CL$', ''
}

# Determine analysis date and time range
$now = Get-Date

# Priority: CustomStart/CustomEnd > DateOffset > DateRange > UseYesterday > Hours
if ($CustomStart -and $CustomEnd) {
    # Custom time range
    $StartTime = [DateTime]::Parse($CustomStart)
    $EndTime = [DateTime]::Parse($CustomEnd)
    $AnalysisDateStr = $StartTime.ToString("yyyyMMdd")
    Write-Host "Time range: Custom ($($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss')))" -ForegroundColor Cyan
}
elseif ($DateOffset -ne 0) {
    # DateOffset: 0=today, -1=yesterday, -2=day before yesterday, etc.
    $TargetDate = $now.Date.AddDays($DateOffset)
    $TargetEnd = $TargetDate.AddDays(1).AddSeconds(-1)
    $AnalysisDateStr = $TargetDate.ToString("yyyyMMdd")
    $StartTime = $TargetDate
    $EndTime = $TargetEnd
    if ($DateOffset -eq -1) {
        Write-Host "Time range: Yesterday ($($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss')))" -ForegroundColor Cyan
    }
    elseif ($DateOffset -eq 0) {
        Write-Host "Time range: Today ($($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss')))" -ForegroundColor Cyan
    }
    else {
        Write-Host "Time range: $([Math]::Abs($DateOffset)) day(s) ago ($($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss')))" -ForegroundColor Cyan
    }
}
elseif ($DateRange -ne "custom") {
    # Preset date ranges
    switch ($DateRange) {
        "yesterday" {
            $StartDate = $now.Date.AddDays(-1)
            $EndDate = $StartDate.AddDays(1).AddSeconds(-1)
            Write-Host "Time range: Yesterday" -ForegroundColor Cyan
        }
        "today" {
            $StartDate = $now.Date
            $EndDate = $now
            Write-Host "Time range: Today (so far)" -ForegroundColor Cyan
        }
        "last7days" {
            $StartDate = $now.Date.AddDays(-7)
            $EndDate = $now
            Write-Host "Time range: Last 7 days" -ForegroundColor Cyan
        }
        "last30days" {
            $StartDate = $now.Date.AddDays(-30)
            $EndDate = $now
            Write-Host "Time range: Last 30 days" -ForegroundColor Cyan
        }
    }
    $AnalysisDateStr = $StartDate.ToString("yyyyMMdd")
    $StartTime = $StartDate
    $EndTime = $EndDate
    Write-Host "  $($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
}
elseif ($UseYesterday) {
    # Yesterday 00:00:00 to 23:59:59 (adaptive, not hardcoded)
    $yesterday = $now.Date.AddDays(-1)
    $yesterdayEnd = $yesterday.AddDays(1).AddSeconds(-1)
    $AnalysisDateStr = $yesterday.ToString("yyyyMMdd")
    $StartTime = $yesterday
    $EndTime = $yesterdayEnd
    Write-Host "Time range: Yesterday ($($yesterday.ToString('yyyy-MM-dd HH:mm:ss')) to $($yesterdayEnd.ToString('yyyy-MM-dd HH:mm:ss')))" -ForegroundColor Cyan
}
else {
    # Last N hours from now
    if ($AnalysisDate) {
        $AnalysisDateStr = $AnalysisDate -replace '-', ''
    } else {
        $AnalysisDateStr = $now.ToString("yyyyMMdd")
    }
    $StartTime = $now.AddHours(-$Hours)
    $EndTime = $now
    Write-Host "Time range: Last $Hours hours ($($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss')))" -ForegroundColor Cyan
}

Write-Host "  StartTime ISO: $($StartTime.ToString('o'))" -ForegroundColor Cyan
Write-Host "  EndTime ISO: $($EndTime.ToString('o'))" -ForegroundColor Cyan

$CsvFile = "$TempDir\$($ShortName)_$AnalysisDateStr.csv"
$HtmlFile = "$TempDir\$($ShortName)_$AnalysisDateStr.html"
$CacheCsv = "$CacheDir\$($ShortName)_$AnalysisDateStr.csv"
$CacheMeta = "$CacheDir\$($ShortName)_$AnalysisDateStr.meta.json"

# Also create a copy with table name + timestamp format
$TimestampStr = Get-Date -Format "yyyyMMdd_HHmmss"
$HtmlFileAlt = "$TempDir\$($ShortName)_$TimestampStr.html"

if (-not (Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
}
if (-not (Test-Path $CacheDir)) {
    New-Item -ItemType Directory -Path $CacheDir -Force | Out-Null
}

Write-Host "============================================" -ForegroundColor Magenta
Write-Host "  Log Analytics One-Click" -ForegroundColor Magenta
Write-Host "============================================" -ForegroundColor Magenta
Write-Host "Table: $TableName" -ForegroundColor Cyan
Write-Host "Hours: $Hours" -ForegroundColor Cyan
Write-Host "Cache: $(if($UseCache){'Enabled'}else{'Disabled'})" -ForegroundColor Cyan
Write-Host "CSV: $CsvFile" -ForegroundColor Cyan
Write-Host "HTML: $HtmlFile" -ForegroundColor Cyan
Write-Host ""

# ============================================================
# Cache Check Function
# ============================================================
function Test-Cache {
    param(
        [string]$CacheCsv,
        [string]$CacheMeta,
        [int]$CacheTTL
    )

    if (-not (Test-Path $CacheCsv)) {
        return @{ Hit = $false; Reason = "Cache file not found" }
    }

    if (-not (Test-Path $CacheMeta)) {
        return @{ Hit = $false; Reason = "Cache metadata not found" }
    }

    try {
        $meta = Get-Content $CacheMeta -Raw | ConvertFrom-Json
        $cacheTime = [DateTime]::Parse($meta.CacheTime)
        $age = (Get-Date) - $cacheTime
        $ttlHours = if ($meta.CacheTTL) { $meta.CacheTTL } else { $CacheTTL }

        if ($age.TotalHours -gt $ttlHours) {
            return @{ Hit = $false; Reason = "Cache expired ($([Math]::Round($age.TotalHours, 1))h > ${ttlHours}h)" }
        }

        return @{
            Hit = $true
            Reason = "Cache hit (age: $([Math]::Round($age.TotalMinutes, 0))min, records: $($meta.RecordCount))"
            RecordCount = $meta.RecordCount
            CacheTime = $cacheTime
        }
    }
    catch {
        return @{ Hit = $false; Reason = "Cache metadata parse error: $_" }
    }
}

function Save-Cache {
    param(
        [string]$SourceCsv,
        [string]$CacheCsv,
        [string]$CacheMeta,
        [string]$TableName,
        [int]$CacheTTL
    )

    Copy-Item -Path $SourceCsv -Destination $CacheCsv -Force

    $recordCount = (Import-Csv -Path $SourceCsv -Encoding UTF8).Count
    $meta = @{
        TableName = $TableName
        CacheTime = (Get-Date).ToString("o")
        CacheTTL = $CacheTTL
        RecordCount = $recordCount
        Hours = $Hours
    }
    $meta | ConvertTo-Json | Out-File -FilePath $CacheMeta -Encoding UTF8 -Force

    Write-Host "Cache saved: $CacheCsv ($recordCount records)" -ForegroundColor Green
}

# ============================================================
# Step 1: Check Cache or Query Data
# ============================================================
Write-Host "[1/4] Checking data source..." -ForegroundColor Yellow

$cacheResult = $null
if ($UseCache -and -not $ForceRefresh) {
    $cacheResult = Test-Cache -CacheCsv $CacheCsv -CacheMeta $CacheMeta -CacheTTL $CacheTTL
}

if ($cacheResult -and $cacheResult.Hit) {
    Write-Host "  $($cacheResult.Reason)" -ForegroundColor Green
    Write-Host "  Using cached data, skipping Azure query!" -ForegroundColor Green

    # Copy cache to working file
    Copy-Item -Path $CacheCsv -Destination $CsvFile -Force
    Write-Host "  Working file: $CsvFile" -ForegroundColor Cyan
} else {
    if ($cacheResult) {
        Write-Host "  $($cacheResult.Reason)" -ForegroundColor Yellow
    }
    Write-Host "  Querying Azure Log Analytics..." -ForegroundColor Yellow

    $QueryParams = @{
        TableName = $TableName
        Hours     = $Hours
        ExportCsv = $CsvFile
        StartTime = $StartTime.ToString("o")
        EndTime   = $EndTime.ToString("o")
    }

    if ($ForceLogin) {
        $QueryParams['ForceLogin'] = $true
    }

    & "$ScriptDir\azure_log_query.ps1" @QueryParams

    if (-not (Test-Path $CsvFile)) {
        Write-Host "Error: CSV file not generated" -ForegroundColor Red
        exit 1
    }

    # Save to cache
    if ($UseCache) {
        Save-Cache -SourceCsv $CsvFile -CacheCsv $CacheCsv -CacheMeta $CacheMeta -TableName $TableName -CacheTTL $CacheTTL
    }

    Write-Host "Data query complete!" -ForegroundColor Green
}
Write-Host ""

# Step 2: Load and compute stats
Write-Host "[2/4] Computing statistics..." -ForegroundColor Yellow

$data = Import-Csv -Path $CsvFile -Encoding UTF8
$totalEvents = $data.Count
Write-Host "Total records: $totalEvents" -ForegroundColor Green

if ($totalEvents -eq 0) {
    Write-Host "Warning: No data found" -ForegroundColor Yellow
    exit 0
}

$allUsers = @()
foreach ($row in $data) {
    $u = if ($row.UserUPN -and $row.UserUPN -ne '') { $row.UserUPN } elseif ($row.UserId -and $row.UserId -ne '') { $row.UserId } else { 'Unknown' }
    $allUsers += $u
}
$uniqueUsers = ($allUsers | Select-Object -Unique).Count

$allOps = @($data | ForEach-Object { $_.Operation })
$uniqueOps = ($allOps | Select-Object -Unique).Count

$workloadMap = @{}
foreach ($row in $data) {
    $wl = if ($row.Workload) { $row.Workload } else { 'Unknown' }
    $workloadMap[$wl] = ($workloadMap[$wl] + 1)
}

$successCount = 0
$failCount = 0
$unknownCount = 0
foreach ($row in $data) {
    $s = $row.IsSuccess
    if ($s -eq 'true') { $successCount++ }
    elseif ($s -eq 'false') { $failCount++ }
    else { $unknownCount++ }
}

Write-Host "Unique users: $uniqueUsers" -ForegroundColor Green
Write-Host "Unique ops: $uniqueOps" -ForegroundColor Green
Write-Host "Workloads: $($workloadMap.Count)" -ForegroundColor Green
Write-Host "Success: $successCount | Failed: $failCount | Unknown: $unknownCount" -ForegroundColor Green
Write-Host ""

# Step 3: Generate HTML
Write-Host "[3/4] Generating HTML report..." -ForegroundColor Yellow

& "$ScriptDir\analyze.ps1" -CsvPath $CsvFile -OutputPath $HtmlFile -AnalysisDate (Get-Date -Format "yyyy-MM-dd")

if (-not (Test-Path $HtmlFile)) {
    Write-Host "Error: HTML file not generated" -ForegroundColor Red
    exit 1
}

# Copy to alternate filename format (TableName_Timestamp.html)
Copy-Item -Path $HtmlFile -Destination $HtmlFileAlt -Force
Write-Host "HTML report generated: $HtmlFile" -ForegroundColor Green
Write-Host "HTML report (alternate): $HtmlFileAlt" -ForegroundColor Green
Write-Host ""

# Step 4: Open in browser
Write-Host "[4/4] Opening browser..." -ForegroundColor Yellow

$htmlUrl = "file:///$($HtmlFile -replace '\\', '/')"
Write-Host "HTML URL: $htmlUrl" -ForegroundColor Cyan

Start-Process $HtmlFile

Write-Host ""
Write-Host "============================================" -ForegroundColor Magenta
Write-Host "  Done!" -ForegroundColor Magenta
Write-Host "============================================" -ForegroundColor Magenta
Write-Host "CSV: $CsvFile" -ForegroundColor Cyan
Write-Host "HTML: $HtmlFile" -ForegroundColor Cyan
Write-Host "URL: $htmlUrl" -ForegroundColor Cyan