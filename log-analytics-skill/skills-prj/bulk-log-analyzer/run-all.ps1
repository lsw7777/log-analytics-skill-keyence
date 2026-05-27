# ============================================================
# Log Analytics One-Click Script
# ============================================================

param(
    [Parameter(Mandatory = $false)]
    [string]$TableName = "",

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
    [string]$StartDate = "",

    [Parameter(Mandatory = $false)]
    [string]$EndDate = "",

    [Parameter(Mandatory = $false)]
    [string]$CustomStart = "",

    [Parameter(Mandatory = $false)]
    [string]$CustomEnd = "",

    [Parameter(Mandatory = $false)]
    [switch]$ClearCache,

    [Parameter(Mandatory = $false)]
    [switch]$ClearCashe
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$TempDir = "$env:USERPROFILE\AppData\Local\Temp\opencode"
$CacheDir = "$TempDir\cache"
$Now = Get-Date
. (Join-Path $ScriptDir 'log-analyzer-core.ps1')

if ($ClearCache -or $ClearCashe) {
    $removedCount = Clear-LogCache -CacheDir $CacheDir
    Write-Host "Cache cleared: $CacheDir ($removedCount items removed)" -ForegroundColor Green
    exit 0
}

if (-not $TableName) {
    $TableName = Select-LogTableInteractive
} else {
    $TableName = Resolve-LogTableSelection -Selection $TableName
}

# Determine analysis date and time range
if ($CustomStart -and $CustomEnd) {
    $StartTime = [DateTime]::Parse($CustomStart)
    $EndTime = [DateTime]::Parse($CustomEnd)
    $AnalysisDateStr = $StartTime.ToString("yyyyMMdd")
    $AnalysisDateDisplay = $StartTime.ToString("yyyy-MM-dd")
    Write-Host "Time range: Custom ($($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss')))" -ForegroundColor Cyan
}
elseif ($AnalysisDate) {
    $range = Get-LogTimeRangeFromDates -StartDate $AnalysisDate -EndDate $AnalysisDate
    $StartTime = $range.StartTime
    $EndTime = $range.EndTime
    $AnalysisDateStr = $range.AnalysisDateStr
    $AnalysisDateDisplay = $range.AnalysisDateDisplay
    Write-Host "Time range: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
}
elseif ($StartDate -or $EndDate) {
    if (-not $StartDate -or -not $EndDate) {
        throw '-StartDate and -EndDate must be provided together.'
    }
    $range = Get-LogTimeRangeFromDates -StartDate $StartDate -EndDate $EndDate
    $StartTime = $range.StartTime
    $EndTime = $range.EndTime
    $AnalysisDateStr = $range.AnalysisDateStr
    $AnalysisDateDisplay = $range.AnalysisDateDisplay
    Write-Host "Time range: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
}
else {
    $range = Select-LogTimeRangeInteractive -Now $Now
    $StartTime = $range.StartTime
    $EndTime = $range.EndTime
    $AnalysisDateStr = $range.AnalysisDateStr
    $AnalysisDateDisplay = $range.AnalysisDateDisplay
    Write-Host "Time range: $AnalysisDateDisplay ($($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss')))" -ForegroundColor Cyan
}

Write-Host "  StartTime ISO: $($StartTime.ToString('o'))" -ForegroundColor Cyan
Write-Host "  EndTime ISO: $($EndTime.ToString('o'))" -ForegroundColor Cyan

$paths = Get-LogArtifactPaths -TempDir $TempDir -TableName $TableName -AnalysisDateStr $AnalysisDateStr -Now $Now
$CsvFile = $paths.CsvFile
$HtmlFile = $paths.HtmlFile
$HtmlFilePath = $paths.HtmlFilePath
$CacheCsv = $paths.CacheCsv
$CacheMeta = $paths.CacheMeta

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
        [int]$CacheTTL,
        [string]$TableName,
        [datetime]$StartTime,
        [datetime]$EndTime
    )

    if (-not (Test-Path $CacheCsv)) {
        return @{ Hit = $false; Reason = "Cache file not found" }
    }

    if (-not (Test-Path $CacheMeta)) {
        return @{ Hit = $false; Reason = "Cache metadata not found" }
    }

    try {
        $meta = Get-Content $CacheMeta -Raw | ConvertFrom-Json
        if (-not (Test-LogCacheMetadataMatches -Meta $meta -TableName $TableName -StartTime $StartTime -EndTime $EndTime)) {
            return @{ Hit = $false; Reason = 'Cache metadata does not match table and time range' }
        }

        if (-not (Test-LogCachePayloadValid -CacheCsv $CacheCsv -RecordCount $meta.RecordCount)) {
            Remove-Item -Path $CacheCsv -Force -ErrorAction SilentlyContinue
            Remove-Item -Path $CacheMeta -Force -ErrorAction SilentlyContinue
            return @{ Hit = $false; Reason = 'Cache payload is empty or invalid' }
        }

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
        [int]$CacheTTL,
        [datetime]$StartTime,
        [datetime]$EndTime
    )
    $recordCount = Get-LogCsvRecordCount -CsvPath $SourceCsv
    if ($recordCount -le 0) {
        Remove-Item -Path $CacheCsv -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $CacheMeta -Force -ErrorAction SilentlyContinue
        Write-Host "Cache skipped: query returned 0 records" -ForegroundColor Yellow
        return
    }

    Copy-Item -Path $SourceCsv -Destination $CacheCsv -Force

    $meta = @{
        TableName = $TableName
        CacheTime = (Get-Date).ToString("o")
        CacheTTL = $CacheTTL
        RecordCount = $recordCount
        Hours = $Hours
        StartTimeUtc = Get-LogCacheTimeKey -Time $StartTime
        EndTimeUtc = Get-LogCacheTimeKey -Time $EndTime
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
    $cacheResult = Test-Cache -CacheCsv $CacheCsv -CacheMeta $CacheMeta -CacheTTL $CacheTTL -TableName $TableName -StartTime $StartTime -EndTime $EndTime
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
        Save-Cache -SourceCsv $CsvFile -CacheCsv $CacheCsv -CacheMeta $CacheMeta -TableName $TableName -CacheTTL $CacheTTL -StartTime $StartTime -EndTime $EndTime
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
}

$allUsers = @()
foreach ($row in $data) {
    $u = Get-UserValue -Row $row -TableName $TableName
    $allUsers += $u
}
$uniqueUsers = ($allUsers | Select-Object -Unique).Count

$allOps = @($data | ForEach-Object { Get-OperationValue -Row $_ -TableName $TableName })
$uniqueOps = ($allOps | Select-Object -Unique).Count

$workloadMap = @{}
foreach ($row in $data) {
    $wl = Get-WorkloadValue -Row $row -TableName $TableName
    $workloadMap[$wl] = ($workloadMap[$wl] + 1)
}

$successCount = 0
$failCount = 0
$unknownCount = 0
foreach ($row in $data) {
    $s = Get-SuccessValue -Row $row -TableName $TableName
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

& "$ScriptDir\analyze.ps1" -CsvPath $CsvFile -OutputPath $HtmlFilePath -AnalysisDate $AnalysisDateDisplay -TableName $TableName

if (-not (Test-Path $HtmlFilePath)) {
    Write-Host "Error: HTML file not generated" -ForegroundColor Red
    exit 1
}

Write-Host "HTML report generated: $HtmlFilePath" -ForegroundColor Green
Write-Host ""

# Step 4: Open in browser
Write-Host "[4/4] Opening browser..." -ForegroundColor Yellow

$htmlUrl = "file:///$($HtmlFilePath -replace '\\', '/')"
Write-Host "HTML URL: $htmlUrl" -ForegroundColor Cyan

Start-Process $HtmlFilePath

Write-Host ""
Write-Host "============================================" -ForegroundColor Magenta
Write-Host "  Done!" -ForegroundColor Magenta
Write-Host "============================================" -ForegroundColor Magenta
Write-Host "CSV: $CsvFile" -ForegroundColor Cyan
Write-Host "HTML: $HtmlFilePath" -ForegroundColor Cyan
Write-Host "URL: $htmlUrl" -ForegroundColor Cyan
