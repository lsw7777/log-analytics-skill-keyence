# ============================================================
# Log Analytics merged risk report
# ============================================================

param(
    [Parameter(Mandatory = $false)]
    [string]$TableName = "",

    [Parameter(Mandatory = $false)]
    [int]$Hours = 168,

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
    [switch]$ClearCashe,

    [Parameter(Mandatory = $false)]
    [switch]$NoOpen,

    [Parameter(Mandatory = $false)]
    [switch]$NoIsolatedQueryProcess,

    [Parameter(Mandatory = $false)]
    [switch]$NoRiskFilter
)

$ErrorActionPreference = 'Stop'
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

if (-not (Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
}
if (-not (Test-Path $CacheDir)) {
    New-Item -ItemType Directory -Path $CacheDir -Force | Out-Null
}

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
        return @{ Hit = $false; Reason = 'Cache file not found' }
    }
    if (-not (Test-Path $CacheMeta)) {
        return @{ Hit = $false; Reason = 'Cache metadata not found' }
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

        return @{ Hit = $true; Reason = "Cache hit (records: $($meta.RecordCount))"; RecordCount = $meta.RecordCount }
    } catch {
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
        Write-Host "  Cache skipped: query returned 0 records" -ForegroundColor Yellow
        return
    }

    Copy-Item -Path $SourceCsv -Destination $CacheCsv -Force
    $meta = @{
        TableName = $TableName
        CacheTime = (Get-Date).ToString('o')
        CacheTTL = $CacheTTL
        RecordCount = $recordCount
        Hours = $Hours
        StartTimeUtc = Get-LogCacheTimeKey -Time $StartTime
        EndTimeUtc = Get-LogCacheTimeKey -Time $EndTime
    }
    $meta | ConvertTo-Json | Out-File -FilePath $CacheMeta -Encoding UTF8 -Force
    Write-Host "  Cache saved: $CacheCsv ($recordCount records)" -ForegroundColor Green
}

if ($CustomStart -and $CustomEnd) {
    $StartTime = [DateTime]::Parse($CustomStart)
    $EndTime = [DateTime]::Parse($CustomEnd)
    Assert-LogTimeRangeWithinLimit -StartTime $StartTime -EndTime $EndTime
    $AnalysisDateStr = "$($StartTime.ToString('yyyyMMddHHmm'))_$($EndTime.ToString('yyyyMMddHHmm'))"
    $AnalysisDateDisplay = "$($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))"
}
elseif ($AnalysisDate) {
    $range = Get-LogTimeRangeFromDates -StartDate $AnalysisDate -EndDate $AnalysisDate
    $StartTime = $range.StartTime
    $EndTime = $range.EndTime
    $AnalysisDateStr = $range.AnalysisDateStr
    $AnalysisDateDisplay = $range.AnalysisDateDisplay
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
}
else {
    $range = Select-LogTimeRangeInteractive -Now $Now
    $StartTime = $range.StartTime
    $EndTime = $range.EndTime
    $AnalysisDateStr = $range.AnalysisDateStr
    $AnalysisDateDisplay = $range.AnalysisDateDisplay
}

if ([string]::IsNullOrWhiteSpace($TableName)) {
    $targetTables = @(Get-SupportedLogTables | ForEach-Object { $_.Name })
} else {
    $targetTables = @($TableName.Split(',') | ForEach-Object { Resolve-LogTableSelection -Selection $_.Trim() })
}

if ($targetTables.Count -eq 1) {
    $reportPaths = Get-LogArtifactPaths -TempDir $TempDir -TableName $targetTables[0] -AnalysisDateStr $AnalysisDateStr -Now $Now
} else {
    $reportPaths = Get-MergedLogArtifactPaths -TempDir $TempDir -AnalysisDateStr $AnalysisDateStr -Now $Now
}
$HtmlFilePath = $reportPaths.HtmlFilePath

Write-Host '============================================' -ForegroundColor Magenta
Write-Host '  Log Analytics Merged Risk Report' -ForegroundColor Magenta
Write-Host '============================================' -ForegroundColor Magenta
Write-Host "Tables: $($targetTables -join ', ')" -ForegroundColor Cyan
Write-Host "Time range: $AnalysisDateDisplay ($($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss')))" -ForegroundColor Cyan
Write-Host "Cache: $(if($UseCache){'Enabled'}else{'Disabled'})" -ForegroundColor Cyan
Write-Host "Risk prefilter: $(if($NoRiskFilter){'Disabled'}else{'Enabled'})" -ForegroundColor Cyan
Write-Host "HTML: $($reportPaths.HtmlFile)" -ForegroundColor Cyan
Write-Host ''

$csvFiles = [System.Collections.Generic.List[string]]::new()
$csvTables = [System.Collections.Generic.List[string]]::new()

foreach ($table in $targetTables) {
    $paths = Get-LogArtifactPaths -TempDir $TempDir -TableName $table -AnalysisDateStr $AnalysisDateStr -Now $Now
    Write-Host "[1/3] Data source: $table" -ForegroundColor Yellow

    $cacheResult = $null
    if ($UseCache -and -not $ForceRefresh) {
        $cacheResult = Test-Cache -CacheCsv $paths.CacheCsv -CacheMeta $paths.CacheMeta -CacheTTL $CacheTTL -TableName $table -StartTime $StartTime -EndTime $EndTime
    }

    if ($cacheResult -and $cacheResult.Hit) {
        Write-Host "  $($cacheResult.Reason)" -ForegroundColor Green
        Copy-Item -Path $paths.CacheCsv -Destination $paths.CsvFile -Force
    } else {
        if ($cacheResult) {
            Write-Host "  $($cacheResult.Reason)" -ForegroundColor Yellow
        }
        Remove-Item -Path $paths.CsvFile -Force -ErrorAction SilentlyContinue

        if ($NoIsolatedQueryProcess) {
            $queryParams = @{
                TableName = $table
                Hours = $Hours
                ExportCsv = $paths.CsvFile
                StartTime = $StartTime.ToString('o')
                EndTime = $EndTime.ToString('o')
            }
            if ($ForceLogin) { $queryParams['ForceLogin'] = $true }
            if (-not $NoRiskFilter) { $queryParams['RiskOnly'] = $true }

            & "$ScriptDir\azure_log_query.ps1" @queryParams
        } else {
            $queryArgs = @(
                '-NoProfile',
                '-ExecutionPolicy', 'Bypass',
                '-File', (Join-Path $ScriptDir 'azure_log_query.ps1'),
                '-TableName', $table,
                '-Hours', ([string]$Hours),
                '-ExportCsv', $paths.CsvFile,
                '-StartTime', $StartTime.ToString('o'),
                '-EndTime', $EndTime.ToString('o')
            )
            if ($ForceLogin) { $queryArgs += '-ForceLogin' }
            if (-not $NoRiskFilter) { $queryArgs += '-RiskOnly' }
            & powershell.exe @queryArgs
            if ($LASTEXITCODE -ne 0) {
                if ($LASTEXITCODE -eq 20) {
                    Write-Host "  Azure PowerShell module conflict detected. Stop querying remaining tables; repair Az modules first." -ForegroundColor Red
                    break
                }
                Write-Host "  Query failed for $table (exit code $LASTEXITCODE); skipping this table." -ForegroundColor Yellow
                continue
            }
        }

        if (-not (Test-Path $paths.CsvFile)) {
            Write-Host "  Query did not generate CSV for $table; skipping." -ForegroundColor Yellow
            continue
        }
        if ($UseCache) {
            Save-Cache -SourceCsv $paths.CsvFile -CacheCsv $paths.CacheCsv -CacheMeta $paths.CacheMeta -TableName $table -CacheTTL $CacheTTL -StartTime $StartTime -EndTime $EndTime
        }
    }

    $csvFiles.Add($paths.CsvFile) | Out-Null
    $csvTables.Add($table) | Out-Null
}

if ($csvFiles.Count -eq 0) {
    throw 'No CSV files were available for report generation.'
}

Write-Host ''
Write-Host '[2/3] Generating merged HTML report...' -ForegroundColor Yellow
& "$ScriptDir\analyze.ps1" -CsvPath $csvFiles.ToArray() -OutputPath $HtmlFilePath -AnalysisDate $AnalysisDateDisplay -TableName $csvTables.ToArray()

if (-not (Test-Path $HtmlFilePath)) {
    throw 'HTML file was not generated.'
}

$htmlUrl = "file:///$($HtmlFilePath -replace '\\', '/')"
Write-Host ''
Write-Host '[3/3] Done' -ForegroundColor Yellow
Write-Host "HTML: $HtmlFilePath" -ForegroundColor Cyan
Write-Host "URL: $htmlUrl" -ForegroundColor Cyan

if (-not $NoOpen) {
    Start-Process $HtmlFilePath
}
