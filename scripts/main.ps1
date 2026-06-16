# ============================================================
# Log Analytics merged risk report
# ============================================================

param(
    [Parameter(Mandatory = $false)]
    [string]$Prompt = "",

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
    [switch]$NoRiskFilter,

    [Parameter(Mandatory = $false)]
    [switch]$SkipTotalCount,

    [Parameter(Mandatory = $false)]
    [int]$TotalCountTimeoutSec = 30

    ,
    [Parameter(Mandatory = $false)]
    [switch]$VerifyLicenseGraph,

    [Parameter(Mandatory = $false)]
    [switch]$SkipLicenseGraph
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$TempDir = "$env:USERPROFILE\AppData\Local\Temp\opencode"
$CacheDir = "$TempDir\cache"
$Now = Get-Date
. (Join-Path $ScriptDir 'log-analyzer-shared.ps1')

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

function Format-LogByteSize {
    param([long]$Bytes)

    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N2} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N2} KB' -f ($Bytes / 1KB)) }
    return "$Bytes B"
}

function Format-LogElapsed {
    param([TimeSpan]$Elapsed)
    if ($Elapsed.TotalMinutes -ge 1) {
        return ('{0:N1} min' -f $Elapsed.TotalMinutes)
    }
    return ('{0:N1} sec' -f $Elapsed.TotalSeconds)
}

function Resolve-SkillPromptTimeRange {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [datetime]$Now = (Get-Date)
    )

    $text = $Prompt.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw 'Prompt cannot be empty.'
    }

    $roundedNow = [datetime]::new($Now.Year, $Now.Month, $Now.Day, $Now.Hour, 0, 0)
    $match = [regex]::Match($text, '(?i)(最近|近|last)\s*(\d{1,2})\s*(天|日|days?|d)')
    if ($match.Success) {
        $days = [int]$match.Groups[2].Value
        return Get-RelativeLogTimeRange -Now $roundedNow -Days $days
    }

    $match = [regex]::Match($text, '(?i)(最近|近|last)\s*(\d{1,3})\s*(小时|小時|hours?|hrs?|h)')
    if ($match.Success) {
        $hours = [int]$match.Groups[2].Value
        if ($hours -le 0 -or $hours -gt (90 * 24)) {
            throw 'Hours must be between 1 and 2160.'
        }
        $startTime = $roundedNow.AddHours(-$hours)
        return [PSCustomObject]@{
            StartTime = $startTime
            EndTime = $roundedNow
            AnalysisDateStr = "$($startTime.ToString('yyyyMMddHHmm'))_$($roundedNow.ToString('yyyyMMddHHmm'))"
            AnalysisDateDisplay = "$($startTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($roundedNow.ToString('yyyy-MM-dd HH:mm:ss'))"
        }
    }

    if ($text -match '(?i)最近\s*三\s*(小时|小時)|last\s*3\s*(hours?|hrs?|h)') {
        return Get-RelativeLogTimeRange -Now $roundedNow -Days 0
    }

    throw "Cannot parse time range from prompt: $Prompt. Try '查询最近15天的微软日志' or use -StartDate/-EndDate."
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
        return 0
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
    return $recordCount
}

if ($CustomStart -and $CustomEnd) {
    $StartTime = [DateTime]::Parse($CustomStart)
    $EndTime = [DateTime]::Parse($CustomEnd)
    Assert-LogTimeRangeWithinLimit -StartTime $StartTime -EndTime $EndTime
    $AnalysisDateStr = "$($StartTime.ToString('yyyyMMddHHmm'))_$($EndTime.ToString('yyyyMMddHHmm'))"
    $AnalysisDateDisplay = "$($StartTime.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))"
}
elseif ($Prompt) {
    $range = Resolve-SkillPromptTimeRange -Prompt $Prompt -Now $Now
    $StartTime = $range.StartTime
    $EndTime = $range.EndTime
    $AnalysisDateStr = $range.AnalysisDateStr
    $AnalysisDateDisplay = $range.AnalysisDateDisplay
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

if ($targetTables -notcontains 'DCRLogErrors') {
    $targetTables += 'DCRLogErrors'
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
Write-Host "License Graph verification: $(if($SkipLicenseGraph){'Skipped for speed'}else{'Enabled'})" -ForegroundColor Cyan
Write-Host "Total count precheck: $(if($SkipTotalCount){'Skipped'}else{'Enabled'})" -ForegroundColor Cyan
if (-not $SkipTotalCount) { Write-Host "Total count timeout: ${TotalCountTimeoutSec}s" -ForegroundColor Cyan }
Write-Host "HTML: $($reportPaths.HtmlFile)" -ForegroundColor Cyan
Write-Host ''

$csvFiles = [System.Collections.Generic.List[string]]::new()
$csvTables = [System.Collections.Generic.List[string]]::new()
$csvTotalCounts = [System.Collections.Generic.List[int]]::new()
$diagnosticRows = [System.Collections.Generic.List[object]]::new()
$overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($table in $targetTables) {
    $tableStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $sourceKind = 'Query'
    $recordCount = $null
    $paths = Get-LogArtifactPaths -TempDir $TempDir -TableName $table -AnalysisDateStr $AnalysisDateStr -Now $Now
    $cacheBypassTables = @(
        'DCRLogErrors',
        'IntuneAuditLogsDCR_CL',
        'MailboxStatisticsDCR_CL',
        'AADManagedIdentitySignInLogs',
        'AADServicePrincipalSignInLogs',
        'SigninLogs'
    )
    $useTableCache = $UseCache -and $table -notin $cacheBypassTables
    Write-Host "[1/3] Data source: $table" -ForegroundColor Yellow
    if ($table -eq 'DCRLogErrors') {
        Write-Host '  Cache skipped: DCRLogErrors uses fixed last-30-days distinct KQL.' -ForegroundColor Yellow
    }
    if ($table -eq 'MailboxStatisticsDCR_CL') {
        Write-Host '  Cache skipped: MailboxStatisticsDCR_CL uses low-space and SharedMailbox KQL.' -ForegroundColor Yellow
    }
    if ($table -eq 'IntuneAuditLogsDCR_CL') {
        Write-Host '  Cache skipped: IntuneAuditLogsDCR_CL uses current audit-record KQL.' -ForegroundColor Yellow
    }
    if ($table -in @('AADManagedIdentitySignInLogs', 'AADServicePrincipalSignInLogs', 'SigninLogs')) {
        Write-Host '  Cache skipped: sign-in risk depends on current trusted IP rules.' -ForegroundColor Yellow
    }

    $cacheResult = $null
    if ($useTableCache -and -not $ForceRefresh) {
        $cacheResult = Test-Cache -CacheCsv $paths.CacheCsv -CacheMeta $paths.CacheMeta -CacheTTL $CacheTTL -TableName $table -StartTime $StartTime -EndTime $EndTime
    }

    # Get total record count for the table (without risk filter)
    $totalCount = 0
    if ($SkipTotalCount) {
        Write-Host "  Total record count skipped (-SkipTotalCount)." -ForegroundColor DarkGray
    } else {
    try {
        $totalCountQuery = New-TableTotalCountQuery -TableName $table -StartTime $StartTime -EndTime $EndTime
        Write-Host "  Getting total record count... (timeout: ${TotalCountTimeoutSec}s; if this is slow, rerun with -SkipTotalCount)" -ForegroundColor DarkGray
        # Do NOT print the query to avoid exposing IP addresses
        # Note: Do NOT pass -TableName here, only pass -Query, otherwise the script will 
        # auto-generate a query based on the table name and ignore our count query
        $totalCountResult = & "$ScriptDir\query-log-analytics.ps1" -Query $totalCountQuery -StartTime $StartTime.ToString('o') -EndTime $EndTime.ToString('o') -RawCount -NoProfile -QueryTimeoutSec $TotalCountTimeoutSec -ErrorAction Stop 2>&1
        $exitCode = $LASTEXITCODE
        Write-Host "  Exit code: $exitCode" -ForegroundColor DarkGray
        Write-Host "  Raw result type: $($totalCountResult.GetType().Name)" -ForegroundColor DarkGray
        Write-Host "  Raw result: $totalCountResult" -ForegroundColor DarkGray
        
        # Parse the output - handle both string and array of objects
        $resultString = if ($totalCountResult -is [array]) {
            $totalCountResult | ForEach-Object { $_.ToString() } | Where-Object { $_ -match '^\d+$' } | Select-Object -Last 1
        } else {
            $totalCountResult.ToString()
        }
        
        if ($resultString -match '^\d+$') {
            $totalCount = [int]$resultString
        }
        
        Write-Host "  Parsed total count: $totalCount" -ForegroundColor DarkGray
    } catch {
        Write-Host "  Failed to get total count: $_" -ForegroundColor DarkYellow
        Write-Host "  Continue without total count for $table. You can use -SkipTotalCount to bypass this precheck." -ForegroundColor DarkYellow
        Write-Host "  Exception type: $($_.Exception.GetType().Name)" -ForegroundColor DarkYellow
    }
    }

    if ($cacheResult -and $cacheResult.Hit) {
        $sourceKind = 'Cache'
        $recordCount = $cacheResult.RecordCount
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

            & "$ScriptDir\query-log-analytics.ps1" @queryParams
        } else {
            $queryArgs = @(
                '-NoProfile',
                '-ExecutionPolicy', 'Bypass',
                '-File', (Join-Path $ScriptDir 'query-log-analytics.ps1'),
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
        # Even if CSV file doesn't exist (filtered count = 0), include the table if total count > 0
        if ($totalCount -gt 0) {
            Write-Host "  Query did not generate CSV for $table (filtered=0), but total=$totalCount; including in report." -ForegroundColor Yellow
            $csvTables.Add($table) | Out-Null
            $csvTotalCounts.Add($totalCount) | Out-Null
            # Create an empty CSV with header for report generation
            'TimeGenerated' | Out-File -FilePath $paths.CsvFile -Encoding UTF8 -Force
            $csvFiles.Add($paths.CsvFile) | Out-Null
        } else {
            Write-Host "  Query did not generate CSV for $table; skipping." -ForegroundColor Yellow
        }
        continue
    }
        if ($useTableCache) {
            $recordCount = Save-Cache -SourceCsv $paths.CsvFile -CacheCsv $paths.CacheCsv -CacheMeta $paths.CacheMeta -TableName $table -CacheTTL $CacheTTL -StartTime $StartTime -EndTime $EndTime
        }
    }

    if ($null -eq $recordCount -and (Test-Path $paths.CsvFile)) {
        $recordCount = Get-LogCsvRecordCount -CsvPath $paths.CsvFile
    }
    $fileSizeBytes = if (Test-Path $paths.CsvFile) { (Get-Item -Path $paths.CsvFile).Length } else { 0 }
    $tableStopwatch.Stop()
    $diagnosticRows.Add([PSCustomObject]@{
        Table = $table
        Source = $sourceKind
        Records = if ($null -ne $recordCount) { [int]$recordCount } else { 0 }
        Size = Format-LogByteSize -Bytes $fileSizeBytes
        Seconds = [Math]::Round($tableStopwatch.Elapsed.TotalSeconds, 1)
    }) | Out-Null
    Write-Host ("  Diagnostics: source={0}, records={1}, size={2}, elapsed={3}" -f $sourceKind, $(if ($null -ne $recordCount) { $recordCount } else { 0 }), (Format-LogByteSize -Bytes $fileSizeBytes), (Format-LogElapsed -Elapsed $tableStopwatch.Elapsed)) -ForegroundColor DarkCyan

    $csvFiles.Add($paths.CsvFile) | Out-Null
    $csvTables.Add($table) | Out-Null
    $csvTotalCounts.Add($totalCount) | Out-Null
}

if ($csvFiles.Count -eq 0) {
    throw 'No CSV files were available for report generation.'
}

Write-Host ''
Write-Host '[2/3] Generating merged HTML report...' -ForegroundColor Yellow
$analysisStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
# 计算 UTC 时间用于 KQL 语句显示
$startUtcIso = $StartTime.ToUniversalTime().ToString('o')
$endUtcIso = $EndTime.ToUniversalTime().ToString('o')
$htmlParams = @{
    CsvPath = $csvFiles.ToArray()
    OutputPath = $HtmlFilePath
    AnalysisDate = $AnalysisDateDisplay
    TableName = $csvTables.ToArray()
    TotalCounts = $csvTotalCounts.ToArray()
    StartUtc = $startUtcIso
    EndUtc = $endUtcIso
}
# 默认启用 License Graph 验证，除非明确指定 -SkipLicenseGraph
if ($SkipLicenseGraph) { $htmlParams.SkipLicenseGraph = $true }
& "$ScriptDir\generate-html-report.ps1" @htmlParams
$analysisStopwatch.Stop()

if (-not (Test-Path $HtmlFilePath)) {
    throw 'HTML file was not generated.'
}

$htmlUrl = "file:///$($HtmlFilePath -replace '\\', '/')"
$overallStopwatch.Stop()
Write-Host ''
Write-Host '[3/3] Done' -ForegroundColor Yellow
Write-Host "HTML: $HtmlFilePath" -ForegroundColor Cyan
Write-Host "URL: $htmlUrl" -ForegroundColor Cyan
Write-Host ''
Write-Host 'Performance summary:' -ForegroundColor Cyan
foreach ($row in $diagnosticRows) {
    Write-Host ("  {0}: {1}, {2} records, {3}, {4:N1}s" -f $row.Table, $row.Source, $row.Records, $row.Size, $row.Seconds) -ForegroundColor Cyan
}
Write-Host ("  HTML analysis: {0}" -f (Format-LogElapsed -Elapsed $analysisStopwatch.Elapsed)) -ForegroundColor Cyan
Write-Host ("  Total elapsed: {0}" -f (Format-LogElapsed -Elapsed $overallStopwatch.Elapsed)) -ForegroundColor Cyan

if (-not $NoOpen) {
    Start-Process $HtmlFilePath
}
