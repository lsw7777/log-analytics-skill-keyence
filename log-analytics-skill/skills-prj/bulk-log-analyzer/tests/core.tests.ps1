$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir = Split-Path -Parent $scriptDir
. (Join-Path $rootDir 'log-analyzer-core.ps1')

function Assert-Equal {
    param(
        [object]$Actual,
        [object]$Expected,
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected', got '$Actual'."
    }
}

$tables = Get-SupportedLogTables
Assert-Equal $tables.Count 7 'Supported table count mismatch.'
Assert-Equal $tables[0].Name 'AssignedLicensesDCR_CL' 'First menu table mismatch.'
Assert-Equal $tables[1].Name 'AuditGeneralDCR_CL' 'Second menu table mismatch.'
Assert-Equal (Resolve-LogTableSelection -Selection '2') 'AuditGeneralDCR_CL' 'Menu selection did not resolve expected table.'

$range = Get-DefaultLogTimeRange -Now ([datetime]'2026-05-25T13:45:00')
Assert-Equal $range.StartTime.ToString('yyyy-MM-dd HH:mm:ss') '2026-05-24 00:00:00' 'Default start time mismatch.'
Assert-Equal $range.EndTime.ToString('yyyy-MM-dd HH:mm:ss') '2026-05-25 00:00:00' 'Default end time mismatch.'
Assert-Equal $range.AnalysisDateStr '20260524' 'Analysis date mismatch.'

$singleDayRange = Get-LogTimeRangeFromDates -StartDate '2026-05-24' -EndDate '2026-05-24'
Assert-Equal $singleDayRange.StartTime.ToString('yyyy-MM-dd HH:mm:ss') '2026-05-24 00:00:00' 'Single-day start mismatch.'
Assert-Equal $singleDayRange.EndTime.ToString('yyyy-MM-dd HH:mm:ss') '2026-05-25 00:00:00' 'Single-day inclusive end mismatch.'
Assert-Equal $singleDayRange.AnalysisDateStr '20260524' 'Single-day analysis date mismatch.'
Assert-Equal $singleDayRange.AnalysisDateDisplay '2026-05-24' 'Single-day display date mismatch.'

$multiDayRange = Get-LogTimeRangeFromDates -StartDate '2026-05-20' -EndDate '2026-05-24'
Assert-Equal $multiDayRange.StartTime.ToString('yyyy-MM-dd HH:mm:ss') '2026-05-20 00:00:00' 'Multi-day start mismatch.'
Assert-Equal $multiDayRange.EndTime.ToString('yyyy-MM-dd HH:mm:ss') '2026-05-25 00:00:00' 'Multi-day inclusive end mismatch.'
Assert-Equal $multiDayRange.AnalysisDateStr '20260520_20260524' 'Multi-day analysis date mismatch.'
Assert-Equal $multiDayRange.AnalysisDateDisplay '2026-05-20 to 2026-05-24' 'Multi-day display date mismatch.'

$invalidRangeFailed = $false
try {
    Get-LogTimeRangeFromDates -StartDate '2026-05-25' -EndDate '2026-05-24' | Out-Null
} catch {
    $invalidRangeFailed = $true
}
Assert-Equal $invalidRangeFailed $true 'Invalid date range should fail.'

$paths = Get-LogArtifactPaths -TempDir 'C:\Temp' -TableName 'AuditGeneralDCR_CL' -AnalysisDateStr '20260524' -Now ([datetime]'2026-05-25T13:45:00')
Assert-Equal $paths.HtmlFile '.\final_report_AuditGeneralDCR_CL_20260524_1345.html' 'HTML file naming mismatch.'
Assert-Equal $paths.CsvFile 'C:\Temp\AuditGeneralDCR_CL_20260524.csv' 'CSV file naming mismatch.'

$resolvedHtml = Join-Path (Get-Location) $paths.HtmlFile
if ($resolvedHtml -notlike '*final_report_AuditGeneralDCR_CL_20260524_1345.html') {
    throw "Relative HTML path should resolve under current repository root, got '$resolvedHtml'."
}
if ([System.IO.Path]::IsPathRooted($paths.HtmlFile)) {
    throw "HTML display path must be relative, got '$($paths.HtmlFile)'."
}
if (-not [System.IO.Path]::IsPathRooted($paths.HtmlFilePath)) {
    throw "HTML write path must be absolute for reliable file creation, got '$($paths.HtmlFilePath)'."
}

$auditProfile = Get-TableAnalysisProfile -TableName 'AuditGeneralDCR_CL'
Assert-Equal ($auditProfile.GroupFields -join ',') 'Activity,Operation,Workload' 'AuditGeneral grouping fields mismatch.'
Assert-Equal $auditProfile.UseCompositeOperationGroup $true 'AuditGeneral should use composite operation grouping.'

$sharePointProfile = Get-TableAnalysisProfile -TableName 'SharePointAuditDCR_CL'
Assert-Equal ($sharePointProfile.GroupFields -join ',') 'Activity,Operation,Workload' 'SharePointAudit grouping fields mismatch.'
Assert-Equal $sharePointProfile.UseCompositeOperationGroup $true 'SharePointAudit should use composite operation grouping.'

$groupRow = [PSCustomObject]@{ Activity = 'FileAccessed'; Operation = 'Open'; Workload = 'SharePoint' }
Assert-Equal (Get-OperationGroupValue -Row $groupRow -TableName 'AuditGeneralDCR_CL') 'FileAccessed | Open | SharePoint' 'Composite operation grouping mismatch.'

$query = New-LogTableQuery -TableName 'WQCLogDCR_CL' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00')
if ($query -notmatch '^WQCLogDCR_CL \| where TimeGenerated >= datetime\(2026-05-23T16:00:00Z') {
    throw "Query does not start with expected table and start filter: $query"
}
if ($query -notmatch 'TimeGenerated < datetime\(2026-05-24T16:00:00Z') {
    throw "Query does not include expected end filter: $query"
}

$matchingMeta = [PSCustomObject]@{
    TableName = 'AuditGeneralDCR_CL'
    StartTimeUtc = Get-LogCacheTimeKey -Time ([datetime]'2026-05-24T00:00:00+08:00')
    EndTimeUtc = Get-LogCacheTimeKey -Time ([datetime]'2026-05-25T00:00:00+08:00')
}
Assert-Equal (Test-LogCacheMetadataMatches -Meta $matchingMeta -TableName 'AuditGeneralDCR_CL' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00')) $true 'Cache metadata should match same table and time range.'
Assert-Equal (Test-LogCacheMetadataMatches -Meta $matchingMeta -TableName 'AuditGeneralDCR_CL' -StartTime ([datetime]'2026-05-24T01:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00')) $false 'Cache metadata should reject different time range.'
Assert-Equal (Test-LogCacheMetadataMatches -Meta $matchingMeta -TableName 'WQCLogDCR_CL' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00')) $false 'Cache metadata should reject different table.'

$cachePayloadDir = Join-Path ([System.IO.Path]::GetTempPath()) 'log-analyzer-cache-payload-tests'
if (Test-Path $cachePayloadDir) { Remove-Item -Path $cachePayloadDir -Recurse -Force }
New-Item -ItemType Directory -Path $cachePayloadDir -Force | Out-Null
$emptyCsv = Join-Path $cachePayloadDir 'empty.csv'
$headerOnlyCsv = Join-Path $cachePayloadDir 'header-only.csv'
$dataCsv = Join-Path $cachePayloadDir 'data.csv'
'' | Out-File -FilePath $emptyCsv -Encoding UTF8 -Force
'"TimeGenerated"' | Out-File -FilePath $headerOnlyCsv -Encoding UTF8 -Force
'"TimeGenerated"' | Out-File -FilePath $dataCsv -Encoding UTF8 -Force
'"2026-05-26T00:00:00Z"' | Out-File -FilePath $dataCsv -Encoding UTF8 -Append
Assert-Equal (Get-LogCsvRecordCount -CsvPath $emptyCsv) 0 'Empty CSV should have zero records.'
Assert-Equal (Get-LogCsvRecordCount -CsvPath $headerOnlyCsv) 0 'Header-only CSV should have zero records.'
Assert-Equal (Get-LogCsvRecordCount -CsvPath $dataCsv) 1 'Data CSV should count records.'
Assert-Equal (Test-LogCachePayloadValid -CacheCsv $emptyCsv -RecordCount 0) $false 'Zero-record empty cache payload should be invalid.'
Assert-Equal (Test-LogCachePayloadValid -CacheCsv $headerOnlyCsv -RecordCount 0) $false 'Zero-record header-only cache payload should be invalid.'
Assert-Equal (Test-LogCachePayloadValid -CacheCsv $dataCsv -RecordCount 1) $true 'Non-empty cache payload should be valid.'

$clearCacheDir = Join-Path ([System.IO.Path]::GetTempPath()) 'log-analyzer-clear-cache-tests'
if (Test-Path $clearCacheDir) { Remove-Item -Path $clearCacheDir -Recurse -Force }
New-Item -ItemType Directory -Path $clearCacheDir -Force | Out-Null
'x' | Out-File -FilePath (Join-Path $clearCacheDir 'sample.csv') -Encoding UTF8 -Force
'{}' | Out-File -FilePath (Join-Path $clearCacheDir 'sample.meta.json') -Encoding UTF8 -Force
Clear-LogCache -CacheDir $clearCacheDir | Out-Null
if (-not (Test-Path $clearCacheDir)) { throw 'Clear cache should preserve cache directory.' }
Assert-Equal @(Get-ChildItem -Path $clearCacheDir -Force).Count 0 'Clear cache should remove all cache entries.'

$explicitRangeMode = Get-LogQueryExecutionMode -QueryStartTime ([datetime]'2026-05-25T16:00:00Z') -QueryEndTime ([datetime]'2026-05-26T16:00:00Z') -Hours 24
Assert-Equal $explicitRangeMode.UseTimespan $false 'Explicit date range should not use relative timespan.'
Assert-Equal $explicitRangeMode.Timespan $null 'Explicit date range should not set timespan.'

$relativeRangeMode = Get-LogQueryExecutionMode -Hours 24
Assert-Equal $relativeRangeMode.UseTimespan $true 'Relative query should use timespan.'
Assert-Equal $relativeRangeMode.Timespan.TotalHours 24 'Relative query timespan mismatch.'

Write-Host 'core.tests.ps1 passed' -ForegroundColor Green
