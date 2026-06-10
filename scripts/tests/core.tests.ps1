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
Assert-Equal $tables.Count 8 'Supported table count mismatch.'
Assert-Equal $tables[0].Name 'AADManagedIdentitySignInLogs' 'First menu table mismatch.'
Assert-Equal $tables[1].Name 'AADServicePrincipalSignInLogs' 'Second menu table mismatch.'
Assert-Equal (Resolve-LogTableSelection -Selection '4') 'AuditLogs' 'Menu selection did not resolve expected table.'

$range = Get-DefaultLogTimeRange -Now ([datetime]'2026-05-25T13:45:00')
Assert-Equal $range.StartTime.ToString('yyyy-MM-dd HH:mm:ss') '2026-05-25 10:45:00' 'Default start time mismatch.'
Assert-Equal $range.EndTime.ToString('yyyy-MM-dd HH:mm:ss') '2026-05-25 13:45:00' 'Default end time mismatch.'
Assert-Equal $range.AnalysisDateStr '202605251045_202605251345' 'Analysis date mismatch.'
Assert-Equal $range.AnalysisDateDisplay '2026-05-25 10:45:00 to 2026-05-25 13:45:00' 'Analysis date display mismatch.'

$oneDayRelativeRange = Get-RelativeLogTimeRange -Now ([datetime]'2026-05-25T13:45:00') -Days 1
Assert-Equal $oneDayRelativeRange.StartTime.ToString('yyyy-MM-dd HH:mm:ss') '2026-05-24 13:45:00' 'One-day relative start mismatch.'
Assert-Equal $oneDayRelativeRange.EndTime.ToString('yyyy-MM-dd HH:mm:ss') '2026-05-25 13:45:00' 'One-day relative end mismatch.'

$tooManyRelativeDaysFailed = $false
try {
    Get-RelativeLogTimeRange -Now ([datetime]'2026-05-25T13:45:00') -Days 91 | Out-Null
} catch {
    $tooManyRelativeDaysFailed = $true
}
Assert-Equal $tooManyRelativeDaysFailed $true 'Relative day selection should reject more than 90 days.'

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

$tooLongRangeFailed = $false
try {
    Get-LogTimeRangeFromDates -StartDate '2026-01-01' -EndDate '2026-05-10' | Out-Null
} catch {
    $tooLongRangeFailed = $true
}
Assert-Equal $tooLongRangeFailed $true 'Date range longer than 90 days should fail.'

$paths = Get-LogArtifactPaths -TempDir 'C:\Temp' -TableName 'AuditLogs' -AnalysisDateStr '20260524' -Now ([datetime]'2026-05-25T13:45:00')
Assert-Equal $paths.HtmlFile '.\final_report_AuditLogs_20260524_1345.html' 'HTML file naming mismatch.'
Assert-Equal $paths.CsvFile 'C:\Temp\AuditLogs_20260524.csv' 'CSV file naming mismatch.'

$resolvedHtml = Join-Path (Get-Location) $paths.HtmlFile
if ($resolvedHtml -notlike '*final_report_AuditLogs_20260524_1345.html') {
    throw "Relative HTML path should resolve under current repository root, got '$resolvedHtml'."
}
if ([System.IO.Path]::IsPathRooted($paths.HtmlFile)) {
    throw "HTML display path must be relative, got '$($paths.HtmlFile)'."
}
if (-not [System.IO.Path]::IsPathRooted($paths.HtmlFilePath)) {
    throw "HTML write path must be absolute for reliable file creation, got '$($paths.HtmlFilePath)'."
}

$mergedPaths = Get-MergedLogArtifactPaths -TempDir 'C:\Temp' -AnalysisDateStr '20260518_20260524' -Now ([datetime]'2026-05-25T13:45:00')
Assert-Equal $mergedPaths.HtmlFile '.\final_report_merged_20260518_20260524_1345.html' 'Merged HTML file naming mismatch.'

$auditLogsProfile = Get-TableAnalysisProfile -TableName 'AuditLogs'
Assert-Equal $auditLogsProfile.UserFields[0] 'Actor' 'AuditLogs should prefer extracted user actor.'

Assert-Equal (Test-DeleteOrDisableOperation -Operation 'Delete user' -TableName 'AuditLogs') $true 'Delete operation should be high privilege.'
Assert-Equal (Test-DeleteOrDisableOperation -Operation 'Disable device' -TableName 'IntuneAuditLogsDCR_CL') $true 'Disable operation should be high privilege.'
Assert-Equal (Test-DeleteOrDisableOperation -Operation 'Search' -TableName 'AuditLogs') $false 'Search should not be high privilege.'

$trustedRules = Get-TrustedIpRules -Paths @((Join-Path (Get-Location) 'scripts\config\TrustedLocation_IDC_Ali.txt'), (Join-Path (Get-Location) 'scripts\config\TrustedLocation_KJ.txt')) -IncludeMicrosoft:$false
Assert-Equal (Test-IpInTrustedRules -IP '47.102.133.2' -Rules $trustedRules) $true 'Trusted IP should match rules.'
Assert-Equal (Test-IpInTrustedRules -IP '47.102.133.2:443' -Rules $trustedRules) $true 'Trusted IP with port should match rules.'
Assert-Equal (Test-IpInTrustedRules -IP 'client=47.102.133.2:443' -Rules $trustedRules) $true 'Trusted IP with prefix and port should match rules.'
Assert-Equal (Test-IpInTrustedRules -IP '8.8.8.8' -Rules $trustedRules) $false 'Untrusted IP should not match rules.'
Assert-Equal (Test-PrivateOrInvalidIp -IP '10.1.2.3:443') $true 'Private IP with port should be invalid for public suspicious IP checks.'
Assert-Equal (Test-PrivateOrInvalidIp -IP '127.0.0.1:443') $true 'Loopback IP with port should be invalid for public suspicious IP checks.'
Assert-Equal (Test-PrivateOrInvalidIp -IP 'client=0.0.0.0:123') $true 'Unspecified IP with prefix and port should be invalid for public suspicious IP checks.'

foreach ($unsupportedTable in @('AuditGeneralDCR_CL', 'AzureADUsersDCR_CL', 'MessageTraceDataDCR_CL', 'SharePointAuditDCR_CL', 'WQCLogDCR_CL')) {
    $selectionFailed = $false
    try {
        Resolve-LogTableSelection -Selection $unsupportedTable | Out-Null
    } catch {
        $selectionFailed = $true
    }
    Assert-Equal $selectionFailed $true "$unsupportedTable should not be a supported table."
}

$riskSigninQuery = New-LogTableQuery -TableName 'SigninLogs' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00') -RiskOnly
if ($riskSigninQuery -notmatch '__isSigninSuspiciousSuccess') {
    throw "Risk-only SigninLogs query should include suspicious success prefilter: $riskSigninQuery"
}
if ($riskSigninQuery -notmatch '__isPublicUntrustedIp') {
    throw "Risk-only query should include public untrusted IP prefilter: $riskSigninQuery"
}
if ($riskSigninQuery -notmatch 'summarize TimeGenerated=max\(TimeGenerated\), FirstTime=min\(TimeGenerated\), LastTime=max\(TimeGenerated\), EventCount=count\(\)') {
    throw "Risk-only SigninLogs query should aggregate duplicate rows in KQL: $riskSigninQuery"
}
if ($riskSigninQuery -match 'offhour|OffHour|hour\(') {
    throw "Risk-only query should not include off-hours activity logic: $riskSigninQuery"
}

$riskSpSigninQuery = New-LogTableQuery -TableName 'AADServicePrincipalSignInLogs' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00') -RiskOnly
if ($riskSpSigninQuery -notmatch 'ServicePrincipalName') {
    throw "Risk-only SP sign-in query should explicitly project ServicePrincipalName: $riskSpSigninQuery"
}
if ($riskSpSigninQuery -notmatch 'where EventCount > 10') {
    throw "Risk-only SP sign-in query should keep only failures above 10 events: $riskSpSigninQuery"
}
if ($riskSpSigninQuery -notmatch 'AggregatedSuspiciousSigninSuccess') {
    throw "Risk-only SP sign-in query should include suspicious successful sign-ins: $riskSpSigninQuery"
}

$riskAuditLogsQuery = New-LogTableQuery -TableName 'AuditLogs' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00') -RiskOnly
if ($riskAuditLogsQuery -notmatch 'Add app role assignment to service principal') {
    throw "Risk-only AuditLogs query should include service-principal app role changes: $riskAuditLogsQuery"
}
if ($riskAuditLogsQuery -notmatch '__actorIsUser') {
    throw "Risk-only AuditLogs query should require user actors: $riskAuditLogsQuery"
}
if ($riskAuditLogsQuery -notmatch 'PIM activation expired') {
    throw "Risk-only AuditLogs query should filter PIM activation expired noise: $riskAuditLogsQuery"
}
if ($riskAuditLogsQuery -match 'sort by TimeGenerated') {
    throw "Risk-only AuditLogs query should not sort full raw result sets: $riskAuditLogsQuery"
}

$riskDcrQuery = New-LogTableQuery -TableName 'DCRLogErrors' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00') -RiskOnly
if ($riskDcrQuery -notmatch 'DCRLogErrors') {
    throw "Risk-only DCRLogErrors query should query DCRLogErrors: $riskDcrQuery"
}
if ($riskDcrQuery -notmatch 'InputStreamId, OperationName, Message') {
    throw "Risk-only DCRLogErrors query should aggregate by InputStreamId, OperationName, Message: $riskDcrQuery"
}

$riskLicenseQuery = New-LogTableQuery -TableName 'AssignedLicensesDCR_CL' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00') -RiskOnly
if ($riskLicenseQuery -notmatch 'summarize TimeGenerated=max\(TimeGenerated\), UsedUsers=dcount\(UserPrincipalName\)') {
    throw "Risk-only AssignedLicenses query should summarize license usage before export: $riskLicenseQuery"
}

$riskIntuneQuery = New-LogTableQuery -TableName 'IntuneAuditLogsDCR_CL' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00') -RiskOnly
if ($riskIntuneQuery -notmatch 'Actor') {
    throw "Risk-only Intune query should project Actor fields: $riskIntuneQuery"
}
if ($riskIntuneQuery -notmatch 'TargetDisplayName') {
    throw "Risk-only Intune query should project TargetDisplayName fields: $riskIntuneQuery"
}

$riskMailboxQuery = New-LogTableQuery -TableName 'MailboxStatisticsDCR_CL' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00') -RiskOnly
if ($riskMailboxQuery -notmatch 'LatestMailboxRiskSnapshot') {
    throw "Risk-only Mailbox query should export latest risk snapshots only: $riskMailboxQuery"
}

$matchingMeta = [PSCustomObject]@{
    TableName = 'AuditLogs'
    StartTimeUtc = Get-LogCacheTimeKey -Time ([datetime]'2026-05-24T00:00:00+08:00')
    EndTimeUtc = Get-LogCacheTimeKey -Time ([datetime]'2026-05-25T00:00:00+08:00')
}
Assert-Equal (Test-LogCacheMetadataMatches -Meta $matchingMeta -TableName 'AuditLogs' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00')) $true 'Cache metadata should match same table and time range.'
Assert-Equal (Test-LogCacheMetadataMatches -Meta $matchingMeta -TableName 'AuditLogs' -StartTime ([datetime]'2026-05-24T01:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00')) $false 'Cache metadata should reject different time range.'
Assert-Equal (Test-LogCacheMetadataMatches -Meta $matchingMeta -TableName 'SigninLogs' -StartTime ([datetime]'2026-05-24T00:00:00+08:00') -EndTime ([datetime]'2026-05-25T00:00:00+08:00')) $false 'Cache metadata should reject different table.'

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
