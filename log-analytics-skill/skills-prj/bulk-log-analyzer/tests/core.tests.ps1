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
Assert-Equal $paths.HtmlFile 'C:\Temp\AuditGeneralDCR_CL_20260525_1345.html' 'HTML file naming mismatch.'
Assert-Equal $paths.CsvFile 'C:\Temp\AuditGeneralDCR_CL_20260524.csv' 'CSV file naming mismatch.'

$query = New-LogTableQuery -TableName 'WQCLogDCR_CL' -StartTime ([datetime]'2026-05-24T00:00:00') -EndTime ([datetime]'2026-05-25T00:00:00')
if ($query -notmatch '^WQCLogDCR_CL \| where TimeGenerated >= datetime\(2026-05-24T00:00:00') {
    throw "Query does not start with expected table and start filter: $query"
}
if ($query -notmatch 'TimeGenerated < datetime\(2026-05-25T00:00:00') {
    throw "Query does not include expected end filter: $query"
}

Write-Host 'core.tests.ps1 passed' -ForegroundColor Green
