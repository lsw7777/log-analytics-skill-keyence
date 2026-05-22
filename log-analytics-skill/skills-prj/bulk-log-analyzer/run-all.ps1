# ============================================================
# Log Analytics One-Click Script
# ============================================================

param(
    [Parameter(Mandatory = $false)]
    [string]$TableName = "AuditGeneralDCR_CL",

    [Parameter(Mandatory = $false)]
    [int]$Hours = 24,

    [Parameter(Mandatory = $false)]
    [switch]$ForceLogin
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$TempDir = "$env:USERPROFILE\AppData\Local\Temp\opencode"
$DateStr = Get-Date -Format "yyyyMMdd_HHmm"
$CsvFile = "$TempDir\$($TableName.Substring(0, $TableName.IndexOf('_DCR')))_$DateStr.csv"
$HtmlFile = "$TempDir\report_$DateStr.html"

if (-not (Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
}

Write-Host "============================================" -ForegroundColor Magenta
Write-Host "  Log Analytics One-Click" -ForegroundColor Magenta
Write-Host "============================================" -ForegroundColor Magenta
Write-Host "Table: $TableName" -ForegroundColor Cyan
Write-Host "Hours: $Hours" -ForegroundColor Cyan
Write-Host "CSV: $CsvFile" -ForegroundColor Cyan
Write-Host "HTML: $HtmlFile" -ForegroundColor Cyan
Write-Host ""

# Step 1: Query data
Write-Host "[1/4] Querying log data..." -ForegroundColor Yellow

$QueryParams = @{
    TableName = $TableName
    Hours     = $Hours
    ExportCsv = $CsvFile
}

if ($ForceLogin) {
    $QueryParams['ForceLogin'] = $true
}

& "$ScriptDir\azure_log_query.ps1" @QueryParams

if (-not (Test-Path $CsvFile)) {
    Write-Host "Error: CSV file not generated" -ForegroundColor Red
    exit 1
}

Write-Host "Data query complete!" -ForegroundColor Green
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

Write-Host "HTML report generated: $HtmlFile" -ForegroundColor Green
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