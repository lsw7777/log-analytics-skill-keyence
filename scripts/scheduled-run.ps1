param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptDir 'log-analyzer-core.ps1')

if (-not (Test-Path $ConfigPath)) {
    throw "Schedule config not found: $ConfigPath"
}

$config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
$tables = @($config.Tables)
if (-not $tables -or $tables.Count -eq 0) {
    $tables = @($SupportedLogTables | ForEach-Object { $_.Name })
}

foreach ($table in $tables) {
    & (Join-Path $scriptDir 'run-all.ps1') -TableName $table -AnalysisDate (Get-Date).AddDays(-1).ToString('yyyy-MM-dd') -NoOpen
}
