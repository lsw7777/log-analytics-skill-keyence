$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir = Split-Path -Parent $scriptDir
. (Join-Path $rootDir 'log-analyzer-core.ps1')

function Assert-Equal {
    param([object]$Actual, [object]$Expected, [string]$Message)
    if ($Actual -ne $Expected) { throw "$Message Expected '$Expected', got '$Actual'." }
}

function Assert-Contains {
    param([string]$Content, [string]$Expected, [string]$Message)
    if (-not $Content.Contains($Expected)) { throw "$Message Expected content to include '$Expected'." }
}

$config = New-LogAnalyzerScheduleConfig
Assert-Equal $config.RunAt '01:00' 'Default schedule time mismatch.'
Assert-Equal $config.Tables.Count 7 'Default schedule table count mismatch.'
Assert-Equal $config.Tables[0] 'AssignedLicensesDCR_CL' 'Default schedule first table mismatch.'
Assert-Equal $config.Tables[6] 'WQCLogDCR_CL' 'Default schedule last table mismatch.'

$customConfig = New-LogAnalyzerScheduleConfig -RunAt '03:30' -Tables @('AuditGeneralDCR_CL', 'WQCLogDCR_CL')
Assert-Equal $customConfig.RunAt '03:30' 'Custom schedule time mismatch.'
Assert-Equal $customConfig.Tables.Count 2 'Custom schedule table count mismatch.'

$batchCommand = Get-LogAnalyzerBatchCommand -RootDir $rootDir -ConfigPath 'C:\Temp\schedule-config.json'
Assert-Contains $batchCommand 'scheduled-run.ps1' 'Batch command should reference scheduled-run.ps1.'
Assert-Contains $batchCommand '-ConfigPath' 'Batch command should pass config path.'

$trayCommand = Get-LogAnalyzerTrayCommand -RootDir $rootDir -ConfigPath 'C:\Temp\schedule-config.json'
Assert-Contains $trayCommand 'tray.ps1' 'Tray command should reference tray.ps1.'
Assert-Contains $trayCommand '-WindowStyle Hidden' 'Tray command should hide console window.'

$nextRun = Get-LogAnalyzerNextRunTime -RunAt '01:00' -Now ([datetime]'2026-05-26T00:30:00')
Assert-Equal $nextRun.ToString('yyyy-MM-dd HH:mm') '2026-05-26 01:00' 'Next run time should stay on same day before schedule.'

$nextDayRun = Get-LogAnalyzerNextRunTime -RunAt '01:00' -Now ([datetime]'2026-05-26T01:30:00')
Assert-Equal $nextDayRun.ToString('yyyy-MM-dd HH:mm') '2026-05-27 01:00' 'Next run time should roll to next day after schedule.'

Write-Host 'schedule.tests.ps1 passed' -ForegroundColor Green
