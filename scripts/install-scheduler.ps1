param(
    [string]$RunAt = '01:00',
    [string[]]$Tables = @(),
    [switch]$CreateTrayShortcut
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptDir 'log-analyzer-core.ps1')

$config = New-LogAnalyzerScheduleConfig -RunAt $RunAt -Tables $Tables
$configPath = Get-LogAnalyzerScheduleConfigPath -RootDir $scriptDir
$statusPath = Get-LogAnalyzerStatusPath -RootDir $scriptDir

$config | ConvertTo-Json -Depth 5 | Out-File -FilePath $configPath -Encoding UTF8 -Force
$status = [PSCustomObject]@{
    RunAt = $config.RunAt
    Tables = $config.Tables
    ConfigPath = $configPath
    StatusPath = $statusPath
}
$status | ConvertTo-Json -Depth 5 | Out-File -FilePath $statusPath -Encoding UTF8 -Force

$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument (Get-LogAnalyzerBatchCommand -RootDir $scriptDir -ConfigPath $configPath)
$trigger = New-ScheduledTaskTrigger -Daily -At ([datetime]::Parse($config.RunAt))
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel LeastPrivilege
$taskName = 'BulkLogAnalyzerDaily'
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null

Write-Host "Scheduled task installed: $taskName" -ForegroundColor Green
Write-Host "Run time: $($config.RunAt)" -ForegroundColor Green
Write-Host "Tables: $($config.Tables -join ', ')" -ForegroundColor Green

if ($CreateTrayShortcut) {
    $wsh = New-Object -ComObject WScript.Shell
    $startup = [Environment]::GetFolderPath('Startup')
    $shortcutPath = Join-Path $startup 'BulkLogAnalyzerTray.lnk'
    $shortcut = $wsh.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = 'powershell.exe'
    $shortcut.Arguments = (Get-LogAnalyzerTrayCommand -RootDir $scriptDir -ConfigPath $configPath)
    $shortcut.WorkingDirectory = $scriptDir
    $shortcut.Save()
    Write-Host "Tray shortcut created: $shortcutPath" -ForegroundColor Green
}
