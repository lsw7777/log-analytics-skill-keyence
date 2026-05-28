param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptDir 'log-analyzer-core.ps1')

$config = if (Test-Path $ConfigPath) { Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json } else { New-LogAnalyzerScheduleConfig }
$nextRun = Get-LogAnalyzerNextRunTime -RunAt $config.RunAt

$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon = [System.Drawing.SystemIcons]::Information
$notify.Visible = $true
$notify.Text = "Bulk Log Analyzer waiting until $($nextRun.ToString('HH:mm'))"

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$menu.Items.Add('Run now') | Out-Null
$menu.Items.Add('Exit') | Out-Null
$notify.ContextMenuStrip = $menu

[System.Windows.Forms.Application]::Run()
