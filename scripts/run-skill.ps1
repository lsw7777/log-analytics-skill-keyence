param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Prompt,

    [Parameter(Mandatory = $false)]
    [string]$TableName = "",

    [Parameter(Mandatory = $false)]
    [switch]$ForceRefresh,

    [Parameter(Mandatory = $false)]
    [switch]$NoRiskFilter,

    [Parameter(Mandatory = $false)]
    [switch]$Open
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

$mainParams = @{
    Prompt = $Prompt
    SkipTotalCount = $true
}
if ($TableName) { $mainParams['TableName'] = $TableName }
if ($ForceRefresh) { $mainParams['ForceRefresh'] = $true }
if ($NoRiskFilter) { $mainParams['NoRiskFilter'] = $true }
if (-not $Open) { $mainParams['NoOpen'] = $true }

Write-Host "Skill prompt: $Prompt" -ForegroundColor Cyan
Write-Host 'Starting Log Analytics report from skill wrapper...' -ForegroundColor Cyan
& (Join-Path $ScriptDir 'main.ps1') @mainParams