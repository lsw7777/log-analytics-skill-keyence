$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir = Split-Path -Parent $scriptDir
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) 'log-analyzer-tests'
if (-not (Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
}

function Assert-Contains {
    param([string]$Content, [string]$Expected, [string]$Message)
    if (-not $Content.Contains($Expected)) {
        throw "$Message Expected content to include '$Expected'."
    }
}

function Assert-NotContains {
    param([string]$Content, [string]$Unexpected, [string]$Message)
    if ($Content.Contains($Unexpected)) {
        throw "$Message Content should not include '$Unexpected'."
    }
}

$csvPath = Join-Path $tempDir 'AzureADUsersDCR_CL_sample.csv'
$htmlPath = Join-Path $tempDir 'AzureADUsersDCR_CL_sample.html'

@'
"accountEnabled","businessPhones","companyName","department","displayName","employeeId","jobTitle","mail","officeLocation","userPrincipalName","TimeGenerated","TenantId","Type","_ResourceId"
"true","","SZW-1","SENSOR","Shoyo Gao","4480","Staff","ShoyoGao@keyence.com.cn","Shenzhen West","C250105@china.keyence.com.cn","2026-05-26T22:02:56.8006694Z","703a5771-97fc-4bf3-a585-f607d18c4479","AzureADUsersDCR_CL",""
"false","","PD","FIGNA","Ichinomiya Yoshinori","1871","Chuzai","Ichinomiya@keyence.com.cn","","P207091@china.keyence.com.cn","2026-05-26T22:02:56.8006694Z","703a5771-97fc-4bf3-a585-f607d18c4479","AzureADUsersDCR_CL",""
"true","","","","SHWH04","","","SHWH04@china.keyence.com.cn","","SHWH04@china.keyence.com.cn","2026-05-26T22:02:56.8006694Z","703a5771-97fc-4bf3-a585-f607d18c4479","AzureADUsersDCR_CL",""
'@ | Out-File -FilePath $csvPath -Encoding UTF8 -Force

& (Join-Path $rootDir 'analyze.ps1') -CsvPath $csvPath -OutputPath $htmlPath -AnalysisDate '2026-05-26' -TableName 'AzureADUsersDCR_CL' | Out-Null
$html = Get-Content -Path $htmlPath -Raw -Encoding UTF8

Assert-Contains $html 'Enabled Account | Department: SENSOR' 'Azure AD users should derive operation from account status and department.'
Assert-Contains $html 'Disabled Account | Department: FIGNA' 'Azure AD users should expose disabled account grouping.'
Assert-Contains $html 'accountEnabled=true' 'Azure AD glossary should explain accountEnabled=true.'
Assert-Contains $html 'This table does not include client IP fields.' 'Azure AD users should explain unavailable client IP data.'
Assert-Contains $html 'AzureADUsersDCR_CL is a directory snapshot table; TimeGenerated is ingestion time, not user activity time.' 'Azure AD users should explain snapshot timeline semantics.'
Assert-NotContains $html 'JSON.parse(''[{"name":"N/A"' 'Client IP chart should not rank N/A values.'
Assert-NotContains $html 'suspiciousCount: ,' 'Risk counts must always render valid JavaScript numbers.'
Assert-Contains $html 'ipVelocityCount: 0' 'Risk counts must render zero instead of an empty JavaScript value.'

$auditCsvPath = Join-Path $tempDir 'AuditGeneralDCR_CL_sample.csv'
$auditHtmlPath = Join-Path $tempDir 'AuditGeneralDCR_CL_sample.html'

@'
"TimeGenerated","UserUPN","UserId","Activity","Operation","Workload","ClientIP","IsSuccess"
"2026-05-26T01:02:03Z","C250105@china.keyence.com.cn","","ExportReport","ExportReport","PowerBI","8.8.8.8","true"
"2026-05-26T02:02:03Z","P207091@china.keyence.com.cn","","Search","Search","PowerBI","8.8.4.4","false"
'@ | Out-File -FilePath $auditCsvPath -Encoding UTF8 -Force

& (Join-Path $rootDir 'analyze.ps1') -CsvPath $auditCsvPath -OutputPath $auditHtmlPath -AnalysisDate '2026-05-26' -TableName 'AuditGeneralDCR_CL' | Out-Null
$auditHtml = Get-Content -Path $auditHtmlPath -Raw -Encoding UTF8

Assert-Contains $auditHtml 'Shoyo Gao (C250105@china.keyence.com.cn)' 'Detailed data should show display name first and email in parentheses.'
Assert-Contains $auditHtml 'Ichinomiya Yoshinori (P207091@china.keyence.com.cn)' 'Risk analysis should show display name first and email in parentheses.'
Assert-Contains $auditHtml '"name":"Shoyo Gao (C250105@china.keyence.com.cn)"' 'Active users chart should show display name first and email in parentheses.'

Write-Host 'report.tests.ps1 passed' -ForegroundColor Green
