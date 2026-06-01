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
@(
    [PSCustomObject]@{ accountEnabled = 'true'; businessPhones = ''; companyName = 'SZW-1'; department = 'SENSOR'; displayName = 'Shoyo Gao'; employeeId = '4480'; jobTitle = 'Staff'; mail = 'ShoyoGao@keyence.com.cn'; officeLocation = 'Shenzhen West'; userPrincipalName = 'C250105@china.keyence.com.cn'; TimeGenerated = '2026-05-26T22:02:56.8006694Z'; TenantId = '703a5771-97fc-4bf3-a585-f607d18c4479'; Type = 'AzureADUsersDCR_CL'; _ResourceId = '' }
    [PSCustomObject]@{ accountEnabled = 'false'; businessPhones = ''; companyName = 'PD'; department = 'FIGNA'; displayName = 'Ichinomiya Yoshinori'; employeeId = '1871'; jobTitle = 'Chuzai'; mail = 'Ichinomiya@keyence.com.cn'; officeLocation = ''; userPrincipalName = 'P207091@china.keyence.com.cn'; TimeGenerated = '2026-05-26T22:02:56.8006694Z'; TenantId = '703a5771-97fc-4bf3-a585-f607d18c4479'; Type = 'AzureADUsersDCR_CL'; _ResourceId = '' }
    [PSCustomObject]@{ accountEnabled = 'true'; businessPhones = ''; companyName = ''; department = ''; displayName = 'SHWH04'; employeeId = ''; jobTitle = ''; mail = 'SHWH04@china.keyence.com.cn'; officeLocation = ''; userPrincipalName = 'SHWH04@china.keyence.com.cn'; TimeGenerated = '2026-05-26T22:02:56.8006694Z'; TenantId = '703a5771-97fc-4bf3-a585-f607d18c4479'; Type = 'AzureADUsersDCR_CL'; _ResourceId = '' }
) | Export-Csv -Path $csvPath -Encoding UTF8 -NoTypeInformation -Force

& (Join-Path $rootDir 'analyze.ps1') -CsvPath $csvPath -OutputPath $htmlPath -AnalysisDate '2026-05-26' -TableName 'AzureADUsersDCR_CL' | Out-Null
$html = Get-Content -Path $htmlPath -Raw -Encoding UTF8

Assert-Contains $html 'Enabled Account | Department: SENSOR' 'Azure AD users should derive operation from account status and department.'
Assert-Contains $html 'Disabled Account | Department: FIGNA' 'Azure AD users should expose disabled account grouping.'
Assert-Contains $html 'decodeURIComponent(' 'HTML should use safe JS-encoded JSON payloads.'
Assert-Contains $html 'clientIpEmptyAzureAD' 'Azure AD users should explain unavailable client IP data with localized text.'
Assert-Contains $html 'timelineNoteAzureAD' 'Azure AD users should explain snapshot timeline semantics with localized text.'
$naClientIpPayload = 'JSON.parse(' + [char]39 + '[{"name":"N/A"'
Assert-NotContains $html $naClientIpPayload 'Client IP chart should not rank N/A values.'
Assert-NotContains $html 'suspiciousCount: ,' 'Risk counts must always render valid JavaScript numbers.'
Assert-Contains $html 'ipVelocityCount: 0' 'Risk counts must render zero instead of an empty JavaScript value.'

$auditCsvPath = Join-Path $tempDir 'AuditGeneralDCR_CL_sample.csv'
$auditHtmlPath = Join-Path $tempDir 'AuditGeneralDCR_CL_sample.html'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T01:02:03Z'; UserUPN = 'C250105@china.keyence.com.cn'; UserId = ''; Activity = 'ExportReport'; Operation = 'ExportReport'; Workload = 'PowerBI'; ClientIP = '8.8.8.8'; IsSuccess = 'true' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T02:02:03Z'; UserUPN = 'P207091@china.keyence.com.cn'; UserId = ''; Activity = 'Search'; Operation = 'Search'; Workload = 'PowerBI'; ClientIP = '8.8.4.4'; IsSuccess = 'false' }
) | Export-Csv -Path $auditCsvPath -Encoding UTF8 -NoTypeInformation -Force

& (Join-Path $rootDir 'analyze.ps1') -CsvPath $auditCsvPath -OutputPath $auditHtmlPath -AnalysisDate '2026-05-26' -TableName 'AuditGeneralDCR_CL' | Out-Null
$auditHtml = Get-Content -Path $auditHtmlPath -Raw -Encoding UTF8

Assert-Contains $auditHtml 'Shoyo Gao (C250105@china.keyence.com.cn)' 'Detailed data should show display name first and email in parentheses.'
Assert-Contains $auditHtml 'Ichinomiya Yoshinori (P207091@china.keyence.com.cn)' 'Risk analysis should show display name first and email in parentheses.'
Assert-Contains $auditHtml 'Shoyo%20Gao%20(C250105%40china.keyence.com.cn)' 'Active users chart should show display name first and email in parentheses.'

$spCsvPath = Join-Path $tempDir 'SharePointAuditDCR_CL_sample.csv'
$spHtmlPath = Join-Path $tempDir 'SharePointAuditDCR_CL_sample.html'
@(
    [PSCustomObject]@{ Operation = "FileAccessed'Broken"; Workload = 'SharePoint'; ClientIP = '47.102.133.2'; TimeGenerated = '2026-05-26T10:13:02Z'; UserId = 'c190433@china.keyence.com.cn'; UserAgent = 'TestAgent'; Type = 'SharePointAuditDCR_CL'; _ResourceId = '' }
) | Export-Csv -Path $spCsvPath -Encoding UTF8 -NoTypeInformation -Force

& (Join-Path $rootDir 'analyze.ps1') -CsvPath $spCsvPath -OutputPath $spHtmlPath -AnalysisDate '2026-05-26' -TableName 'SharePointAuditDCR_CL' | Out-Null
$spHtml = Get-Content -Path $spHtmlPath -Raw -Encoding UTF8
Assert-Contains $spHtml 'decodeURIComponent(' 'SharePointAudit HTML should use safe JS-encoded JSON payloads.'
Assert-Contains $spHtml 'FileAccessed&#39;Broken' 'SharePoint sample should preserve apostrophe in visible HTML text.'

$licenseCsvPath = Join-Path $tempDir 'AssignedLicensesDCR_CL_sample.csv'
$licenseHtmlPath = Join-Path $tempDir 'AssignedLicensesDCR_CL_sample.html'
@(
    [PSCustomObject]@{ AppliesTo = 'User'; DisplayName = 'Wing Liao'; ProvisioningStatus = 'PendingInput'; ServicePlanId = 'c1ec4a95-1f05-45b3-a911-aa3fa01094f5'; ServicePlanName = 'INTUNE_A'; TimeGenerated = '2026-05-26T23:40:10.2750393Z'; UserPrincipalName = 'wing@china.keyence.com.cn'; TenantId = '703a5771-97fc-4bf3-a585-f607d18c4479'; Type = 'AssignedLicensesDCR_CL'; _ResourceId = '' }
    [PSCustomObject]@{ AppliesTo = 'User'; DisplayName = 'Simon Yan'; ProvisioningStatus = 'Success'; ServicePlanId = '70d33638-9c74-4d01-bfd3-562de28bd4ba'; ServicePlanName = 'BI_AZURE_P2'; TimeGenerated = '2026-05-26T23:40:10.2750393Z'; UserPrincipalName = 'simon@china.keyence.com.cn'; TenantId = '703a5771-97fc-4bf3-a585-f607d18c4479'; Type = 'AssignedLicensesDCR_CL'; _ResourceId = '' }
) | Export-Csv -Path $licenseCsvPath -Encoding UTF8 -NoTypeInformation -Force

& (Join-Path $rootDir 'analyze.ps1') -CsvPath $licenseCsvPath -OutputPath $licenseHtmlPath -AnalysisDate '2026-05-26' -TableName 'AssignedLicensesDCR_CL' | Out-Null
$licenseHtml = Get-Content -Path $licenseHtmlPath -Raw -Encoding UTF8
Assert-Contains $licenseHtml 'PendingInput | INTUNE_A' 'AssignedLicenses operations should use provisioning status and service plan.'
Assert-Contains $licenseHtml 'timelineNoteAssignedLicenses' 'AssignedLicenses timeline should explain snapshot ingestion time.'
Assert-Contains $licenseHtml 'clientIpEmptyAssignedLicenses' 'AssignedLicenses should explain unavailable client IP data.'
Assert-Contains $licenseHtml "renderSuccessRatio('success-ratio', 1, 1, 0)" 'AssignedLicenses should classify non-success provisioning status as failed.'
Assert-Contains $licenseHtml '"riskIndicators":"1 ' 'AssignedLicenses non-success provisioning records should count as risk.'

Write-Host 'report.tests.ps1 passed' -ForegroundColor Green
