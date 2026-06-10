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

$managedCsv = Join-Path $tempDir 'AADManagedIdentitySignInLogs_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T01:00:00Z'; FirstTime = '2026-05-26T01:00:00Z'; LastTime = '2026-05-26T01:00:00Z'; EventCount = '1'; UserPrincipalName = 'mi-backup'; ServicePrincipalName = 'mi-backup'; AppDisplayName = 'mi-backup'; OperationName = 'mi-backup'; IPAddress = '8.8.4.4'; ResultType = '0'; ResultDescription = 'Success'; __RecordKind = 'AggregatedSuspiciousSigninSuccess' }
) | Export-Csv -Path $managedCsv -Encoding UTF8 -NoTypeInformation -Force

$spCsv = Join-Path $tempDir 'AADServicePrincipalSignInLogs_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T02:00:00Z'; FirstTime = '2026-05-26T01:00:00Z'; LastTime = '2026-05-26T02:00:00Z'; EventCount = '11'; UserPrincipalName = 'sp-risk'; ServicePrincipalName = 'sp-risk'; ResourceDisplayName = 'resource-api'; AppDisplayName = 'sp-risk'; OperationName = 'sp-risk'; IPAddress = '8.8.8.8'; ResultType = '500011'; ResultDescription = 'Failed'; __RecordKind = 'AggregatedFailedSignin' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T02:10:00Z'; FirstTime = '2026-05-26T02:00:00Z'; LastTime = '2026-05-26T02:10:00Z'; EventCount = '5'; UserPrincipalName = 'sp-low'; ServicePrincipalName = 'sp-low'; ResourceDisplayName = 'resource-low'; AppDisplayName = 'sp-low'; OperationName = 'sp-low'; IPAddress = '9.9.9.9'; ResultType = '500011'; ResultDescription = 'Failed'; __RecordKind = 'AggregatedFailedSignin' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T03:00:00Z'; FirstTime = '2026-05-26T03:00:00Z'; LastTime = '2026-05-26T03:00:00Z'; EventCount = '1'; UserPrincipalName = 'sp-success'; ServicePrincipalName = 'sp-success'; ResourceDisplayName = 'resource-success'; AppDisplayName = 'sp-success'; OperationName = 'sp-success'; IPAddress = '11.11.11.11'; ResultType = '0'; ResultDescription = 'Success'; __RecordKind = 'AggregatedSuspiciousSigninSuccess' }
) | Export-Csv -Path $spCsv -Encoding UTF8 -NoTypeInformation -Force

$signinCsv = Join-Path $tempDir 'SigninLogs_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T13:35:00Z'; FirstTime = '2026-05-26T13:30:00Z'; LastTime = '2026-05-26T13:35:00Z'; EventCount = '2'; UserPrincipalName = 'john@example.com'; AppDisplayName = 'Unknown SaaS'; IPAddress = '8.8.8.8'; ResultType = '0'; ResultDescription = 'Success'; __RecordKind = 'AggregatedSuspiciousSigninSuccess' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T13:40:00Z'; UserPrincipalName = 'john@example.com'; AppDisplayName = 'Unknown SaaS'; IPAddress = '47.102.133.2:443'; ResultType = '0'; ResultDescription = 'Success trusted IP' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T02:10:00Z'; FirstTime = '2026-05-26T02:00:00Z'; LastTime = '2026-05-26T02:10:00Z'; EventCount = '2'; UserPrincipalName = 'john@example.com'; AppDisplayName = 'Microsoft Office'; IPAddress = '8.8.4.4'; ResultType = '50074'; ResultDescription = 'MFA failed'; __RecordKind = 'AggregatedFailedSignin' }
) | Export-Csv -Path $signinCsv -Encoding UTF8 -NoTypeInformation -Force

$auditCsv = Join-Path $tempDir 'AuditLogs_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T03:05:00Z'; ActivityDateTime = '2026-05-26T03:05:00Z'; FirstTime = '2026-05-26T03:00:00Z'; LastTime = '2026-05-26T03:05:00Z'; EventCount = '2'; Actor = 'Kathy Cao / C250126@china.keyence.com.cn'; InitiatedBy = '{"user":{"userPrincipalName":"C250126@china.keyence.com.cn","displayName":"Kathy Cao"}}'; OperationName = 'Add service principal'; Target = 'sp-target'; PermissionName = ''; Result = 'success'; ResultReason = ''; __RecordKind = 'AggregatedServicePrincipalAudit' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T05:10:00Z'; ActivityDateTime = '2026-05-26T05:10:00Z'; FirstTime = '2026-05-26T05:00:00Z'; LastTime = '2026-05-26T05:10:00Z'; EventCount = '2'; Actor = 'Kathy Cao / C250126@china.keyence.com.cn'; InitiatedBy = '{"user":{"userPrincipalName":"C250126@china.keyence.com.cn","displayName":"Kathy Cao"}}'; OperationName = 'Add app role assignment to service principal'; Target = 'sp-target'; PermissionName = 'Mail.Read'; Result = 'success'; ResultReason = ''; __RecordKind = 'AggregatedServicePrincipalAudit' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T06:10:00Z'; Actor = 'Kathy Cao / C250126@china.keyence.com.cn'; InitiatedBy = '{"user":{"userPrincipalName":"C250126@china.keyence.com.cn","displayName":"Kathy Cao"}}'; OperationName = 'Remove member from role'; Target = 'role-target'; PermissionName = ''; Result = 'success'; ResultReason = 'PIM activation expired'; __RecordKind = 'AggregatedDeleteDisable' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T07:10:00Z'; Actor = 'app-only'; InitiatedBy = '{"app":{"appId":"00000000-0000-0000-0000-000000000001"}}'; OperationName = 'Hard delete service principal'; Target = 'app-target'; PermissionName = ''; Result = 'success'; ResultReason = ''; __RecordKind = 'AggregatedServicePrincipalAudit' }
) | Export-Csv -Path $auditCsv -Encoding UTF8 -NoTypeInformation -Force

$licenseCsv = Join-Path $tempDir 'AssignedLicensesDCR_CL_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = ''; SkuPartNumber = 'LICENSE_A'; ProvisioningStatus = 'Success'; TotalLicenses = '10'; UsedUsers = '2'; __RecordKind = 'LicenseSummary' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = ''; SkuPartNumber = 'LICENSE_B'; ProvisioningStatus = 'Success'; TotalLicenses = '5'; UsedUsers = '1'; __RecordKind = 'LicenseSummary' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = ''; SkuPartNumber = 'LICENSE_C'; ProvisioningStatus = 'Success'; TotalLicenses = '3'; UsedUsers = '1'; __RecordKind = 'LicenseSummary' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = ''; SkuPartNumber = 'LICENSE_D'; ProvisioningStatus = 'Success'; TotalLicenses = '2'; UsedUsers = '1'; __RecordKind = 'LicenseSummary' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = 'john@example.com'; SkuPartNumber = 'LICENSE_C'; ProvisioningStatus = 'PendingInput'; TotalLicenses = '3'; UsedUsers = ''; __RecordKind = 'LicenseFailure' }
) | Export-Csv -Path $licenseCsv -Encoding UTF8 -NoTypeInformation -Force

$dcrCsv = Join-Path $tempDir 'DCRLogErrors_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; FirstTime = '2026-05-20T00:00:00Z'; LastTime = '2026-05-26T00:00:00Z'; EventCount = '4'; InputStreamId = 'Custom-Stream'; OperationName = 'TransformKql'; Message = 'Failed to transform record'; Status = 'Failed'; __RecordKind = 'AggregatedDcrLogError' }
) | Export-Csv -Path $dcrCsv -Encoding UTF8 -NoTypeInformation -Force

$intuneCsv = Join-Path $tempDir 'IntuneAuditLogsDCR_CL_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; FirstTime = '2026-05-26T00:00:00Z'; LastTime = '2026-05-26T00:00:00Z'; EventCount = '3'; Actor = 'admin@example.com'; OperationName = 'Delete device configuration'; TargetDisplayName = 'Windows policy'; Result = 'failed'; ResultDescription = 'Policy delete failed'; __RecordKind = 'AggregatedIntuneAuditRisk' }
) | Export-Csv -Path $intuneCsv -Encoding UTF8 -NoTypeInformation -Force

$mailboxCsv = Join-Path $tempDir 'MailboxStatisticsDCR_CL_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = 'john@example.com'; RecipientTypeDetails = 'UserMailbox'; AvailableSpaceGB = '4'; QuotaLimitGB = '100'; TotalItemSizeGB = '96' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = 'shared@example.com'; RecipientTypeDetails = 'SharedMailbox'; AvailableSpaceGB = '20'; QuotaLimitGB = '100'; TotalItemSizeGB = '80' }
) | Export-Csv -Path $mailboxCsv -Encoding UTF8 -NoTypeInformation -Force

$htmlPath = Join-Path $tempDir 'merged_report.html'
& (Join-Path $rootDir 'analyze.ps1') `
    -CsvPath @($managedCsv, $spCsv, $signinCsv, $auditCsv, $licenseCsv, $dcrCsv, $intuneCsv, $mailboxCsv) `
    -TableName @('AADManagedIdentitySignInLogs', 'AADServicePrincipalSignInLogs', 'SigninLogs', 'AuditLogs', 'AssignedLicensesDCR_CL', 'DCRLogErrors', 'IntuneAuditLogsDCR_CL', 'MailboxStatisticsDCR_CL') `
    -OutputPath $htmlPath `
    -AnalysisDate '2026-05-20 to 2026-05-26' | Out-Null

$html = Get-Content -Path $htmlPath -Raw -Encoding UTF8

Assert-Contains $html 'Log Analytics 合并风险报告' 'Merged report title should be present.'
Assert-Contains $html '<nav class="side-nav">' 'Report should include side navigation.'
Assert-Contains $html '<details class="section" id="failed-signins" open>' 'Sections should be collapsible details elements.'
Assert-NotContains $html '非工作时间' 'Off-hours section should be removed.'
Assert-Contains $html 'sp-risk' 'ServicePrincipalName should be shown for SP failures.'
Assert-NotContains $html 'sp-low' 'SP failures with 10 or fewer events should be hidden.'
Assert-NotContains $html 'resource-low' 'Low-count SP failure resource should be hidden.'
Assert-Contains $html 'sp-success' 'SP suspicious successful sign-in should be shown.'
Assert-Contains $html 'mi-backup' 'Managed Identity suspicious successful sign-in should be shown.'
Assert-Contains $html 'Unknown SaaS' 'Suspicious SigninLogs app should be shown.'
Assert-Contains $html '8.8.8.8' 'Untrusted suspicious SigninLogs IP should be shown.'
Assert-NotContains $html '47.102.133.2:443' 'Trusted IP with port should be normalized and not appear as suspicious raw value.'
Assert-Contains $html 'Kathy Cao / C250126@china.keyence.com.cn' 'AuditLogs should show user actor.'
Assert-Contains $html 'Add service principal' 'SP object audit operation should be shown.'
Assert-Contains $html 'Add app role assignment to service principal' 'SP app role audit event should be shown.'
Assert-Contains $html 'Mail.Read' 'SP app role permission name should be shown.'
Assert-NotContains $html 'PIM activation expired' 'PIM activation expiry should be filtered.'
Assert-NotContains $html 'app-target' 'AuditLogs app-only actor should be filtered.'
Assert-Contains $html 'LICENSE_A' 'License names should be inferred from data.'
Assert-Contains $html '<td>LICENSE_A</td><td>2</td><td>10</td><td>8</td><td>Log</td>' 'License remaining count should be calculated when total exists.'
Assert-Contains $html 'DCRLogErrors' 'DCRLogErrors section should be present.'
Assert-Contains $html 'Custom-Stream' 'DCRLogErrors InputStreamId should be shown.'
Assert-Contains $html 'Failed to transform record' 'DCRLogErrors Message should be shown.'
Assert-Contains $html 'Intune 审计风险' 'Intune audit section should be present.'
Assert-Contains $html 'admin@example.com' 'Intune Actor should be shown.'
Assert-Contains $html 'Windows policy' 'Intune Target should be shown.'
Assert-Contains $html 'AvailableSpaceGB' 'Mailbox low available space risk should be explained.'
Assert-Contains $html 'shared@example.com' 'SharedMailbox should be shown.'
Assert-Contains $html '<td>shared@example.com</td><td>SharedMailbox</td><td>80 GB</td><td>20</td><td>100</td>' 'SharedMailbox size and quota data should be shown in GB.'
Assert-NotContains $html 'AzureADUsersDCR_CL' 'AzureADUsersDCR_CL should not be queried or reported.'
Assert-NotContains $html 'MessageTraceDataDCR_CL' 'MessageTraceDataDCR_CL should not be queried or reported.'
Assert-NotContains $html 'SharePointAuditDCR_CL' 'SharePointAuditDCR_CL should not be queried or reported.'
Assert-NotContains $html 'AuditGeneralDCR_CL' 'AuditGeneralDCR_CL should not be queried or reported.'
Assert-NotContains $html 'timeline-chart' 'Removed timeline section should not be present.'
Assert-NotContains $html 'donut-chart' 'Removed workload section should not be present.'
Assert-NotContains $html 'users-chart' 'Removed top users section should not be present.'
Assert-NotContains $html 'ops-chart' 'Removed top operations section should not be present.'
Assert-NotContains $html 'group-breakdown' 'Removed group section should not be present.'
Assert-NotContains $html 'table-body' 'Removed detailed data section should not be present.'
Assert-NotContains $html '来源表/工作负载：AADServicePrincipalSignInLogs' 'Suspicious IP section should not include AADServicePrincipalSignInLogs IPs.'

Write-Host 'report.tests.ps1 passed' -ForegroundColor Green
