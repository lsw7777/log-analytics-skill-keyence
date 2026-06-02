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

$azureCsv = Join-Path $tempDir 'AzureADUsersDCR_CL_sample.csv'
@(
    [PSCustomObject]@{ accountEnabled = 'true'; displayName = 'John Risk'; mail = 'john@example.com'; userPrincipalName = 'john@example.com'; department = 'IT'; disabledDateTime = ''; TimeGenerated = '2026-05-26T01:00:00Z' }
    [PSCustomObject]@{ accountEnabled = 'false'; displayName = 'Shared Owner'; mail = 'shared@example.com'; userPrincipalName = 'shared@example.com'; department = 'OPS'; disabledDateTime = '2026-05-20T02:00:00Z'; TimeGenerated = '2026-05-26T01:00:00Z' }
    [PSCustomObject]@{ accountEnabled = 'true'; displayName = ''; mail = ''; userPrincipalName = 'missing@example.com'; department = ''; TimeGenerated = '2026-05-26T01:00:00Z' }
) | Export-Csv -Path $azureCsv -Encoding UTF8 -NoTypeInformation -Force

$signinCsv = Join-Path $tempDir 'SigninLogs_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T13:30:00Z'; UserPrincipalName = 'john@example.com'; AppDisplayName = 'Unknown SaaS'; IPAddress = '8.8.8.8'; ResultType = '0'; ResultDescription = 'Success' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T13:40:00Z'; UserPrincipalName = 'john@example.com'; AppDisplayName = 'Unknown SaaS'; IPAddress = '47.102.133.2:443'; ResultType = '0'; ResultDescription = 'Success trusted IP' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T02:00:00Z'; UserPrincipalName = 'john@example.com'; AppDisplayName = 'Microsoft Office'; IPAddress = '8.8.4.4'; ResultType = '50074'; ResultDescription = 'MFA failed' }
) | Export-Csv -Path $signinCsv -Encoding UTF8 -NoTypeInformation -Force

$auditCsv = Join-Path $tempDir 'AuditLogs_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T03:00:00Z'; InitiatedByUserPrincipalName = 'john@example.com'; InitiatedBy = '{"user":{"userPrincipalName":"jsonblob@example.com"}}'; OperationName = 'Delete user'; TargetResources = 'target@example.com'; Result = 'success'; ResultReason = '' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T04:00:00Z'; InitiatedByUserPrincipalName = 'john@example.com'; InitiatedBy = '{"user":{"userPrincipalName":"jsonblob@example.com"}}'; OperationName = 'Search'; TargetResources = 'Directory'; Result = 'success'; ResultReason = '' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T05:00:00Z'; InitiatedByUserPrincipalName = 'john@example.com'; InitiatedBy = '{"user":{"userPrincipalName":"jsonblob@example.com"}}'; OperationName = 'Add app role assignment to service principal'; TargetResources = 'spn'; Result = 'success'; ResultReason = 'permission changed' }
) | Export-Csv -Path $auditCsv -Encoding UTF8 -NoTypeInformation -Force

$licenseCsv = Join-Path $tempDir 'AssignedLicensesDCR_CL_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = ''; SkuPartNumber = 'LICENSE_A'; ProvisioningStatus = 'Success'; TotalLicenses = '10'; UsedUsers = '2'; __RecordKind = 'LicenseSummary' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = ''; SkuPartNumber = 'LICENSE_B'; ProvisioningStatus = 'Success'; TotalLicenses = '5'; UsedUsers = '1'; __RecordKind = 'LicenseSummary' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = ''; SkuPartNumber = 'LICENSE_C'; ProvisioningStatus = 'Success'; TotalLicenses = '3'; UsedUsers = '1'; __RecordKind = 'LicenseSummary' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = ''; SkuPartNumber = 'LICENSE_D'; ProvisioningStatus = 'Success'; TotalLicenses = '2'; UsedUsers = '1'; __RecordKind = 'LicenseSummary' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = 'john@example.com'; SkuPartNumber = 'LICENSE_C'; ProvisioningStatus = 'PendingInput'; TotalLicenses = '3'; UsedUsers = ''; __RecordKind = 'LicenseFailure' }
) | Export-Csv -Path $licenseCsv -Encoding UTF8 -NoTypeInformation -Force

$mailboxCsv = Join-Path $tempDir 'MailboxStatisticsDCR_CL_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = 'john@example.com'; RecipientTypeDetails = 'UserMailbox'; AvailableSpaceGB = '4'; QuotaLimitGB = '100'; TotalItemSizeGB = '96' }
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; UserPrincipalName = 'shared@example.com'; RecipientTypeDetails = 'SharedMailbox'; AvailableSpaceGB = '20'; QuotaLimitGB = '100'; TotalItemSizeGB = '80' }
) | Export-Csv -Path $mailboxCsv -Encoding UTF8 -NoTypeInformation -Force

$messageCsv = Join-Path $tempDir 'MessageTraceDataDCR_CL_sample.csv'
@(
    [PSCustomObject]@{ TimeGenerated = '2026-05-26T00:00:00Z'; SenderAddress = 'pbi@example.com'; RecipientAddress = 'john@example.com'; Status = 'Failed'; Subject = 'PBI refresh failed'; FromIP = '1.2.3.4' }
) | Export-Csv -Path $messageCsv -Encoding UTF8 -NoTypeInformation -Force

$htmlPath = Join-Path $tempDir 'merged_report.html'
& (Join-Path $rootDir 'analyze.ps1') `
    -CsvPath @($azureCsv, $signinCsv, $auditCsv, $licenseCsv, $mailboxCsv, $messageCsv) `
    -TableName @('AzureADUsersDCR_CL', 'SigninLogs', 'AuditLogs', 'AssignedLicensesDCR_CL', 'MailboxStatisticsDCR_CL', 'MessageTraceDataDCR_CL') `
    -OutputPath $htmlPath `
    -AnalysisDate '2026-05-20 to 2026-05-26' | Out-Null

$html = Get-Content -Path $htmlPath -Raw -Encoding UTF8

Assert-Contains $html 'Log Analytics 合并风险报告' 'Merged report title should be present.'
Assert-Contains $html '非工作时间范围: 21:00 - 08:00' 'Off-hours range should be displayed.'
Assert-Contains $html 'John Risk (john@example.com)' 'Other tables should join AzureADUsers displayName.'
Assert-NotContains $html 'jsonblob@example.com' 'AuditLogs should not use dynamic InitiatedBy JSON when scalar actor fields are present.'
Assert-Contains $html 'Unknown SaaS' 'Suspicious SigninLogs app should be shown.'
Assert-Contains $html '8.8.8.8' 'Untrusted suspicious sign-in IP should be shown.'
Assert-NotContains $html '47.102.133.2:443' 'Trusted IP with port should be normalized and not appear as suspicious raw value.'
Assert-Contains $html 'Delete user' 'AuditLogs delete operation should be shown.'
Assert-NotContains $html 'Search</td><td>Directory' 'AuditLogs Search should not be listed as delete/disable.'
Assert-Contains $html 'LICENSE_A' 'License names should be inferred from data.'
Assert-Contains $html '<td>LICENSE_A</td><td>2</td><td>10</td><td>8</td>' 'License remaining count should be calculated when total exists.'
Assert-Contains $html 'AvailableSpaceGB' 'Mailbox low available space risk should be explained.'
Assert-Contains $html 'Shared Owner (shared@example.com)' 'SharedMailbox should join disabled owner display name.'
Assert-Contains $html '2026-05-20 10:00:00' 'SharedMailbox disabled time should be shown in local time.'
Assert-Contains $html 'PBI refresh failed' 'MessageTrace PBI risk should be shown.'
Assert-Contains $html 'AzureADUsersDCR_CL' 'AzureADUsers should be reported as join source.'
Assert-Contains $html 'Managed Identity / SP' 'Identity permission change section should be present.'
Assert-Contains $html 'Add app role assignment to service principal' 'SP permission audit event should be shown.'
Assert-NotContains $html 'timeline-chart' 'Removed timeline section should not be present.'
Assert-NotContains $html 'donut-chart' 'Removed workload section should not be present.'
Assert-NotContains $html 'users-chart' 'Removed top users section should not be present.'
Assert-NotContains $html 'ops-chart' 'Removed top operations section should not be present.'
Assert-NotContains $html 'group-breakdown' 'Removed group section should not be present.'
Assert-NotContains $html 'table-body' 'Removed detailed data section should not be present.'
Assert-NotContains $html '<td>8.8.8.8</td><td>1</td>' 'Client IP ranking should not include suspicious IPs.'

Write-Host 'report.tests.ps1 passed' -ForegroundColor Green
