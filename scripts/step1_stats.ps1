# Step 1: Compute all statistics from CSV and save as JSON
$jsonPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\stats_20260508.json"
$csvPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\General_20260508.csv"

Write-Host "Loading CSV..." -ForegroundColor Cyan
$data = Import-Csv -Path $csvPath -Encoding UTF8
Write-Host "Loaded $($data.Count) records" -ForegroundColor Green

# Core metrics
$allUsers = @()
foreach ($row in $data) {
    $u = if ($row.UserUPN) { $row.UserUPN } elseif ($row.UserId) { $row.UserId } else { 'Unknown' }
    $allUsers += $u
}
$uniqueUsers = ($allUsers | Select-Object -Unique).Count

$allOps = @($data | ForEach-Object { if ($_.Operation) { $_.Operation } else { 'Unknown' } })
$uniqueOps = ($allOps | Select-Object -Unique).Count

$workloadMap = @{}
foreach ($row in $data) {
    $wl = if ($row.Workload) { $row.Workload } else { 'Unknown' }
    $workloadMap[$wl] = ($workloadMap[$wl] + 1)
}

$userMap = @{}
foreach ($u in $allUsers) { $userMap[$u] = ($userMap[$u] + 1) }
$topUsers = $userMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 15

$opMap = @{}
foreach ($o in $allOps) { $opMap[$o] = ($opMap[$o] + 1) }
$topOps = $opMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 15

$ipMap = @{}
foreach ($row in $data) {
    $ip = if ($row.ClientIP) { $row.ClientIP } else { 'Unknown' }
    $ipMap[$ip] = ($ipMap[$ip] + 1)
}
$topIPs = $ipMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10

$successCount = 0; $failCount = 0; $unknownCount = 0
foreach ($row in $data) {
    if ($row.IsSuccess -eq 'true') { $successCount++ }
    elseif ($row.IsSuccess -eq 'false') { $failCount++ }
    else { $unknownCount++ }
}

# Timeline
$hourMap = @{}
foreach ($row in $data) {
    $tg = $row.TimeGenerated
    if ($tg) {
        try {
            $dt = [DateTime]::Parse($tg)
            $h = $dt.ToString('yyyy-MM-dd HH:00')
            $hourMap[$h] = ($hourMap[$h] + 1)
        } catch {}
    }
}
$timelineSorted = $hourMap.GetEnumerator() | Sort-Object Name

# Failed ops
$failedEvents = @($data | Where-Object { $_.IsSuccess -eq 'false' })
$failedByOp = @{}
foreach ($row in $failedEvents) {
    $op = if ($row.Operation) { $row.Operation } else { 'Unknown' }
    $failedByOp[$op] = ($failedByOp[$op] + 1)
}
$failedByOpSorted = $failedByOp.GetEnumerator() | Sort-Object Value -Descending

# High-priv
$highPrivByUser = @{}
foreach ($row in $data) {
    if ($row.Operation -match '(?i)(^|[^a-z])(delete|deleted|remove|removed|disable|disabled|deactivate|deactivated)([^a-z]|$)') {
        $u = if ($row.UserUPN) { $row.UserUPN } elseif ($row.UserId) { $row.UserId } else { 'Unknown' }
        $key = "$u | $($row.Operation)"
        $highPrivByUser[$key] = ($highPrivByUser[$key] + 1)
    }
}
$highPrivSorted = $highPrivByUser.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 20

# Sensitive
$sensitiveOpsList = @('SensitivityLabeledFileOpened', 'SensitivityLabeledFileRenamed', 'IrmContent', 'AppliedSensitivityLabel', 'ChangedSensitivityLabel')
$sensitiveByOp = @{}
foreach ($row in $data) {
    if ($sensitiveOpsList -contains $row.Operation) {
        $sensitiveByOp[$row.Operation] = ($sensitiveByOp[$row.Operation] + 1)
    }
}
$sensitiveSorted = $sensitiveByOp.GetEnumerator() | Sort-Object Value -Descending

# Service accounts
$serviceAcctByOp = @{}
foreach ($row in $data) {
    if ($row.UserId -match '^00000009-') {
        $op = if ($row.Operation) { $row.Operation } else { 'Unknown' }
        $serviceAcctByOp[$op] = ($serviceAcctByOp[$op] + 1)
    }
}
$serviceAcctSorted = $serviceAcctByOp.GetEnumerator() | Sort-Object Value -Descending

# IP velocity
$ipUsers = @{}
foreach ($row in $data) {
    $ip = if ($row.ClientIP) { $row.ClientIP } else { continue }
    $u = if ($row.UserUPN) { $row.UserUPN } elseif ($row.UserId) { $row.UserId } else { 'Unknown' }
    if (-not $ipUsers.ContainsKey($ip)) { $ipUsers[$ip] = @{} }
    $ipUsers[$ip][$u] = 1
}
$ipVelocityList = @()
foreach ($ip in $ipUsers.Keys) {
    if ($ipUsers[$ip].Count -gt 5) {
        $ipVelocityList += [PSCustomObject]@{
            IP = $ip; UserCount = $ipUsers[$ip].Count
            Users = ($ipUsers[$ip].Keys -join ', ')
        }
    }
}
$ipVelocityList = $ipVelocityList | Sort-Object UserCount -Descending

# Suspicious IPs
$ipWorkloads = @{}
foreach ($row in $data) {
    $ip = if ($row.ClientIP) { $row.ClientIP } else { continue }
    if ($ip -match '^(10\.|172\.(1[6-9]|2[0-9]|3[01])\.|192\.168\.)' -or $ip -eq 'Unknown') { continue }
    $wl = if ($row.Workload) { $row.Workload } else { 'Unknown' }
    if (-not $ipWorkloads.ContainsKey($ip)) { $ipWorkloads[$ip] = @{} }
    $ipWorkloads[$ip][$wl] = 1
}
$suspiciousIPsList = @()
foreach ($ip in $ipWorkloads.Keys) {
    if ($ipWorkloads[$ip].Count -gt 1) {
        $suspiciousIPsList += [PSCustomObject]@{
            IP = $ip; WorkloadCount = $ipWorkloads[$ip].Count
            Workloads = ($ipWorkloads[$ip].Keys -join ', ')
        }
    }
}
$suspiciousIPsList = $suspiciousIPsList | Sort-Object WorkloadCount -Descending

# Operations present in data
$opNamesInData = $opMap.Keys | Sort-Object

# Unique workloads
$wlNames = $workloadMap.Keys | Sort-Object

# Build stats object
$stats = @{
    totalEvents = $data.Count
    uniqueUsers = $uniqueUsers
    uniqueOps = $uniqueOps
    workloadCount = $workloadMap.Count
    successCount = $successCount
    failCount = $failCount
    unknownCount = $unknownCount
    topUsers = ($topUsers | ForEach-Object { @{ name = $_.Name; value = $_.Value } })
    topOps = ($topOps | ForEach-Object { @{ name = $_.Name; value = $_.Value } })
    topIPs = ($topIPs | ForEach-Object { @{ name = $_.Name; value = $_.Value } })
    timeline = ($timelineSorted | ForEach-Object { @{ key = $_.Name; value = $_.Value } })
    workloads = ($workloadMap.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { @{ name = $_.Name; value = $_.Value } })
    failedOps = ($failedByOpSorted | ForEach-Object { @{ name = $_.Name; value = $_.Value } })
    highPriv = ($highPrivSorted | ForEach-Object { @{ name = $_.Name; value = $_.Value } })
    sensitive = ($sensitiveSorted | ForEach-Object { @{ name = $_.Name; value = $_.Value } })
    serviceAcct = ($serviceAcctSorted | ForEach-Object { @{ name = $_.Name; value = $_.Value } })
    suspiciousIPs = ($suspiciousIPsList | ForEach-Object { @{ key = $_.IP; value = $_.Workloads } })
    ipVelocity = ($ipVelocityList | ForEach-Object { @{ key = $_.IP; value = $_.UserCount } })
    opNames = @($opNamesInData)
    wlNames = @($wlNames)
    riskCount = $(
        $c = 0
        if ($failCount -gt 0) { $c++ }
        if ($suspiciousIPsList.Count -gt 0) { $c++ }
        if ($highPrivByUser.Count -gt 0) { $c++ }
        if ($sensitiveByOp.Count -gt 0) { $c++ }
        if ($ipVelocityList.Count -gt 0) { $c++ }
        $c
    )
}

$stats | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonPath -Encoding UTF8
Write-Host "Stats saved to $jsonPath" -ForegroundColor Green
