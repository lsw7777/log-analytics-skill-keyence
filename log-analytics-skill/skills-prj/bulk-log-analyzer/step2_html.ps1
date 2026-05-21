# Step 2: Read template, inject table data, write final HTML with BOM
$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'

$csvPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\General_20260508.csv"
$tmplPath = "$env:USERPROFILE\Desktop\log-skill\skills-prj\bulk-log-analyzer\final_report.html"
$outPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\report_20260508.html"

Write-Host "Loading template..." -ForegroundColor Cyan
$tmplBytes = [System.IO.File]::ReadAllBytes($tmplPath)
$tmpl = [System.Text.Encoding]::UTF8.GetString($tmplBytes)

Write-Host "Loading CSV for table..." -ForegroundColor Cyan
$data = Import-Csv -Path $csvPath -Encoding UTF8
$limit = [Math]::Min(500, $data.Count)

$rows = ""
for ($i = 0; $i -lt $limit; $i++) {
    $r = $data[$i]
    $tg = if ($r.TimeGenerated) { $r.TimeGenerated } else { "" }
    $user = if ($r.UserUPN) { $r.UserUPN } elseif ($r.UserId) { $r.UserId } else { "" }
    $op = if ($r.Operation) { $r.Operation } else { "" }
    $wl = if ($r.Workload) { $r.Workload } else { "" }
    $ip = if ($r.ClientIP) { $r.ClientIP } else { "" }
    $st = if ($r.IsSuccess) { $r.IsSuccess } else { "" }
    $rows += "<tr><td>$i</td><td>$tg</td><td class='op-cell' data-op='$op'>$op</td><td>$user</td><td>$wl</td><td>$ip</td><td class='status-$st'>$st</td></tr>`n"
}

Write-Host "Injecting $limit rows into template..." -ForegroundColor Cyan
$html = $tmpl.Replace('<tbody id="table-body"></tbody>', "<tbody id=`"table-body`">$rows</tbody>")

# Write with UTF-8 BOM
$utf8BOM = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($outPath, $html, $utf8BOM)
Write-Host "Report saved to: $outPath" -ForegroundColor Green
Get-Item $outPath | Select-Object Name, Length
