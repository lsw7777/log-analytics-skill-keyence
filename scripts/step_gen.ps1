$ErrorActionPreference = 'Stop'
$csvPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\General_20260508.csv"
$tmplPath = Join-Path $PSScriptRoot 'template\final_report_v2.html'
$outPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\report_20260508_v4.html"

Write-Host "Loading template..." -ForegroundColor Cyan
$tmpl = Get-Content -Path $tmplPath -Raw -Encoding UTF8
Write-Host "Loading CSV..." -ForegroundColor Cyan
$data = Import-Csv -Path $csvPath -Encoding UTF8
$limit = [Math]::Min(500, $data.Count)

$rowsHtml = ""
for ($i = 0; $i -lt $limit; $i++) {
    $r = $data[$i]
    $tg = ""
    if ($r.TimeGenerated) { $tg = $r.TimeGenerated }
    $user = ""
    if ($r.UserUPN) { $user = $r.UserUPN } elseif ($r.UserId) { $user = $r.UserId }
    $op = ""
    if ($r.Operation) { $op = $r.Operation }
    $wl = ""
    if ($r.Workload) { $wl = $r.Workload }
    $ip = ""
    if ($r.ClientIP) { $ip = $r.ClientIP }
    $st = ""
    if ($r.IsSuccess) { $st = $r.IsSuccess }

    $rowsHtml += "<tr><td>$i</td><td>$tg</td><td class='op-cell' data-op='$op'>$op</td><td>$user</td><td>$wl</td><td>$ip</td><td class='status-$st'>$st</td></tr>`n"
}

Write-Host "Injecting $limit rows..." -ForegroundColor Cyan
$html = $tmpl.Replace('<tbody id="table-body"></tbody>', "<tbody id=`"table-body`">$($rowsHtml)</tbody>")

$utf8BOM = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($outPath, $html, $utf8BOM)
Write-Host "Report saved to: $outPath" -ForegroundColor Green
Get-Item $outPath | Select-Object Name, Length
