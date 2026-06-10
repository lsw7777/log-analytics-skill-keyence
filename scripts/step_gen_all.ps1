$ErrorActionPreference = 'Stop'
$csvPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\General_20260508.csv"
$tmplPath = Join-Path $PSScriptRoot 'template\template_v3.html'
$outPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\report_20260508_allfields.html"

Write-Host "Loading..." -ForegroundColor Cyan
$tmpl = Get-Content -Path $tmplPath -Raw -Encoding UTF8
$data = Import-Csv -Path $csvPath -Encoding UTF8
$limit = [Math]::Min(500, $data.Count)
$fields = $data[0].PSObject.Properties.Name

Write-Host "Building rows..." -ForegroundColor Cyan
$jsonList = New-Object System.Collections.ArrayList
$htmlRows = ""
$q = [char]34

for ($i = 0; $i -lt $limit; $i++) {
    $r = $data[$i]
    $obj = @{}
    foreach ($f in $fields) {
        $obj[$f] = if ($r.$f) { $r.$f } else { "" }
    }
    $json = $obj | ConvertTo-Json -Compress -Depth 1
    [void]$jsonList.Add($json)
    
    $tg = [string]$r.TimeGenerated
    $op = [string]$r.Operation
    $rt = [string]$r.RecordType
    $act = [string]$r.Activity
    $oid = [string]$r.ObjectId
    $u = [string]$r.UserId
    $wl = [string]$r.Workload
    $ip = [string]$r.ClientIP
    $st = [string]$r.IsSuccess
    $stCls = ""
    if ($st -eq 'true') { $stCls = 'status-true' } elseif ($st -eq 'false') { $stCls = 'status-false' }
    
    $htmlRows += "<tr class='data-row' onclick='showDetail($i)'><td>$i</td><td>$tg</td><td>$op</td><td>$rt</td><td>$act</td><td>$oid</td><td>$u</td><td>$wl</td><td>$ip</td><td class='$stCls'>$st</td></tr>`n"
}

$jsonStr = "var ALL_DATA=[$($jsonList -join ',')];"

Write-Host "Injecting $limit rows..." -ForegroundColor Cyan
$html = $tmpl.Replace('<tbody id="table-body"></tbody>', "<tbody id=`"table-body`">$htmlRows</tbody>")
$html = $html.Replace('/* ALL_DATA_PLACEHOLDER */', $jsonStr)

$utf8BOM = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($outPath, $html, $utf8BOM)
Write-Host "Done!" -ForegroundColor Green
Get-Item $outPath | Select-Object Name, Length
