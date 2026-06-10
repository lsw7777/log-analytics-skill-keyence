$ErrorActionPreference = 'Stop'
$csvPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\General_20260508.csv"
$tmplPath = Join-Path $PSScriptRoot 'template\final_report_v2.html'
$outPath = "$env:USERPROFILE\AppData\Local\Temp\opencode\report_20260508_v4.html"

Write-Host "Loading..." -ForegroundColor Cyan
$tmpl = Get-Content -Path $tmplPath -Raw -Encoding UTF8
$data = Import-Csv -Path $csvPath -Encoding UTF8
$limit = [Math]::Min(500, $data.Count)

$rowsHtml = ""
$rowsJsParts = @()
for ($i = 0; $i -lt $limit; $i++) {
    $r = $data[$i]
    $tg = if ($r.TimeGenerated) { $r.TimeGenerated } else { '' }
    $user = if ($r.UserUPN) { $r.UserUPN } elseif ($r.UserId) { $r.UserId } else { '' }
    $op = if ($r.Operation) { $r.Operation } else { '' }
    $wl = if ($r.Workload) { $r.Workload } else { '' }
    $ip = if ($r.ClientIP) { $r.ClientIP } else { '' }
    $st = if ($r.IsSuccess) { $r.IsSuccess } else { '' }

    $userE = $user -replace "'", "\'"
    $opE = $op -replace "'", "\'"
    $tgE = $tg -replace "'", "\'"

    $rowsHtml += "<tr><td>$i</td><td>$tg</td><td class='op-cell' data-op='$op'>$op</td><td>$user</td><td>$wl</td><td>$ip</td><td class='status-$st'>$st</td></tr>`n"
    $rowsJsParts += "{idx:$i,tg:'$tgE',user:'$userE',op:'$opE',wl:'$wl',ip:'$ip',st:'$st'}"
}

$rowsJs = $rowsJsParts -join ','
$q = '"'
$jsLine = 'var _raw=[' + $rowsJs + '];allTableRows=_raw.map(function(r){var c=' + $q + '<td>'+r.idx+'</td><td>'+r.tg+'</td><td class=' + $q + '+$q+'+ "op-cell" +'+$q+' data-op='+$q+'+r.op+'+$q+'>'+r.op+'</td><td>'+r.user+'</td><td>'+r.wl+'</td><td>'+r.ip+'</td><td class=' + $q + '+$q+'+ "status-" +'+$q+'+r.st+'+$q+'>'+r.st+'</td>';return{idx:r.idx,user:r.user,op:r.op,wl:r.wl,ip:r.ip,st:r.st,tg:r.tg,cells:c,attr:''};});'

Write-Host "Injecting $limit rows..." -ForegroundColor Cyan
$html = $tmpl.Replace('<tbody id="table-body"></tbody>', "<tbody id=`"table-body`">$($rowsHtml)</tbody>")
$html = $html.Replace('var allTableRows = [];', $jsLine)

$utf8BOM = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($outPath, $html, $utf8BOM)
Write-Host "Report saved to: $outPath" -ForegroundColor Green
Get-Item $outPath | Select-Object Name, Length
