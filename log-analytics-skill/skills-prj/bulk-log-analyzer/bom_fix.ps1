$src = "C:\Users\d9347\AppData\Local\Temp\opencode\report_20260508.html"
$bytes = [System.IO.File]::ReadAllBytes($src)
$bom = [byte[]](0xEF, 0xBB, 0xBF)
$n = New-Object byte[] ($bytes.Length + 3)
[System.Array]::Copy($bom, 0, $n, 0, 3)
[System.Array]::Copy($bytes, 0, $n, 3, $bytes.Length)
[System.IO.File]::WriteAllBytes($src, $n)
Get-Item $src | Select-Object Name, Length
