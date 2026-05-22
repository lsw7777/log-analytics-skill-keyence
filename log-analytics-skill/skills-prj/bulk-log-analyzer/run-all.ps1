# ============================================================
# Log Analytics 一键执行脚本
# 功能: 获取日志 -> 统计分析 -> 生成HTML报告 -> 在浏览器中打开
# ============================================================

param(
    [Parameter(Mandatory = $false)]
    [string]$TableName = "AuditGeneralDCR_CL",

    [Parameter(Mandatory = $false)]
    [int]$Hours = 24,

    [Parameter(Mandatory = $false)]
    [switch]$ForceLogin
)

# ============================================================
# 配置
# ============================================================
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$TempDir = "$env:USERPROFILE\AppData\Local\Temp\opencode"
$DateStr = Get-Date -Format "yyyyMMdd_HHmm"
$CsvFile = "$TempDir\$($TableName.Substring(0, $TableName.IndexOf('_DCR')))_$DateStr.csv"
$HtmlFile = "$TempDir\report_$DateStr.html"

# 确保临时目录存在
if (-not (Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
}

Write-Host "============================================" -ForegroundColor Magenta
Write-Host "  Log Analytics 一键执行" -ForegroundColor Magenta
Write-Host "============================================" -ForegroundColor Magenta
Write-Host "表名: $TableName" -ForegroundColor Cyan
Write-Host "时间范围: 过去 $Hours 小时" -ForegroundColor Cyan
Write-Host "CSV输出: $CsvFile" -ForegroundColor Cyan
Write-Host "HTML输出: $HtmlFile" -ForegroundColor Cyan
Write-Host ""

# ============================================================
# 步骤 1: 获取日志数据
# ============================================================
Write-Host "[1/4] 正在获取日志数据..." -ForegroundColor Yellow

$QueryParams = @{
    TableName = $TableName
    Hours     = $Hours
    ExportCsv = $CsvFile
}

if ($ForceLogin) {
    $QueryParams['ForceLogin'] = $true
}

& "$ScriptDir\azure_log_query.ps1" @QueryParams

if (-not (Test-Path $CsvFile)) {
    Write-Host "错误: CSV文件未生成，查询可能失败" -ForegroundColor Red
    exit 1
}

Write-Host "数据获取完成!" -ForegroundColor Green
Write-Host ""

# ============================================================
# 步骤 2: 加载并统计数据
# ============================================================
Write-Host "[2/4] 正在加载并统计数据..." -ForegroundColor Yellow

$data = Import-Csv -Path $CsvFile -Encoding UTF8
$totalEvents = $data.Count
Write-Host "总记录数: $totalEvents" -ForegroundColor Green

if ($totalEvents -eq 0) {
    Write-Host "警告: 没有获取到数据" -ForegroundColor Yellow
    exit 0
}

# 计算统计信息
$allUsers = @()
foreach ($row in $data) {
    $u = if ($row.UserUPN -and $row.UserUPN -ne '') { $row.UserUPN } elseif ($row.UserId -and $row.UserId -ne '') { $row.UserId } else { 'Unknown' }
    $allUsers += $u
}
$uniqueUsers = ($allUsers | Select-Object -Unique).Count

$allOps = @($data | ForEach-Object { $_.Operation })
$uniqueOps = ($allOps | Select-Object -Unique).Count

$workloadMap = @{}
foreach ($row in $data) {
    $wl = if ($row.Workload) { $row.Workload } else { 'Unknown' }
    $workloadMap[$wl] = ($workloadMap[$wl] + 1)
}

$successCount = 0
$failCount = 0
$unknownCount = 0
foreach ($row in $data) {
    $s = $row.IsSuccess
    if ($s -eq 'true') { $successCount++ }
    elseif ($s -eq 'false') { $failCount++ }
    else { $unknownCount++ }
}

Write-Host "唯一用户: $uniqueUsers" -ForegroundColor Green
Write-Host "唯一操作: $uniqueOps" -ForegroundColor Green
Write-Host "工作负载: $($workloadMap.Count)" -ForegroundColor Green
Write-Host "成功: $successCount | 失败: $failCount | 未知: $unknownCount" -ForegroundColor Green
Write-Host ""

# ============================================================
# 步骤 3: 生成HTML报告
# ============================================================
Write-Host "[3/4] 正在生成HTML报告..." -ForegroundColor Yellow

# 调用 analyze.ps1 生成报告
& "$ScriptDir\analyze.ps1" -CsvPath $CsvFile -OutputPath $HtmlFile -AnalysisDate (Get-Date -Format "yyyy-MM-dd")

if (-not (Test-Path $HtmlFile)) {
    Write-Host "错误: HTML文件未生成" -ForegroundColor Red
    exit 1
}

Write-Host "HTML报告生成完成: $HtmlFile" -ForegroundColor Green
Write-Host ""

# ============================================================
# 步骤 4: 在浏览器中打开
# ============================================================
Write-Host "[4/4] 正在打开浏览器..." -ForegroundColor Yellow

# 将文件路径转换为URL格式
$htmlUrl = "file:///$($HtmlFile -replace '\\', '/')"
Write-Host "HTML网址: $htmlUrl" -ForegroundColor Cyan

# 在默认浏览器中打开
Start-Process $HtmlFile

Write-Host ""
Write-Host "============================================" -ForegroundColor Magenta
Write-Host "  完成!" -ForegroundColor Magenta
Write-Host "============================================" -ForegroundColor Magenta
Write-Host "CSV文件: $CsvFile" -ForegroundColor Cyan
Write-Host "HTML文件: $HtmlFile" -ForegroundColor Cyan
Write-Host "HTML网址: $htmlUrl" -ForegroundColor Cyan