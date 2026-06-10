# ============================================================
# Cache Manager for Log Analytics
# View, clean, and manage cached data
# ============================================================

param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("list", "clear", "stats", "clear-all")]
    [string]$Action = "list",

    [Parameter(Mandatory = $false)]
    [string]$TableName = "",

    [Parameter(Mandatory = $false)]
    [int]$OlderThanHours = 48
)

$CacheDir = "$env:USERPROFILE\AppData\Local\Temp\opencode\cache"

if (-not (Test-Path $CacheDir)) {
    Write-Host "Cache directory not found: $CacheDir" -ForegroundColor Yellow
    Write-Host "No cache exists yet." -ForegroundColor Yellow
    exit 0
}

# ============================================================
# List Cache
# ============================================================
function Show-CacheList {
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host "  Log Analytics Cache Manager" -ForegroundColor Magenta
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host ""

    $metaFiles = Get-ChildItem -Path $CacheDir -Filter "*.meta.json" -ErrorAction SilentlyContinue

    if ($metaFiles.Count -eq 0) {
        Write-Host "No cached data found." -ForegroundColor Yellow
        return
    }

    $cacheItems = @()
    foreach ($metaFile in $metaFiles) {
        try {
            $meta = Get-Content $metaFile.FullName -Raw | ConvertFrom-Json
            $cacheTime = [DateTime]::Parse($meta.CacheTime)
            $age = (Get-Date) - $cacheTime
            $ttlHours = if ($meta.CacheTTL) { $meta.CacheTTL } else { 24 }
            $isExpired = $age.TotalHours -gt $ttlHours

            $csvFile = $metaFile.FullName -replace '\.meta\.json$', '.csv'
            $csvSize = if (Test-Path $csvFile) {
                $size = (Get-Item $csvFile).Length
                if ($size -gt 1MB) { "{0:N1} MB" -f ($size / 1MB) }
                elseif ($size -gt 1KB) { "{0:N1} KB" -f ($size / 1KB) }
                else { "$size B" }
            } else { "N/A" }

            $cacheItems += [PSCustomObject]@{
                Table = $meta.TableName
                Records = $meta.RecordCount
                Age = "$([Math]::Round($age.TotalMinutes, 0))min"
                TTL = "${ttlHours}h"
                Status = if ($isExpired) { "EXPIRED" } else { "VALID" }
                Size = $csvSize
                File = $metaFile.Name -replace '\.meta\.json$', ''
            }
        }
        catch {
            Write-Host "  Error reading $($metaFile.Name): $_" -ForegroundColor Red
        }
    }

    $cacheItems | Format-Table -AutoSize -Property Table, Records, Age, TTL, Status, Size

    $totalSize = (Get-ChildItem -Path $CacheDir -File | Measure-Object -Property Length -Sum).Sum
    Write-Host "Total cache size: $(if($totalSize -gt 1MB){'{0:N1} MB' -f ($totalSize/1MB)}else{'{0:N1} KB' -f ($totalSize/1KB)})" -ForegroundColor Cyan
    Write-Host "Cache entries: $($cacheItems.Count)" -ForegroundColor Cyan
}

# ============================================================
# Clear Expired Cache
# ============================================================
function Clear-ExpiredCache {
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host "  Clearing Expired Cache" -ForegroundColor Magenta
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host ""

    $metaFiles = Get-ChildItem -Path $CacheDir -Filter "*.meta.json" -ErrorAction SilentlyContinue
    $cleared = 0
    $freedSpace = 0

    foreach ($metaFile in $metaFiles) {
        try {
            $meta = Get-Content $metaFile.FullName -Raw | ConvertFrom-Json
            $cacheTime = [DateTime]::Parse($meta.CacheTime)
            $age = (Get-Date) - $cacheTime
            $ttlHours = if ($meta.CacheTTL) { $meta.CacheTTL } else { 24 }

            if ($age.TotalHours -gt $ttlHours) {
                $csvFile = $metaFile.FullName -replace '\.meta\.json$', '.csv'
                if (Test-Path $csvFile) {
                    $freedSpace += (Get-Item $csvFile).Length
                    Remove-Item $csvFile -Force
                }
                Remove-Item $metaFile.FullName -Force
                $cleared++
                Write-Host "  Cleared: $($metaFile.Name -replace '\.meta\.json$', '') ($($meta.RecordCount) records)" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "  Error processing $($metaFile.Name): $_" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "Cleared $cleared expired cache entries" -ForegroundColor Green
    Write-Host "Freed $(if($freedSpace -gt 1MB){'{0:N1} MB' -f ($freedSpace/1MB)}else{'{0:N1} KB' -f ($freedSpace/1KB)})" -ForegroundColor Green
}

# ============================================================
# Clear All Cache
# ============================================================
function Clear-AllCache {
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host "  Clearing All Cache" -ForegroundColor Magenta
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host ""

    $totalSize = (Get-ChildItem -Path $CacheDir -File | Measure-Object -Property Length -Sum).Sum
    $fileCount = (Get-ChildItem -Path $CacheDir -File).Count

    Remove-Item -Path "$CacheDir\*" -Force -Recurse

    Write-Host "Cleared $fileCount cache files" -ForegroundColor Green
    Write-Host "Freed $(if($totalSize -gt 1MB){'{0:N1} MB' -f ($totalSize/1MB)}else{'{0:N1} KB' -f ($totalSize/1KB)})" -ForegroundColor Green
}

# ============================================================
# Show Cache Stats
# ============================================================
function Show-CacheStats {
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host "  Cache Statistics" -ForegroundColor Magenta
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host ""

    $metaFiles = Get-ChildItem -Path $CacheDir -Filter "*.meta.json" -ErrorAction SilentlyContinue
    $validCount = 0
    $expiredCount = 0
    $totalRecords = 0

    foreach ($metaFile in $metaFiles) {
        try {
            $meta = Get-Content $metaFile.FullName -Raw | ConvertFrom-Json
            $cacheTime = [DateTime]::Parse($meta.CacheTime)
            $age = (Get-Date) - $cacheTime
            $ttlHours = if ($meta.CacheTTL) { $meta.CacheTTL } else { 24 }

            if ($age.TotalHours -gt $ttlHours) {
                $expiredCount++
            } else {
                $validCount++
                $totalRecords += $meta.RecordCount
            }
        }
        catch {}
    }

    $totalSize = (Get-ChildItem -Path $CacheDir -File | Measure-Object -Property Length -Sum).Sum

    Write-Host "Cache directory: $CacheDir" -ForegroundColor Cyan
    Write-Host "Total entries: $($metaFiles.Count)" -ForegroundColor Cyan
    Write-Host "Valid entries: $validCount" -ForegroundColor Green
    Write-Host "Expired entries: $expiredCount" -ForegroundColor Yellow
    Write-Host "Total cached records: $totalRecords" -ForegroundColor Cyan
    Write-Host "Total cache size: $(if($totalSize -gt 1MB){'{0:N1} MB' -f ($totalSize/1MB)}else{'{0:N1} KB' -f ($totalSize/1KB)})" -ForegroundColor Cyan
}

# ============================================================
# Main
# ============================================================
switch ($Action) {
    "list" { Show-CacheList }
    "clear" { Clear-ExpiredCache }
    "clear-all" { Clear-AllCache }
    "stats" { Show-CacheStats }
}