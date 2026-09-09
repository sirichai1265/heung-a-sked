#requires -version 5
<#
  update.ps1 - regenerate the Vessel Schedule dashboard from a new SKED file
  and push it to GitHub Pages.

  Usage:
    - Double-click update.bat  ............ uses the newest *-SKED.xls in this folder
    - Drag a .xls onto update.bat ......... uses that file as "today"
    - powershell -File update.ps1 9-10-SKED.xls
#>
param([string]$Today)

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

function Fail($msg) {
    Write-Host ''
    Write-Host $msg -ForegroundColor Red
    Read-Host 'กด Enter เพื่อปิด'
    exit 1
}

# --- pick the files -------------------------------------------------------
$skeds = Get-ChildItem -Filter '*-SKED.xls' -File | Sort-Object LastWriteTime -Descending
if ([string]::IsNullOrWhiteSpace($Today)) {
    if ($skeds.Count -eq 0) { Fail 'ไม่พบไฟล์ *-SKED.xls ในโฟลเดอร์นี้' }
    $Today = $skeds[0].Name
}
$todayName = Split-Path $Today -Leaf
if (-not (Test-Path -LiteralPath $todayName)) { Fail "ไม่พบไฟล์: $todayName" }

$yesterday = $skeds | Where-Object { $_.Name -ne $todayName } | Select-Object -First 1

Write-Host '======================================================'
Write-Host " TODAY     : $todayName"
if ($yesterday) { Write-Host " YESTERDAY : $($yesterday.Name)  (Early/Delay เทียบให้)" }
else            { Write-Host ' YESTERDAY : (ไม่มี - ข้าม Early/Delay)' }
Write-Host '======================================================'
Write-Host ''

# --- generate ------------------------------------------------------------
$genArgs = @('generate_vessel_schedule.py', '--today', $todayName)
if ($yesterday) { $genArgs += @('--yesterday', $yesterday.Name) }

python @genArgs
if ($LASTEXITCODE -ne 0) { Fail 'generate_vessel_schedule.py ล้มเหลว (ดู error ด้านบน)' }

Copy-Item 'Vessel Schedule.html' 'index.html' -Force

# --- commit + push -----------------------------------------------------
$tag = [System.IO.Path]::GetFileNameWithoutExtension($todayName)
git add -A
git commit -m "Update: $tag" | Out-Host
git push origin main | Out-Host
if ($LASTEXITCODE -ne 0) { Fail 'git push ล้มเหลว (ตรวจ internet / สิทธิ์ GitHub)' }

Write-Host ''
Write-Host 'เสร็จแล้ว — เว็บจะอัปเดตอัตโนมัติใน ~1 นาที:' -ForegroundColor Green
Write-Host '  https://sirichai1265.github.io/heung-a-sked/'
Write-Host ''
Read-Host 'กด Enter เพื่อปิด'
