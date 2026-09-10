#requires -version 5
<#
  update.ps1 - regenerate the Vessel Schedule dashboard and push it to
  GitHub Pages.

  Modes:
    -Mode full   (default) replace the whole dataset with a complete SKED
    -Mode merge  overlay a partial SKED (some vessels only, each with its
                 full rotation) onto the current dataset

  Entry points:
    update.bat            -> full, newest *-SKED.xls in this folder
    drag .xls on update.bat        -> full, that file
    drag .xls on update-partial.bat -> merge, that file

  Early/Delay is computed automatically against the snapshot taken just
  before this update (sked_prev.xlsx).
#>
param(
    [string]$File,
    [ValidateSet('full', 'merge')][string]$Mode = 'full'
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$OutputEncoding = [System.Text.Encoding]::UTF8
Set-Location -LiteralPath $PSScriptRoot

function Fail($msg) {
    Write-Host ''
    Write-Host $msg -ForegroundColor Red
    Read-Host 'กด Enter เพื่อปิด'
    exit 1
}

# --- pick the input file -----------------------------------------------
if ([string]::IsNullOrWhiteSpace($File)) {
    if ($Mode -eq 'merge') {
        Fail 'โหมด merge ต้องลากไฟล์ SKED (บางเรือ) มาวางบน update-partial.bat'
    }
    $newest = Get-ChildItem -Filter '*-SKED.xls' -File |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $newest) { Fail 'ไม่พบไฟล์ *-SKED.xls ในโฟลเดอร์นี้' }
    $File = $newest.Name
}
$name = Split-Path $File -Leaf
if (-not (Test-Path -LiteralPath $name)) { Fail "ไม่พบไฟล์: $name" }

Write-Host '======================================================'
if ($Mode -eq 'merge') {
    Write-Host " MERGE (เฉพาะบางเรือ) : $name" -ForegroundColor Yellow
} else {
    Write-Host " FULL (แทนทั้งหมด)    : $name"
}
Write-Host '======================================================'
Write-Host ''

# --- generate --------------------------------------------------------
if ($Mode -eq 'merge') { $genArgs = @('generate_vessel_schedule.py', '--merge', $name) }
else                   { $genArgs = @('generate_vessel_schedule.py', '--full',  $name) }

python @genArgs
if ($LASTEXITCODE -ne 0) { Fail 'generate_vessel_schedule.py ล้มเหลว (ดู error ด้านบน)' }

Copy-Item 'Vessel Schedule.html' 'index.html' -Force

# --- commit + push --------------------------------------------------
$tag = [System.IO.Path]::GetFileNameWithoutExtension($name)
git add -A
git commit -m "Update ($Mode): $tag" | Out-Host
git push origin main | Out-Host
if ($LASTEXITCODE -ne 0) { Fail 'git push ล้มเหลว (ตรวจ internet / สิทธิ์ GitHub)' }

Write-Host ''
Write-Host 'เสร็จแล้ว — เว็บจะอัปเดตอัตโนมัติใน ~1 นาที:' -ForegroundColor Green
Write-Host '  https://sirichai1265.github.io/heung-a-sked/'
Write-Host ''
Read-Host 'กด Enter เพื่อปิด'
