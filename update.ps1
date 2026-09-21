#requires -version 5
<#
  update.ps1 - regenerate the Vessel Schedule dashboard and push it to
  GitHub Pages.

  Modes:
    -Mode full   (default) replace the whole dataset with a complete SKED
    -Mode merge  overlay a partial SKED (some vessels only, each with its
                 full rotation) onto the current dataset

  Entry points:
    update.bat            -> full, newest *-SKED.xls in .\input
    drag .xls on update.bat        -> full, that file (archived into .\input)
    drag .xls on update-partial.bat -> merge, that file (archived into .\input)

  Raw SKED files live in .\input — a dragged file from anywhere else is
  copied in there first, so .\input ends up holding every source file
  ever processed.

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
$inputDir = Join-Path $PSScriptRoot 'input'
if (-not (Test-Path -LiteralPath $inputDir)) { New-Item -ItemType Directory -Path $inputDir | Out-Null }

if ([string]::IsNullOrWhiteSpace($File)) {
    if ($Mode -eq 'merge') {
        Fail 'โหมด merge ต้องลากไฟล์ SKED (บางเรือ) มาวางบน update-partial.bat'
    }
    $newest = Get-ChildItem -LiteralPath $inputDir -Filter '*-SKED.xls' -File |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $newest) { Fail "ไม่พบไฟล์ *-SKED.xls ใน $inputDir" }
    $name = $newest.Name
} else {
    if (-not (Test-Path -LiteralPath $File)) { Fail "ไม่พบไฟล์: $File" }
    $srcFull = (Resolve-Path -LiteralPath $File).Path
    $name = Split-Path $File -Leaf
    $destFull = Join-Path $inputDir $name
    if ($srcFull -ne $destFull) {
        Copy-Item -LiteralPath $srcFull -Destination $destFull -Force
    }
}
$relPath = Join-Path 'input' $name

Write-Host '======================================================'
if ($Mode -eq 'merge') {
    Write-Host " MERGE (เฉพาะบางเรือ) : $name" -ForegroundColor Yellow
} else {
    Write-Host " FULL (แทนทั้งหมด)    : $name"
}
Write-Host '======================================================'
Write-Host ''

# --- generate --------------------------------------------------------
if ($Mode -eq 'merge') { $genArgs = @('generate_vessel_schedule.py', '--merge', $relPath) }
else                   { $genArgs = @('generate_vessel_schedule.py', '--full',  $relPath) }

python @genArgs
if ($LASTEXITCODE -ne 0) { Fail 'generate_vessel_schedule.py ล้มเหลว (ดู error ด้านบน)' }

Copy-Item 'Vessel Schedule.html' 'index.html' -Force

# --- commit + push --------------------------------------------------
# Targeted add (never -A) so stray files sitting in this folder never
# ride along into the public repo.
$tag = [System.IO.Path]::GetFileNameWithoutExtension($name)
git add -- $relPath 'Vessel Schedule.html' 'Vessel Schedule.xlsx' 'index.html' 'sked_current.xlsx' 'sked_prev.xlsx'
git commit -m "Update ($Mode): $tag" | Out-Host
git push origin main | Out-Host
if ($LASTEXITCODE -ne 0) { Fail 'git push ล้มเหลว (ตรวจ internet / สิทธิ์ GitHub)' }

Write-Host ''
Write-Host 'เสร็จแล้ว — เว็บจะอัปเดตอัตโนมัติใน ~1 นาที:' -ForegroundColor Green
Write-Host '  https://sirichai1265.github.io/heung-a-sked/'
Write-Host ''
Read-Host 'กด Enter เพื่อปิด'
