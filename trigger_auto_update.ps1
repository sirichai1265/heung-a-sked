# Fires the GitHub "Auto-update vessel schedule" workflow from this PC.
# Run by the Windows Task Scheduler task "Heung-A Vessel Schedule auto-update"
# at 07:01, 12:01, 17:01, 22:01 (and on wake if a time was missed).
# scheduled=true makes the workflow skip if that slot was already updated
# by another trigger. Each attempt is logged to
# %LOCALAPPDATA%\heung-a-sked-trigger.log

$log = Join-Path $env:LOCALAPPDATA 'heung-a-sked-trigger.log'
$gh = (Get-Command gh -ErrorAction SilentlyContinue).Source
if (-not $gh) { $gh = Join-Path $env:USERPROFILE 'bin\gh.exe' }

$out = & $gh workflow run 'Auto-update vessel schedule' --repo sirichai1265/heung-a-sked -f scheduled=true 2>&1 | Out-String
"$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  exit=$LASTEXITCODE  $($out.Trim())" |
    Add-Content -Path $log -Encoding utf8
exit $LASTEXITCODE
