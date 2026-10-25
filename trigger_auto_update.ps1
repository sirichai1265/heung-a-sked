# Fires the GitHub "Auto-update vessel schedule" workflow from this PC, waits
# for it, then pulls the result so the local "Vessel Schedule.html" /
# index.html are current (there is no online copy -- GitHub Pages is off).
# Run by the Windows Task Scheduler task "Heung-A Vessel Schedule auto-update"
# at 07:01, 12:01, 17:01, 22:01 (and on wake if a time was missed).
# scheduled=true makes the workflow skip if that slot was already updated
# by another trigger; the pull still picks up that other run's update.
# Each attempt is logged to %LOCALAPPDATA%\heung-a-sked-trigger.log

$log = Join-Path $env:LOCALAPPDATA 'heung-a-sked-trigger.log'
$repo = 'sirichai1265/heung-a-sked'
$gh = (Get-Command gh -ErrorAction SilentlyContinue).Source
if (-not $gh) { $gh = Join-Path $env:USERPROFILE 'bin\gh.exe' }
function Log($msg) { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg" | Add-Content -Path $log -Encoding utf8 }

$out = & $gh workflow run 'Auto-update vessel schedule' --repo $repo -f scheduled=true 2>&1 | Out-String
$code = $LASTEXITCODE
Log "exit=$code  $($out.Trim())"

# Wait for the newest run (ours, or one already in flight) to finish.
Start-Sleep -Seconds 15
$id = & $gh run list --repo $repo --workflow auto-update.yml --limit 1 --json databaseId -q '.[0].databaseId' 2>$null
if ($id) { & $gh run watch $id --repo $repo --interval 15 *> $null }

$pull = & git -C $PSScriptRoot pull --ff-only --quiet origin main 2>&1 | Out-String
Log "pull exit=$LASTEXITCODE  run=$id  $($pull.Trim())"
exit $code
