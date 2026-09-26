# OpenOrder - worship bulletin & slide generator for churches
# Copyright (c) 2026 Rev. Dr. Jonathan Mellette (TheRevDrJ)
# Licensed under AGPL-3.0 - see LICENSE file for details
#
# Stops ONLY a stale OpenOrder BACKEND process - a pythonw/python running this
# project's uvicorn (app.main:app on port 6801) - and nothing else. The
# frontend is handled by OpenOrder.bat purely by its hard-coded port (6800),
# NOT here, on purpose: you may run several Vite servers at once, so this
# script never matches node/vite by name - that would risk killing another
# project's frontend.
#
# Safety (process-kills are destructive - kill NARROW, never
# broad): a process is a target ONLY if its command line contains BOTH
# 'app.main:app' AND the backend port '6801' - a fingerprint unique to this
# project's backend. An empty match list kills nothing; there is no fall-through
# to a broad match. Each PID is validated > 0 before Stop-Process is called with
# that explicit -Id. Exits with the number of processes killed, so the caller
# (OpenOrder.bat stop) can tell whether anything was running.
#
#   -Quiet    suppress the per-process line (used by 'start' pre-launch cleanup)
#   -DryRun   report what WOULD be killed, kill nothing (verification)

param([switch]$Quiet, [switch]$DryRun)

$targets = Get-CimInstance Win32_Process | Where-Object {
    ($_.Name -eq 'pythonw.exe' -or $_.Name -eq 'python.exe') -and
    $_.CommandLine -and
    ($_.CommandLine -like '*app.main:app*') -and
    ($_.CommandLine -like '*6801*')
}

$count = 0
foreach ($p in $targets) {
    if ($p.ProcessId -gt 0) {
        if (-not $Quiet) {
            $verb = if ($DryRun) { 'Would kill' } else { 'Killing' }
            Write-Host "  $verb stale OpenOrder backend (PID: $($p.ProcessId))..."
        }
        if (-not $DryRun) {
            Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
        }
        $count++
    }
}

exit $count
