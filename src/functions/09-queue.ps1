# ==============================================================================
# SOFTWARE INSTALLATION QUEUE
# ==============================================================================
# Runs the Software multi-select in selection order, one module at a time.
# Order is the SoftwareSelection ArrayList order (see 06-cards.ps1), never
# sorted or re-derived. Filtering never mutates the selection.
#
# Safety: Medium/High entries get one combined pre-install review showing each
# item's reason; any High item requires explicit confirmation. Low/None items
# never force per-item popups. Windows security (UAC, signatures) is never
# bypassed. Failures continue to the next item; nothing retries forever and a
# failure never silently swaps in an unrelated installer.

function Get-FmkQueueRiskItems {
    param(
        [Parameter(Mandatory)][string[]]$Keys
    )
    $items = @()
    foreach ($k in $Keys) {
        if (-not $ModuleMeta.Contains($k)) { continue }
        $meta = $ModuleMeta[$k]
        $risk = [string]$meta.Risk
        if ($risk -eq 'Medium' -or $risk -eq 'High') { $items += $meta }
    }
    return $items
}

function Confirm-FmkQueueRisks {
    <#
        One combined review for the queue's Medium/High items. Returns $true to
        proceed, $false to cancel. High items require explicit confirmation via
        the dialog's Continue path; the message names each High concern.
    #>
    param(
        [Parameter(Mandatory)][string[]]$Keys
    )
    $risky = @(Get-FmkQueueRiskItems -Keys $Keys)
    if ($risky.Count -eq 0) { return $true }

    $lines = @('The selection includes rated items:')
    $hasHigh = $false
    foreach ($m in $risky) {
        $short = ''
        try { $short = Get-FmkRiskShort -Meta $m } catch { $short = "$($m.Risk)" }
        $lines += "- $($m.Label): $short"
        if ([string]$m.Risk -eq 'High') { $hasHigh = $true }
    }
    if ($hasHigh) {
        $lines += ''
        $lines += 'High-risk items require explicit confirmation. Continue only if you accept the specific concerns above.'
    }
    else {
        $lines += ''
        $lines += 'Continue with the installation queue?'
    }
    $msg = $lines -join "`n"
    return (Show-FmkConfirmDialog -Message $msg -Title 'Review rated items' -ConfirmText 'Continue' -CancelText 'Cancel')
}

function Invoke-FmkSoftwareQueue {
    <#
        UI entry point for Install Selected. Guards against duplicate runs via
        QueueRunning plus the shared Busy flag, snapshots the order up front,
        installs sequentially, supports Cancel for pending items, and reports a
        success/failed/cancelled summary with failure details in the log.
    #>
    param([Parameter(Mandatory)]$Sync)

    if ([bool]$Sync.QueueRunning) {
        Write-Log 'Install queue is already running; ignoring duplicate request.' 'Warn'
        return $null
    }
    if ([bool]$Sync.Busy) {
        Write-Log "Already running: $($Sync.BusyLabel). Queue request ignored." 'Warn'
        return $null
    }

    $keys = @(Get-FmkSoftwareSelection -Sync $Sync)
    if ($keys.Count -eq 0) { return $null }

    # Snapshot the order: later selection toggles must not mutate a running queue.
    $queue = @($keys)

    if (-not (Confirm-FmkQueueRisks -Keys $queue)) {
        Write-Log 'Queue canceled at the risk review.' 'Warn'
        return $null
    }

    $Sync.QueueRunning = $true
    $Sync.QueueCancel = $false
    $Sync.Busy = $true
    $Sync.BusyLabel = 'Install queue'
    $Sync.QueueTotal = $queue.Count
    $Sync.QueueDone = 0
    $Sync.ObservedItemPercent = -1
    $Sync.ObservedItemSent = -1
    Update-FmkInstallButton -Sync $Sync
    Write-Log "Queue: preparing $($queue.Count) app(s)" 'Muted'
    Set-FmkBusy -Busy $true -Text "Queue: preparing $($queue.Count) app(s)" -Percent 0

    $ok = 0
    $failed = 0
    $cancelled = 0
    $failures = @()
    $total = $queue.Count
    $index = 0

    try {
        foreach ($key in $queue) {
            $index++
            if ([bool]$Sync.QueueCancel) {
                # Remaining items count as cancelled, including this one.
                $cancelled = $total - $index + 1
                Write-Log "Queue canceled by user; $cancelled pending item(s) skipped." 'Warn'
                break
            }

            $label = $key
            if ($ModuleMeta.Contains($key)) { $label = $ModuleMeta[$key].Label }
            # Step/total status with the completed-task floor as determinate
            # progress through the one shared mechanism. Observed installer
            # percentages refine it within the item when the output carries
            # any; otherwise the floor stands with the waiting message.
            $Sync.QueueDone = $index - 1
            $Sync.ObservedItemPercent = -1
            $Sync.ObservedItemSent = -1
            $statusMsg = "Installing $label ($index of $total)"
            Write-Log $statusMsg 'Muted'
            Set-FmkBusy -Busy $true -Text $statusMsg -Percent ([double](100 * ($index - 1) / $total))
            if ($Sync.SelectionStatus) {
                $Sync.SelectionStatus.Text = "$statusMsg..."
            }

            $result = $null
            try {
                $result = Invoke-FmkRunspace -ActionKey $key -Sync $Sync
            }
            catch {
                $result = [pscustomobject]@{
                    Status        = 'Failed'
                    Message       = $_.Exception.Message
                    Duration      = [timespan]::Zero
                    RollbackHint  = ''
                    PreviousState = ''
                }
            }

            if (-not $result -or -not $result.PSObject.Properties['Status']) {
                $result = [pscustomobject]@{
                    Status        = 'Failed'
                    Message       = 'Module produced no result object.'
                    Duration      = [timespan]::Zero
                    RollbackHint  = ''
                    PreviousState = ''
                }
            }

            $rb = ''
            try { $rb = Get-ModuleMetaValue -ActionKey $key -Name 'RollbackHint' -DefaultValue '' } catch { $rb = '' }
            if (-not $rb -and $result.PSObject.Properties['RollbackHint']) { $rb = [string]$result.RollbackHint }
            $prev = ''
            if ($result.PSObject.Properties['PreviousState']) { $prev = [string]$result.PreviousState }
            $dur = [timespan]::Zero
            if ($result.PSObject.Properties['Duration'] -and $result.Duration -is [timespan]) { $dur = $result.Duration }

            Add-ModuleResult -Module $key -Status $result.Status -Message $result.Message -Duration $dur -RollbackHint $rb -PreviousState $prev

            if ($result.Status -eq 'Success') {
                $ok++
                Write-Log "Queue: $label installed ($index/$total)." 'Success'
            }
            elseif ($result.Status -eq 'Canceled') {
                $cancelled++
                Write-Log "Queue: $label canceled ($index/$total)." 'Warn'
            }
            else {
                $failed++
                $failures += "$label`: $($result.Message)"
                Write-Log "Queue: $label failed ($index/$total) - $($result.Message)" 'Error'
            }
        }
    }
    finally {
        # Guaranteed restore: the overlay, progress bar, and interaction lock
        # all clear here even when a step throws or the user cancels.
        $Sync.QueueRunning = $false
        $Sync.QueueCancel = $false
        $Sync.Busy = $false
        $Sync.BusyLabel = $null
        Set-FmkBusy -Busy $false
        Update-FmkInstallButton -Sync $Sync
    }

    $summary = "Queue finished: $ok installed, $failed failed, $cancelled cancelled (of $total)."
    if ($failed -eq 0 -and $cancelled -eq 0) {
        Write-Log $summary 'Success'
        Set-FmkTaskComplete -Status 'Success'
        Set-FmkStatus "Completed: $ok of $total installed." 'FmkRiskLow'
    }
    elseif ($failed -gt 0) {
        Write-Log $summary 'Error'
        Set-FmkTaskComplete -Status 'Failed'
        Set-FmkStatus 'Finished with errors -- see output.' 'FmkRiskHigh'
    }
    else {
        Write-Log $summary 'Warn'
        Set-FmkTaskComplete -Status 'Cancelled'
        Set-FmkStatus 'Queue canceled.' 'FmkRiskMed'
    }
    foreach ($f in $failures) { Write-Log "Failed: $f" 'Error' }
    if ($Sync.SelectionStatus) { $Sync.SelectionStatus.Text = $summary }

    if ($failures.Count -gt 0) {
        $detail = $summary + "`n`nFailures:`n- " + ($failures -join "`n- ") + "`n`nSee the output pane for details."
        [void](Show-FmkConfirmDialog -Message $detail -Title 'Queue finished' -ConfirmText 'OK' -CancelText 'Close')
    }

    return [pscustomobject]@{
        Total     = $total
        Success   = $ok
        Failed    = $failed
        Cancelled = $cancelled
        Failures  = $failures
    }
}

function Request-FmkQueueCancel {
    param([Parameter(Mandatory)]$Sync)
    if (-not [bool]$Sync.QueueRunning) { return }
    $Sync.QueueCancel = $true
    Write-Log 'Cancel requested; finishing the current install, then stopping.' 'Warn'
}
