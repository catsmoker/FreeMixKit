# ==============================================================================
# BACKGROUND EXECUTION
# ==============================================================================
# Modules block for minutes (sfc, DISM, winget). Running them on the UI thread
# would freeze the window, so each one runs on its own runspace.
#
# A fresh runspace has an empty session state -- none of our functions or
# modules exist in it. Rather than marshalling every function and variable
# across by hand, the runspace re-runs this script in "library mode"
# (-FmkLibraryMode), which defines everything and returns before touching the
# UI. That keeps a single source of truth for all module code.
#
# One runspace per invocation, not a pool: modules mutate process-wide state
# (scheduled tasks, registry, services) and reusing a runspace risks leaking
# that between runs. The ~1s startup cost is acceptable for operations that
# take minutes.

function Get-FmkLibraryPath {
    <#
        A path to a copy of this script that a runspace can dot-source.
        Running from a file we already have one; running from `irm | iex` we do
        not, so fetch a copy into the temp dir (that is what $ScriptUrl is for).
    #>
    if ($PSCommandPath -and (Test-Path -LiteralPath $PSCommandPath)) {
        return $PSCommandPath
    }

    $libPath = Join-Path $DataTemp 'FreeMixKit.Library.ps1'
    if (-not (Test-Path -LiteralPath $libPath)) {
        Write-Host "Fetching a runspace bootstrap copy..." -ForegroundColor DarkGray
        Invoke-WebRequest -Uri $ScriptUrl -OutFile $libPath -UseBasicParsing -ErrorAction Stop
    }
    return $libPath
}

function Wait-FmkPump {
    <#
        Runs a nested dispatcher frame for a short interval, which is what
        actually lets the UI repaint and lets background threads deliver their
        Dispatcher.Invoke calls while we are blocked waiting on a runspace.

        NOTE: this must be a real frame. Dispatcher.Invoke called *on* the
        dispatcher thread executes inline and does not drain the queue, so
        `Invoke([Background], {})` would look like a pump but process nothing.

        The frame and timer are held in globals on purpose: PowerShell
        scriptblocks do not capture the local scope they were created in, so a
        Tick handler cannot see function-local variables.
    #>
    param([int]$Milliseconds = 60)

    $global:FmkPumpFrame = New-Object System.Windows.Threading.DispatcherFrame
    $global:FmkPumpTimer = New-Object System.Windows.Threading.DispatcherTimer
    $global:FmkPumpTimer.Interval = [TimeSpan]::FromMilliseconds($Milliseconds)
    # [void]: this runs inside the wait loop in Invoke-FmkRunspace, so anything
    # that reached the pipeline here would be emitted by *that* function and turn
    # its single result object into an array.
    [void]$global:FmkPumpTimer.Add_Tick({
            # Null-checked because the globals are cleared in the finally block:
            # a tick that somehow outlives its frame must not throw inside a
            # dispatcher callback.
            if ($global:FmkPumpTimer) { $global:FmkPumpTimer.Stop() }
            if ($global:FmkPumpFrame) { $global:FmkPumpFrame.Continue = $false }
        })

    try {
        $global:FmkPumpTimer.Start()
        [System.Windows.Threading.Dispatcher]::PushFrame($global:FmkPumpFrame)
    }
    finally {
        $global:FmkPumpTimer.Stop()
        $global:FmkPumpTimer = $null
        $global:FmkPumpFrame = $null
    }
}

function Invoke-FmkRunspace {
    <#
        Runs one module on a background runspace and returns the same
        {Status, Message, Duration, ...} object Invoke-ModuleAction produces.
    #>
    param(
        [Parameter(Mandatory)][string]$ActionKey,
        [Parameter(Mandatory)]$Sync
    )

    $libPath = $null
    try {
        $libPath = Get-FmkLibraryPath
    }
    catch {
        Write-Log "Could not obtain a runspace bootstrap: $($_.Exception.Message)" "Warn"
    }

    if (-not $libPath) {
        # Degrade rather than fail: the module still runs, the window just
        # blocks for the duration.
        Write-Log "No bootstrap available; running on the UI thread (window will block)." "Warn"
        return Invoke-ModuleAction -ActionKey $ActionKey
    }

    $rs = $null
    $ps = $null
    try {
        $rs = [runspacefactory]::CreateRunspace()
        # STA, not MTA. Modules reach late-bound COM (WScript.Shell for the
        # shortcut module, Windows Installer) which the original script only ever
        # touched from PowerShell 5.1's STA thread. One module runs at a time, so
        # there is no throughput to gain from MTA and a real regression to lose.
        $rs.ApartmentState = 'STA'
        $rs.Open()

        $ps = [powershell]::Create()
        $ps.Runspace = $rs

        [void]$ps.AddScript({
                param($Sync, $LibraryPath, $ActionKey)

                # Library mode defines every function and module in this
                # runspace's scope, then returns.
                . $LibraryPath -FmkLibraryMode

                # Must be set AFTER the dot-source -- library mode initialises
                # $global:sync to $null on the way through.
                $global:sync = $Sync

                return Invoke-ModuleAction -ActionKey $ActionKey
            }).AddArgument($Sync).AddArgument($libPath).AddArgument($ActionKey)

        $handle = $ps.BeginInvoke()

        # Keep the window alive while the module runs. Without this the UI
        # thread would sit inside EndInvoke: no repaint, no log lines, and the
        # window would be reported as "not responding".
        while (-not $handle.IsCompleted) {
            Wait-FmkPump -Milliseconds 60
        }

        # EndInvoke hands back a Collection[PSObject]. Take the last element
        # explicitly rather than relying on pipeline unrolling, so a module that
        # accidentally emits extra output cannot turn $result into an array.
        $output = @($ps.EndInvoke($handle))
        if ($output.Count -gt 0) { return $output[$output.Count - 1] }
        return $null
    }
    finally {
        if ($ps) { $ps.Dispose() }
        if ($rs) {
            $rs.Close()
            $rs.Dispose()
        }
    }
}

function Invoke-FmkModule {
    <#
        UI-facing entry point: runs a module with busy state and result
        bookkeeping, swallowing failures into the exported result list.
    #>
    param(
        [Parameter(Mandatory)][string]$ActionKey,
        [Parameter(Mandatory)]$Sync
    )

    $label = Get-ModuleMetaValue -ActionKey $ActionKey -Name "Label" -DefaultValue $ActionKey

    # Re-entrancy guard. Set-FmkBusy disables every action button before the
    # module starts, which is what normally stops a second click, but a queued
    # click that was already dispatched can still arrive here. Refusing it
    # explicitly is what makes "clicking the same button twice installs once"
    # true of the dispatch path rather than only of the button's enabled state.
    if ([bool]$Sync.QueueRunning) {
        Write-Log "Install queue is running; ignored the request for '$label'." "Warn"
        return $null
    }
    if ($Sync.Busy) {
        Write-Log "Already running: $($Sync.BusyLabel). Ignored the request for '$label'." "Warn"
        return $null
    }
    $Sync.Busy = $true
    $Sync.BusyLabel = $label

    Set-FmkBusy -Busy $true -Text "Running: $label"
    Write-Log "Executing: $label" "Muted"

    try {
        $result = Invoke-FmkRunspace -ActionKey $ActionKey -Sync $Sync
    }
    catch {
        $result = [pscustomobject]@{
            Status        = "Failed"
            Message       = $_.Exception.Message
            Duration      = [timespan]::Zero
            RollbackHint  = Get-ModuleMetaValue -ActionKey $ActionKey -Name "RollbackHint" -DefaultValue ""
            PreviousState = ""
        }
    }
    finally {
        $Sync.Busy = $false
        $Sync.BusyLabel = $null
        Set-FmkBusy -Busy $false
    }

    # A module that returned nothing at all still gets recorded as a failure
    # rather than crashing the result bookkeeping below.
    if (-not $result -or -not $result.PSObject.Properties['Status']) {
        $result = [pscustomobject]@{
            Status        = "Failed"
            Message       = "Module produced no result object."
            Duration      = [timespan]::Zero
            RollbackHint  = ""
            PreviousState = ""
        }
    }

    Add-ModuleResult -Module $ActionKey `
        -Status $result.Status `
        -Message $result.Message `
        -Duration $result.Duration `
        -RollbackHint $result.RollbackHint `
        -PreviousState $result.PreviousState

    if ($result.Status -eq 'Success') {
        Write-Log "Result: $($result.Status) in $([Math]::Round($result.Duration.TotalSeconds, 2))s" "Success"
        Set-FmkTaskComplete -Status 'Success'
        Set-FmkStatus "Completed: $label" 'FmkRiskLow'
    }
    elseif ($result.Status -eq 'Canceled') {
        Write-Log "Result: canceled by user" "Warn"
        Set-FmkTaskComplete -Status 'Cancelled'
        Set-FmkStatus 'Canceled.' 'FmkRiskMed'
    }
    else {
        Write-Log "Result: $($result.Status) - $($result.Message)" "Error"
        Set-FmkTaskComplete -Status 'Failed'
        Set-FmkStatus 'Failed -- see output.' 'FmkRiskHigh'
    }

    if ($result.RollbackHint) { Write-Log "Rollback hint: $($result.RollbackHint)" "Warn" }
    if ($result.PreviousState) { Write-Log "Previous state: $($result.PreviousState)" "Info" }

    return $result
}
