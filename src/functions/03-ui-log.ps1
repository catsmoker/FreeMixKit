# ==============================================================================
# LOGGING AND STATUS
# ==============================================================================
# Write-Log keeps its original signature and file-append behaviour, so the ~30
# call sites inside the module bodies work unchanged. The only difference is
# that output also goes to the UI log pane when a window exists.
#
# SCOPE: the UI is reached through delegates built on the UI thread at boot and
# stored on $sync (see New-FmkSyncState). Scriptblocks are never created in a
# background runspace and handed to Dispatcher.Invoke -- a scriptblock carries
# the session state it was created in, so that would re-enter the runspace that
# is currently blocked on the Invoke. Instead the caller parks its arguments on
# $sync, which is thread-safe, and the UI-thread delegate reads them back.
#
# Write-Log is reached from the UI thread *and* from background runspaces, so
# both paths exist and Invoke (not BeginInvoke) keeps log order matching
# execution order. On the UI thread it runs inline.

$FmkLogColorKeys = @{
    Info    = 'FmkFg'
    Success = 'FmkRiskLow'
    Warn    = 'FmkRiskMed'
    Error   = 'FmkRiskHigh'
    # Framework lifecycle chatter (dispatch notices, selection echoes) renders
    # muted gray so it stays structurally distinct from real tool output
    # (white/amber/red) and from result verdicts (green/amber/red).
    Muted   = 'FmkFgMuted'
}

# Bounded so a chatty module cannot grow the document without limit.
$FmkMaxLogLines = 2000

function Invoke-FmkUiAction {
    <#
        Runs one of the UI-thread delegates from wherever we happen to be.
        Returns quietly if the UI is not up yet, so logging during boot and in
        library mode falls through to the console path instead of throwing.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $window = $null
    if ($global:sync) { $window = $global:sync.FmkWindow }
    if (-not $window) { return $false }

    $delegate = $null
    if ($global:sync.ContainsKey($Name)) { $delegate = $global:sync[$Name] }
    if (-not $delegate) { return $false }

    try {
        [void]$window.Dispatcher.Invoke($delegate)
        return $true
    }
    catch {
        # Never let a display failure break the operation being displayed.
        return $false
    }
}

function Add-FmkLogLine {
    <# Appends one colored line. Must run on the UI thread. #>
    param(
        [string]$Text,
        [string]$ColorKey = 'FmkFg'
    )

    # Not [Parameter(Mandatory)]: this is invoked from a no-argument delegate, so
    # a missing value must be a silent no-op rather than an interactive prompt.
    if (-not $Text) { return }

    $box = $global:sync.LogBox
    if (-not $box) { return }

    $para = New-Object System.Windows.Documents.Paragraph
    $para.Margin = New-Object System.Windows.Thickness(0)
    $para.LineHeight = 15

    $brush = $null
    if ($global:sync.FmkWindow -and $global:sync.FmkWindow.Resources.Contains($ColorKey)) {
        $brush = $global:sync.FmkWindow.Resources[$ColorKey]
    }
    if (-not $brush) { $brush = [System.Windows.Media.Brushes]::LightGray }

    $run = New-Object System.Windows.Documents.Run($Text)
    $run.Foreground = $brush
    [void]$para.Inlines.Add($run)
    [void]$box.Document.Blocks.Add($para)

    while ($box.Document.Blocks.Count -gt $FmkMaxLogLines) {
        $box.Document.Blocks.Remove($box.Document.Blocks.FirstBlock)
    }

    $box.ScrollToEnd()
}

function Write-Log($Message, $Type = "Info") {
    # File append first, so a crash mid-log still leaves a record on disk.
    try {
        $tag = $Type.ToUpperInvariant()
        "[{0}] [{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $tag, $Message |
            Add-Content -Path $SessionLogPath -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch { }

    $line = "[{0}] {1}" -f (Get-Date).ToString('HH:mm:ss'), $Message
    $colorKey = if ($FmkLogColorKeys.ContainsKey($Type)) { $FmkLogColorKeys[$Type] } else { 'FmkFg' }

    $window = $null
    if ($global:sync -and $global:sync.FmkWindow) { $window = $global:sync.FmkWindow }

    if ($window) {
        $global:sync.PendingLogText = $line
        $global:sync.PendingLogColor = $colorKey
        try {
            [void](Invoke-FmkUiAction -Name 'ActLogFlush')
        }
        finally {
            $global:sync.PendingLogText = $null
            $global:sync.PendingLogColor = $null
        }
    }
    else {
        $consoleColor = switch ($Type) {
            "Success" { "Cyan" }
            "Warn" { "Yellow" }
            "Error" { "Red" }
            default { "White" }
        }
        Write-Host " $line" -ForegroundColor $consoleColor
    }
}

function Set-FmkStatusNow {
    <# Applies $global:sync.PendingStatusText. Must run on the UI thread. #>
    $state = $global:sync
    if (-not $state.StatusText) { return }

    $state.StatusText.Text = $state.PendingStatusText
    if ($state.FmkWindow.Resources.Contains($state.PendingStatusColor)) {
        $state.StatusText.Foreground = $state.FmkWindow.Resources[$state.PendingStatusColor]
    }
}

function Set-FmkStatus {
    <# Updates the status line shown to the right of the progress bar. #>
    param(
        [string]$Text,
        [string]$ColorKey = 'FmkFgMuted'
    )

    if (-not ($global:sync -and $global:sync.FmkWindow -and $global:sync.StatusText)) { return }

    $global:sync.PendingStatusText = $Text
    $global:sync.PendingStatusColor = $ColorKey
    [void](Invoke-FmkUiAction -Name 'ActStatus')
}

function Get-FmkProgressRequest {
    <#
        Normalizes a caller-supplied percentage into a progress request for the
        one shared bar. A finite number is trustworthy determinate progress
        (clamped to 0..100); $null, NaN, infinity, or anything unparseable is
        indeterminate. Percentages are never invented here -- callers only pass
        numbers they actually measured.
    #>
    param($Percent)

    if ($null -eq $Percent) { return @{ Mode = 'Indeterminate'; Value = 0 } }
    try {
        $value = [double]$Percent
        if ([double]::IsNaN($value) -or [double]::IsInfinity($value)) {
            return @{ Mode = 'Indeterminate'; Value = 0 }
        }
        if ($value -lt 0) { $value = 0 }
        if ($value -gt 100) { $value = 100 }
        return @{ Mode = 'Determinate'; Value = $value }
    }
    catch {
        return @{ Mode = 'Indeterminate'; Value = 0 }
    }
}

function Set-FmkProgress {
    <#
        Progress-only update for the shared bar (busy state untouched). Rides
        the ActBusy delegate so it works from any thread without a second
        marshaller; a no-op while nothing is busy, since the bar is hidden.
    #>
    param(
        $Percent = $null,
        [switch]$Indeterminate,
        [string]$Text = ''
    )

    if (-not ($global:sync -and $global:sync.FmkWindow)) { return }

    if ($Indeterminate -or ($null -eq $Percent)) {
        $global:sync.PendingProgressMode = 'Indeterminate'
        $global:sync.PendingProgressValue = 0
    }
    else {
        $req = Get-FmkProgressRequest -Percent $Percent
        $global:sync.PendingProgressMode = $req.Mode
        $global:sync.PendingProgressValue = $req.Value
    }
    if ($Text) {
        $global:sync.PendingBusyText = $Text
        $global:sync.PendingStatusText = $Text
        $global:sync.PendingStatusColor = 'FmkFgMuted'
    }
    [void](Invoke-FmkUiAction -Name 'ActBusy')
}

function Get-FmkQueueOverallPercent {
    <#
        Queue-wide percentage from completed tasks plus the current task's
        measured fraction -- the only honest determinate number available when
        installers emit no progress of their own. Pure math, unit tested.
    #>
    param(
        [int]$Done = 0,
        [int]$Total = 0,
        [double]$ItemPercent = 0
    )
    if ($Total -le 0) { return 0 }
    if ($Done -lt 0) { $Done = 0 }
    if ($Done -gt $Total) { $Done = $Total }
    if ($ItemPercent -lt 0) { $ItemPercent = 0 }
    if ($ItemPercent -gt 100) { $ItemPercent = 100 }
    return ($Done + ($ItemPercent / 100)) / $Total * 100
}

function Update-FmkObservedProgress {
    <#
        Forwards an installer-observed percentage (parsed from real output
        text, never estimated) to the shared bar. Inside a queue it is folded
        into the overall completed/total math; for a lone task it is the task
        percentage. Forwards only on integer change. Safe to call from the
        background runspace or with no window at all (then a no-op).
    #>
    param([int]$ItemPercent)

    $st = $global:sync
    if (-not ($st -and $st.FmkWindow)) { return }
    if ($ItemPercent -lt 0) { $ItemPercent = 0 }
    if ($ItemPercent -gt 100) { $ItemPercent = 100 }

    $overall = $ItemPercent
    if ([bool]$st.QueueRunning) {
        $total = 0
        $done = 0
        try { $total = [int]$st.QueueTotal } catch { $total = 0 }
        try { $done = [int]$st.QueueDone } catch { $done = 0 }
        if ($total -gt 0) {
            $overall = Get-FmkQueueOverallPercent -Done $done -Total $total -ItemPercent $ItemPercent
        }
    }
    $rounded = [int][Math]::Round($overall)
    $sent = -1
    try { $sent = [int]$st.ObservedItemSent } catch { $sent = -1 }
    if ($rounded -ne $sent) {
        $st.ObservedItemSent = $rounded
        Set-FmkProgress -Percent $overall
    }
}

function Set-FmkBusyNow {
    <#
        Applies PendingBusy / PendingBusyText / PendingProgressMode /
        PendingProgressValue. UI thread only. This is the single place that
        locks and unlocks the interface, and every task path reaches it from a
        finally block -- so success, failure, and cancellation all restore.
        Terminal modes (Complete/Failed/Cancelled) keep a truthful end state
        on the bar after the overlay lifts, until the next task starts.
    #>
    $state = $global:sync
    $busy = [bool]$state.PendingBusy
    $text = [string]$state.PendingBusyText
    $mode = [string]$state.PendingProgressMode
    if (-not $mode) { $mode = if ($busy) { 'Indeterminate' } else { 'None' } }
    $value = 0
    try { $value = [double]$state.PendingProgressValue } catch { $value = 0 }

    # Loading overlay: dims the body, swallows its mouse input, keeps the
    # current operation and a cancel action visible. The log pane is outside
    # the overlay, so output keeps streaming and the shared bar stays in view.
    if ($state.BusyOverlay) {
        $state.BusyOverlay.Visibility = if ($busy) {
            [System.Windows.Visibility]::Visible
        }
        else {
            [System.Windows.Visibility]::Collapsed
        }
    }
    # Indeterminate means no measured percentage exists (winget, for example,
    # emits phase lines but never a number on a redirected pipe), so the
    # overlay says so explicitly instead of implying measurable progress.
    if ($state.BusyOverlayText -and $text) {
        $state.BusyOverlayText.Text = if ($busy -and ($mode -eq 'Indeterminate')) {
            "$text -- waiting for installer progress information"
        }
        else {
            $text
        }
    }
    # Only the queue is safely cancellable (it stops after the current step;
    # a running installer is never killed). Single tasks show no Cancel.
    if ($state.TaskCancelButton) {
        $showCancel = $busy -and [bool]$state.QueueRunning
        $state.TaskCancelButton.Visibility = if ($showCancel) {
            [System.Windows.Visibility]::Visible
        }
        else {
            [System.Windows.Visibility]::Collapsed
        }
    }

    if ($state.ProgressBar) {
        # Terminal states stay visible after the task ends (green 100 on true
        # success, red/amber frozen at the last measured value otherwise), so
        # failure never masquerades as completion. The next task start resets.
        $showBar = $busy -or ($mode -eq 'Complete') -or ($mode -eq 'Failed') -or ($mode -eq 'Cancelled')
        $state.ProgressBar.Visibility = if ($showBar) {
            [System.Windows.Visibility]::Visible
        }
        else {
            [System.Windows.Visibility]::Collapsed
        }
        try { $state.ProgressBar.IsIndeterminate = ($busy -and ($mode -eq 'Indeterminate')) } catch { }
        $barBrush = $null
        try {
            if ($mode -eq 'Failed') { $barBrush = $state.FmkWindow.Resources['FmkRiskHigh'] }
            elseif ($mode -eq 'Cancelled') { $barBrush = $state.FmkWindow.Resources['FmkRiskMed'] }
            else { $barBrush = $state.FmkWindow.Resources['FmkRiskLow'] }
            if ($barBrush) { $state.ProgressBar.Foreground = $barBrush }
        }
        catch { }
        if ($mode -eq 'Determinate' -or $mode -eq 'Failed' -or $mode -eq 'Cancelled') {
            try { $state.ProgressBar.Value = $value } catch { }
        }
        elseif ($mode -eq 'Complete') {
            try { $state.ProgressBar.Value = 100 } catch { }
        }
        elseif (-not $busy) {
            try { $state.ProgressBar.Value = 0 } catch { }
        }
    }

    # Numeric percentage beside the bar. Set only from measured values; empty
    # (not zero) whenever none exists, so indeterminate never fakes a number.
    if ($state.ProgressPercent) {
        $pctText = ''
        if ($mode -eq 'Complete') { $pctText = '100%' }
        elseif ($mode -eq 'Determinate' -or $mode -eq 'Failed' -or $mode -eq 'Cancelled') {
            $pctText = "$([int][Math]::Round($value))%"
        }
        $state.ProgressPercent.Text = $pctText
    }

    foreach ($btn in $state.RunButtons) { $btn.IsEnabled = -not $busy }

    # Keep the Software queue action bar in step with single-module runs without
    # adding it to RunButtons (the queue owns its enabled state).
    try {
        if ($state.InstallSelectedButton -or $state.ClearSelectionButton) {
            $count = 0
            if ($state.SoftwareSelection) { $count = $state.SoftwareSelection.Count }
            $queueBusy = [bool]$state.QueueRunning
            if ($state.InstallSelectedButton) {
                $state.InstallSelectedButton.IsEnabled = ($count -gt 0) -and (-not $busy) -and (-not $queueBusy)
            }
            if ($state.ClearSelectionButton) {
                $state.ClearSelectionButton.IsEnabled = ($count -gt 0) -and (-not $busy) -and (-not $queueBusy)
            }
        }
    }
    catch { }

    # Interaction lock: the overlay above swallows mouse input over the body,
    # and disabling these containers also drops their keyboard tab stops, so
    # tabs, cards, and page buttons are unreachable while exclusive work runs.
    # Header search, risk help, and the log toggle stay live on purpose:
    # filtering and log inspection never conflict with a running task.
    try {
        if ($state.SideBar) { $state.SideBar.IsEnabled = -not $busy }
        if ($state.MainContent) { $state.MainContent.IsEnabled = -not $busy }
    }
    catch { }

    if (-not $state.StatusText) { return }
    if ($busy -and $text) { $state.StatusText.Text = $text }
    elseif (-not $busy) { $state.StatusText.Text = 'Ready' }
}

function Set-FmkBusy {
    <#
        The single authoritative task-state switch: overlay, shared progress
        bar, status text, interaction lock, and action buttons. Busy without an
        explicit percentage is indeterminate -- a percentage is only shown
        when the caller measured one. Quick interactions (selecting cards,
        filtering, navigating) never call this, so they never lock.
    #>
    param(
        [Parameter(Mandatory)][bool]$Busy,
        [string]$Text = '',
        $Percent = $null
    )

    if (-not ($global:sync -and $global:sync.FmkWindow)) { return }

    $global:sync.PendingBusy = $Busy
    $global:sync.PendingBusyText = $Text
    if ($Busy) {
        if ($null -eq $Percent) {
            $global:sync.PendingProgressMode = 'Indeterminate'
            $global:sync.PendingProgressValue = 0
        }
        else {
            $req = Get-FmkProgressRequest -Percent $Percent
            $global:sync.PendingProgressMode = $req.Mode
            $global:sync.PendingProgressValue = $req.Value
        }
    }
    else {
        $global:sync.PendingProgressMode = 'None'
        $global:sync.PendingProgressValue = 0
    }
    [void](Invoke-FmkUiAction -Name 'ActBusy')
}

function Set-FmkTaskComplete {
    <#
        Parks the truthful end state after a task finishes: 100% green only on
        real success, red/amber frozen at the last measured value on failure
        or cancellation. The overlay is already down (Busy=$false); this only
        sets what the bar keeps showing until the next task starts.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('Success', 'Failed', 'Cancelled')][string]$Status
    )

    if (-not ($global:sync -and $global:sync.FmkWindow)) { return }

    $global:sync.PendingBusy = $false
    $global:sync.PendingProgressMode = if ($Status -eq 'Success') { 'Complete' }
    elseif ($Status -eq 'Failed') { 'Failed' }
    else { 'Cancelled' }
    if ($Status -eq 'Success') { $global:sync.PendingProgressValue = 100 }
    [void](Invoke-FmkUiAction -Name 'ActBusy')
}
