# ==============================================================================
# BOOT
# ==============================================================================
# Loads the XAML, builds the $sync state object every other piece talks
# through, wires the chrome events, and shows the window.
#
# $sync is a [hashtable]::Synchronized() shared with the background runspaces --
# the same pattern WinUtil uses. Event handlers reach it through $global:sync,
# never through a captured local (PowerShell scriptblocks do not close over the
# scope they were created in).

function Get-FmkConsoleHandle {
    if (-not ('FmkNative.ConsoleHost' -as [type])) {
        Add-Type -Namespace FmkNative -Name ConsoleHost -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
    }
    return [FmkNative.ConsoleHost]::GetConsoleWindow()
}

function Hide-FmkConsole {
    <# SW_HIDE. Called once the window is up so only the GUI is visible. #>
    try {
        $handle = Get-FmkConsoleHandle
        if ($handle -ne [IntPtr]::Zero) { [void][FmkNative.ConsoleHost]::ShowWindow($handle, 0) }
    }
    catch { }
}

function Show-FmkConsole {
    <# SW_SHOW. Restored on exit and on crash, so a failure is never silent. #>
    try {
        $handle = Get-FmkConsoleHandle
        if ($handle -ne [IntPtr]::Zero) { [void][FmkNative.ConsoleHost]::ShowWindow($handle, 5) }
    }
    catch { }
}

function New-FmkSyncState {
    <#
        Builds $sync and populates it from every x:Name in the XAML, so adding
        a named element to ui.xaml is enough to make it reachable as
        $sync.<Name> -- no second list to keep in step.
    #>
    param([Parameter(Mandatory)]$Window, [Parameter(Mandatory)][string]$Xaml)

    $sync = [hashtable]::Synchronized(@{})
    $sync.FmkWindow = $Window
    # Only ever touched on the UI thread, but the busy-state code walks it.
    $sync.RunButtons = New-Object System.Collections.ArrayList
    # Re-entrancy flag read by Invoke-FmkModule. Lives on $sync rather than in a
    # function-local because the click handlers are separate scopes.
    $sync.Busy = $false
    $sync.BusyLabel = $null
    # Filled by Build-FmkCards: category -> list of @{ Heading; Body } sections.
    $sync.CardSections = @{}
    # category -> the search term last applied to it, so Apply-FmkFilter can tell
    # a changed query (reset the scroll) from a plain page switch (keep it).
    $sync.FilterTerms = @{}
    $sync.ActiveCategory = $null
    # Selected Software group chip (unset means All). Lives on $sync so the
    # click handlers -- separate scopes that only see $global:sync -- can share
    # it with Apply-FmkFilter.
    $sync.SoftwareGroup = $null
    # Software multi-select queue. SoftwareSelection is an ArrayList whose order
    # IS the install order (append on select, remove on deselect). QueueRunning
    # plus Busy guards duplicate runs; QueueCancel lets Cancel skip pending
    # items after the current one finishes.
    $sync.SoftwareSelection = New-Object System.Collections.ArrayList
    $sync.QueueRunning = $false
    $sync.QueueCancel = $false
    # Queue progress accounting, read by Update-FmkObservedProgress: totals
    # and completed counts are task facts, never estimates.
    $sync.QueueTotal = 0
    $sync.QueueDone = 0
    $sync.ObservedItemPercent = -1
    $sync.ObservedItemSent = -1
    # Shared progress state, applied by Set-FmkBusyNow. Empty mode means
    # "decide from Busy": indeterminate while busy, hidden when idle.
    $sync.PendingProgressMode = ''
    $sync.PendingProgressValue = 0
    # Which palette was baked into the XAML. Set by Initialize-FmkUi; recorded
    # here because it is UI state, not because anything reads it back yet.
    $sync.Theme = $null
    $sync.Dialog = $null
    $sync.DialogResult = $false
    # Set by module actions that want to report what they changed.
    $sync.LastResult = $null

    foreach ($match in [regex]::Matches($Xaml, 'x:Name="([^"]+)"')) {
        $name = $match.Groups[1].Value
        $element = $Window.FindName($name)
        if ($element) { $sync[$name] = $element }
    }

    # --------------------------------------------------------------------------
    # Marshalling delegates.
    #
    # These are built HERE, on the UI thread, and stored on $sync so a background
    # runspace can call $window.Dispatcher.Invoke($sync.ActLogFlush).
    #
    # Why it matters: a scriptblock created in the runspace carries the
    # runspace's own session state, and invoking that delegate from the UI thread
    # would re-enter the very runspace that is currently blocked waiting on the
    # Invoke -- PowerShell session state is not thread-safe, so that is a
    # deadlock at best. Passing a plain .NET delegate that was compiled against
    # the UI thread's session state avoids the problem entirely: the runspace
    # only ever hands over an object reference.
    #
    # The delegates take no arguments by design. Everything they need is parked
    # on $sync by the caller first, which is also what makes them reusable.
    # --------------------------------------------------------------------------
    $sync.ActLogFlush = [action]{ Add-FmkLogLine -Text $global:sync.PendingLogText -ColorKey $global:sync.PendingLogColor }
    $sync.ActStatus = [action]{ Set-FmkStatusNow }
    $sync.ActBusy = [action]{ Set-FmkBusyNow }
    $sync.ActShowDialog = [action]{ New-FmkConfirmWindow }

    return $sync
}

function Set-FmkTitleBar {
    <#
        Paints the standard window chrome to match the palette.

        WinUtil keeps the ordinary title bar rather than drawing its own, and so
        do we -- but Windows draws that bar light unless asked otherwise, which
        leaves a white strip across the top of an otherwise dark window.
    #>
    param(
        [Parameter(Mandatory)]$Window,
        [bool]$Dark = $true
    )

    try {
        if (-not ('FmkNative.Dwm' -as [type])) {
            Add-Type -Namespace FmkNative -Name Dwm -MemberDefinition @'
[DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
'@
        }

        # Creates the HWND if it does not exist yet, which is why this can run
        # before Show.
        $handle = (New-Object System.Windows.Interop.WindowInteropHelper($Window)).Handle
        if ($handle -eq [IntPtr]::Zero) { return }

        $value = if ($Dark) { 1 } else { 0 }
        # 20 is DWMWA_USE_IMMERSIVE_DARK_MODE from Win10 20H1 on. Builds in the
        # 17763..18984 window only understood 19, and return a failure code for
        # 20, so fall back rather than guessing the build.
        $hr = [FmkNative.Dwm]::DwmSetWindowAttribute($handle, 20, [ref]$value, 4)
        if ($hr -ne 0) { [void][FmkNative.Dwm]::DwmSetWindowAttribute($handle, 19, [ref]$value, 4) }
    }
    catch {
        # Cosmetic only: a light title bar is not worth failing startup over.
    }
}

function Initialize-FmkUi {
    param(
        [Parameter(Mandatory)][string]$Xaml,
        [string]$Version = '',
        [string]$IconPath = '',
        [ValidateSet('Dark', 'Light')][string]$Theme = 'Dark'
    )

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase -ErrorAction SilentlyContinue

    # The palette is substituted into the XAML text first, so the window is born
    # with its final colors. Applying a theme to a live window is not possible
    # here -- see the comment above the palette in ui.xaml.
    $themed = Get-FmkThemedXaml -Xaml $Xaml -Theme $Theme
    $window = [System.Windows.Markup.XamlReader]::Parse($themed)

    if ($IconPath -and (Test-Path -LiteralPath $IconPath)) {
        try {
            # OnLoad decodes eagerly and releases the file: the embedded icon
            # is rewritten to temp on every launch, so a lingering read lock
            # here would break the next startup. This Icon feeds the title bar
            # and the taskbar button.
            $iconUri = New-Object System.Uri($IconPath)
            $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create(
                $iconUri,
                [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
                [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        }
        catch { }
    }

    $sync = New-FmkSyncState -Window $window -Xaml $Xaml
    $sync.Theme = $Theme
    # Must be global before anything can log or marshal a call.
    $global:sync = $sync

    if ($Version -and $sync.VersionBadge) { $sync.VersionBadge.Text = "v$Version" }

    Build-FmkNav -Sync $sync
    Build-FmkCards -Sync $sync
    Set-FmkPage -Category $FmkCategories[0] -Sync $sync
    try { Update-FmkInstallButton -Sync $sync } catch { }

    # --- chrome events --------------------------------------------------------
    $sync.SearchBox.Add_TextChanged({
            Apply-FmkFilter -Sync $global:sync
        })

    # Software multi-select action bar. Handlers read $global:sync, never
    # locals, because scriptblocks do not close over the creating scope.
    try {
        if ($sync.InstallSelectedButton) {
            $sync.InstallSelectedButton.Add_Click({
                    [void](Invoke-FmkSoftwareQueue -Sync $global:sync)
                })
        }
        if ($sync.ClearSelectionButton) {
            $sync.ClearSelectionButton.Add_Click({
                    Clear-FmkSoftwareSelection -Sync $global:sync
                })
        }
        # The overlay Cancel: visible only while the queue runs (the one
        # safely cancellable task). Handlers read $global:sync, never locals,
        # because scriptblocks do not close over the creating scope.
        if ($sync.TaskCancelButton) {
            $sync.TaskCancelButton.Add_Click({
                    Request-FmkQueueCancel -Sync $global:sync
                })
        }
        if ($sync.RiskHelpButton) {
            $sync.RiskHelpButton.Add_Click({ Show-FmkRiskHelp })
        }
        if ($sync.GlobalRiskHelpButton) {
            $sync.GlobalRiskHelpButton.Add_Click({ Show-FmkRiskHelp })
        }
    }
    catch { }

    # Title bar colour. On SourceInitialized rather than before Show: the
    # attribute is dropped if it is set on an HWND DWM has not composited yet,
    # which is the state the window is in until it is shown.
    $window.Add_SourceInitialized({
            Set-FmkTitleBar -Window $this -Dark ($global:sync.Theme -eq 'Dark')
            # Keep the window inside the work area (taskbar excluded) at any DPI.
            # WPF units are device-independent, and WorkArea is reported in the
            # same units, so no manual DPI math is needed. Only shrink, never
            # grow: the XAML size stands on large displays, and the user can
            # still resize (ResizeMode stays CanResize).
            try {
                $area = [System.Windows.SystemParameters]::WorkArea
                if ($this.Width -gt $area.Width) { $this.Width = $area.Width }
                if ($this.Height -gt $area.Height) { $this.Height = $area.Height }
                if ($this.MaxWidth -gt $area.Width -or $this.MaxWidth -le 0) { $this.MaxWidth = $area.Width }
                if ($this.MaxHeight -gt $area.Height -or $this.MaxHeight -le 0) { $this.MaxHeight = $area.Height }
            }
            catch { }
        })

    # Console collapse. The row height is what sizes the pane -- collapsing the
    # RichTextBox instead would leave the 150px row standing empty.
    $sync.LogToggle.Add_Click({
            $row = $global:sync.RootGrid.RowDefinitions[2]
            if ($row.Height.Value -gt 40) {
                $row.Height = New-Object System.Windows.GridLength(32)
                # Up arrow: the pane is closed, so the button offers to open it.
                $global:sync.LogToggle.Content = [string][char]0x25B2
            }
            else {
                $row.Height = New-Object System.Windows.GridLength(150)
                $global:sync.LogToggle.Content = [string][char]0x25BC
            }
        })

    $window.Add_Closing({
            try { Export-SessionResults } catch { }
        })

    $window.Add_Closed({
            Show-FmkConsole
        })

    Write-Log "FreeMixKit v$Version started." "Success"
    Write-Log "OS: $($SysInfo.OS) | CPU: $($SysInfo.CPU) | RAM: $($SysInfo.RAM)" "Info"

    Hide-FmkConsole
    try {
        [void]$window.ShowDialog()
    }
    finally {
        # Also covers the path where ShowDialog throws rather than closing.
        Show-FmkConsole
        $global:sync = $null
    }
}
