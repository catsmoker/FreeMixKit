# ==============================================================================
# DIALOGS
# ==============================================================================
# Confirm-Action keeps its original signature so Invoke-ModuleAction's
# ConfirmMessage path works unchanged -- it just renders a real modal instead of
# a console prompt.
#
# SCOPE: Confirm-Action is reached from a *background runspace* (a module body
# calls it), so the dialog has to be built on the UI thread. The runspace cannot
# do that itself, and it must not hand Dispatcher.Invoke a scriptblock of its
# own -- that scriptblock would carry the runspace's session state and re-enter
# the runspace that is currently blocked on the Invoke. So the arguments travel
# through $global:sync and the work happens in New-FmkConfirmWindow, reached via
# the ActShowDialog delegate built on the UI thread at boot.

function New-FmkConfirmWindow {
    <#
        Builds and shows the modal. Runs on the UI thread. Reads everything it
        needs out of $global:sync -- see the scope note above.
    #>
    $state = $global:sync
    $owner = $state.FmkWindow

    try {
        $dlg = New-Object System.Windows.Window
        $dlg.Title = $state.DlgTitle
        $dlg.Width = 500
        $dlg.MaxWidth = 640
        $dlg.SizeToContent = [System.Windows.SizeToContent]::Height
        $dlg.ResizeMode = [System.Windows.ResizeMode]::NoResize
        $dlg.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
        $dlg.ShowInTaskbar = $false
        $dlg.Owner = $owner
        $dlg.Background = $owner.Resources['FmkBgAlt']
        $dlg.Foreground = $owner.Resources['FmkFg']
        $dlg.FontFamily = $owner.FontFamily
        $dlg.FontSize = 13

        $root = New-Object System.Windows.Controls.Grid
        $root.Margin = New-Object System.Windows.Thickness(22)

        $rowText = New-Object System.Windows.Controls.RowDefinition
        $rowText.Height = [System.Windows.GridLength]::Auto
        $rowBar = New-Object System.Windows.Controls.RowDefinition
        $rowBar.Height = [System.Windows.GridLength]::Auto
        [void]$root.RowDefinitions.Add($rowText)
        [void]$root.RowDefinitions.Add($rowBar)

        $text = New-Object System.Windows.Controls.TextBlock
        $text.Text = $state.DlgMessage
        $text.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $text.Margin = New-Object System.Windows.Thickness(0, 0, 0, 20)
        [System.Windows.Controls.Grid]::SetRow($text, 0)
        [void]$root.Children.Add($text)

        $bar = New-Object System.Windows.Controls.StackPanel
        $bar.Orientation = [System.Windows.Controls.Orientation]::Horizontal
        $bar.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
        [System.Windows.Controls.Grid]::SetRow($bar, 1)

        $cancel = New-Object System.Windows.Controls.Button
        $cancel.Content = $state.DlgCancel
        $cancel.MinWidth = 96
        $cancel.Padding = New-Object System.Windows.Thickness(14, 7, 14, 7)
        $cancel.Margin = New-Object System.Windows.Thickness(0, 0, 8, 0)
        $cancel.IsCancel = $true
        $cancel.Add_Click({
                $global:sync.DialogResult = $false
                if ($global:sync.Dialog) { $global:sync.Dialog.Close() }
            })

        $ok = New-Object System.Windows.Controls.Button
        $ok.Content = $state.DlgConfirm
        $ok.MinWidth = 96
        $ok.Padding = New-Object System.Windows.Thickness(14, 7, 14, 7)
        $ok.IsDefault = $true
        $ok.Add_Click({
                $global:sync.DialogResult = $true
                if ($global:sync.Dialog) { $global:sync.Dialog.Close() }
            })

        [void]$bar.Children.Add($cancel)
        [void]$bar.Children.Add($ok)
        [void]$root.Children.Add($bar)

        $dlg.Content = $root
        $state.Dialog = $dlg

        [void]$dlg.ShowDialog()
    }
    finally {
        # Cleared even if ShowDialog throws, so a later Confirm-Action is not
        # left holding a dead window.
        $global:sync.Dialog = $null
    }
}

function Show-FmkConfirmDialog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$Title = 'FreeMixKit',
        [string]$ConfirmText = 'Continue',
        [string]$CancelText = 'Cancel'
    )

    $state = $global:sync
    if (-not $state) { return $false }

    $state.DialogResult = $false
    $state.Dialog = $null
    $state.DlgMessage = $Message
    $state.DlgTitle = $Title
    $state.DlgConfirm = $ConfirmText
    $state.DlgCancel = $CancelText

    try {
        [void](Invoke-FmkUiAction -Name 'ActShowDialog')
    }
    finally {
        foreach ($key in 'DlgMessage', 'DlgTitle', 'DlgConfirm', 'DlgCancel') {
            $state.Remove($key)
        }
    }

    return [bool]$state.DialogResult
}

function Confirm-Action([string]$Prompt) {
    # No window (library mode, or called before the UI is up): fall back to the
    # console so modules stay usable outside the GUI.
    if (-not ($global:sync -and $global:sync.FmkWindow)) {
        $response = Read-Host "$Prompt (Y/N)"
        return $response -match '^[Yy]$'
    }

    # ConfirmMessage strings already carry real newlines -- they were written as
    # double-quoted PowerShell literals, so `n was expanded at parse time. No
    # unescaping needed here.
    return Show-FmkConfirmDialog -Message $Prompt
}

function Show-FmkError {
    param([Parameter(Mandatory)][string]$Message, [string]$Title = 'FreeMixKit - Error')
    if (-not ($global:sync -and $global:sync.FmkWindow)) {
        Write-Host $Message -ForegroundColor Red
        return
    }
    [void](Show-FmkConfirmDialog -Message $Message -Title $Title -ConfirmText 'OK' -CancelText 'Close')
}
