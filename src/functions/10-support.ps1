# ==============================================================================
# ABOUT, ISSUE REPORTING, AND DIAGNOSTIC LOG SHARING
# ==============================================================================
# The About page lead-in (built by New-FmkAboutPanel, called from
# Build-FmkCards) plus the two support actions behind its buttons. Both actions
# run on the UI thread from button clicks -- they are plain functions, not
# $Modules entries, so they never enter the background runspace machinery and
# never touch the busy state.
#
# URLs come from $FmkRepoUrl (src/w.template.ps1), the single source of truth
# for the project repository. Nothing here guesses a URL.

function New-FmkAboutPanel {
    <#
        The About page lead-in: what Free Mix Kit is, what this version can
        actually do, and the support buttons. Descriptions cover only
        implemented capabilities.
    #>
    param([Parameter(Mandatory)]$Sync)

    $resources = $Sync.FmkWindow.Resources
    $muted = $null
    try {
        if ($resources.Contains('FmkFgMuted')) { $muted = $resources['FmkFgMuted'] }
    }
    catch { }

    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Margin = New-Object System.Windows.Thickness(2, 0, 0, 4)

    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = 'About Free Mix Kit'
    $title.FontSize = 14
    $title.FontWeight = [System.Windows.FontWeights]::Bold
    $title.Margin = New-Object System.Windows.Thickness(0, 0, 0, 6)
    [void]$panel.Children.Add($title)

    $blurb = New-Object System.Windows.Controls.TextBlock
    $blurb.Text = 'Free Mix Kit is a PowerShell-based Windows utility that helps you discover and install supported software, run available system utilities, and access useful Windows maintenance and configuration tools from one interface.'
    $blurb.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $blurb.Margin = New-Object System.Windows.Thickness(0, 0, 0, 6)
    if ($muted) { $blurb.Foreground = $muted }
    [void]$panel.Children.Add($blurb)

    $featHead = New-Object System.Windows.Controls.TextBlock
    $featHead.Text = 'Features'
    $featHead.FontWeight = [System.Windows.FontWeights]::SemiBold
    $featHead.Margin = New-Object System.Windows.Thickness(0, 6, 0, 2)
    [void]$panel.Children.Add($featHead)

    $features = @(
        'Software catalog: search, filter, select several applications, and install them in order with Install Selected.',
        'System cleanup and repair: junk cleanup, SFC/DISM repair, malware scans, and system reports.',
        'Utilities and fixes, including Windows tweak launchers.',
        'Activation helpers with risk ratings and explicit confirmations.',
        'Risk badges with per-item explanations, live search, and a streaming output log.'
    )
    foreach ($f in $features) {
        $item = New-Object System.Windows.Controls.TextBlock
        # ASCII hyphen, not a Unicode bullet: the built script ships without a
        # BOM, so a non-ASCII glyph depends on file encoding, host-codepage
        # script parsing, and font coverage all agreeing. A hyphen matches the
        # list style used everywhere else in the app (queue summaries, etc.).
        $item.Text = "- $f"
        $item.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $item.Margin = New-Object System.Windows.Thickness(8, 0, 0, 2)
        if ($muted) { $item.Foreground = $muted }
        [void]$panel.Children.Add($item)
    }

    $srcLine = New-Object System.Windows.Controls.TextBlock
    $srcLine.Text = "Source code: $FmkRepoUrl"
    $srcLine.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $srcLine.Margin = New-Object System.Windows.Thickness(0, 8, 0, 2)
    if ($muted) { $srcLine.Foreground = $muted }
    [void]$panel.Children.Add($srcLine)

    $row = New-Object System.Windows.Controls.StackPanel
    $row.Orientation = [System.Windows.Controls.Orientation]::Horizontal
    $row.Margin = New-Object System.Windows.Thickness(0, 8, 0, 10)

    $report = New-Object System.Windows.Controls.Button
    $report.Content = 'Report an Issue'
    $report.Style = $Sync.FmkWindow.FindResource('GhostButton')
    $report.Margin = New-Object System.Windows.Thickness(0, 0, 8, 0)
    $report.ToolTip = 'Open the official GitHub Issues page in your browser'
    # [void]: Add_Click returns a value that must not reach the pipeline.
    [void]$report.Add_Click({ [void](Open-FmkIssuesPage) })
    [void]$row.Children.Add($report)

    $share = New-Object System.Windows.Controls.Button
    $share.Content = 'Share Logs'
    $share.Style = $Sync.FmkWindow.FindResource('GhostButton')
    $share.ToolTip = 'Save sanitized diagnostics to attach to a bug report'
    [void]$share.Add_Click({ [void](Export-FmkDiagnosticLog) })
    [void]$row.Children.Add($share)

    [void]$panel.Children.Add($row)
    return $panel
}

function Get-FmkIssuesUrl {
    <# The official Issues page, derived from the single repo-URL source. #>
    return "$FmkRepoUrl/issues"
}

function Open-FmkIssuesPage {
    <#
        Opens the official GitHub Issues page. Never submits anything: the
        user files the issue themselves. A failed launch shows the URL to
        copy instead of failing silently.
    #>
    $url = Get-FmkIssuesUrl
    try {
        Start-Process $url | Out-Null
        Write-Log "Opened: $url" 'Info'
        return $true
    }
    catch {
        $msg = "Could not open the browser. Copy this URL manually:`n$url"
        if ($global:sync -and $global:sync.FmkWindow) {
            [void](Show-FmkConfirmDialog -Message $msg -Title 'Report an Issue' -ConfirmText 'OK' -CancelText 'Close')
        }
        else {
            Write-Host $msg
        }
        return $false
    }
}

function Protect-FmkLogSecrets {
    <#
        Redacts likely secrets from diagnostic text before export: key/value
        secrets (password, token, API key, ...), bearer tokens, and private
        key blocks. Pure function -- safe to unit test headlessly.
    #>
    param([string]$Text)

    if (-not $Text) { return '' }
    $out = $Text
    $out = [regex]::Replace($out,
        '(?i)(password|passwd|pwd|secret|api[-_]?key|access[-_]?token|auth[-_]?token|client[-_]?secret)(\s*[:=]\s*)([^\s,;"]+)',
        '$1$2[REDACTED]')
    $out = [regex]::Replace($out,
        '(?i)(Bearer\s+)([A-Za-z0-9\-._~+/=]{8,})',
        '$1[REDACTED]')
    $out = [regex]::Replace($out,
        '(?is)(-----BEGIN [A-Z0-9 ]*PRIVATE KEY-----).+?(-----END [A-Z0-9 ]*PRIVATE KEY-----)',
        '$1 [REDACTED] $2')
    return $out
}

function Export-FmkDiagnosticLog {
    <#
        Sanitized diagnostics export for bug reports. Collects the app version,
        timestamp, OS, recorded task results, and the session-log tail the app
        already retains -- nothing else, no credentials, no unrelated files.
        The user picks the save location; nothing is uploaded anywhere. The
        user attaches the file to a GitHub issue themselves.
    #>
    param(
        [string[]]$SessionLog = $null,
        [object[]]$Results = $null,
        [string]$SavePath = ''
    )

    if ($null -eq $SessionLog) {
        $SessionLog = @()
        try {
            if (Test-Path -LiteralPath $SessionLogPath) {
                $SessionLog = @(Get-Content -LiteralPath $SessionLogPath -Tail 300 -ErrorAction Stop)
            }
        }
        catch { $SessionLog = @() }
    }
    if ($null -eq $Results) {
        $Results = @()
        try { $Results = @($ModuleResults) } catch { $Results = @() }
    }

    if ((@($SessionLog).Count -eq 0) -and (@($Results).Count -eq 0)) {
        $msg = 'No diagnostic logs are available yet. Run a task first, then try Share Logs again.'
        if ($global:sync -and $global:sync.FmkWindow) {
            [void](Show-FmkConfirmDialog -Message $msg -Title 'Share Logs' -ConfirmText 'OK' -CancelText 'Close')
        }
        else {
            Write-Host $msg
        }
        return $null
    }

    if (-not $SavePath) {
        # Save dialog must run on the UI thread; this function is only ever
        # reached from a button click, which already is the UI thread.
        try {
            $dlg = New-Object Microsoft.Win32.SaveFileDialog
            $dlg.Filter = 'Text files (*.txt)|*.txt|Log files (*.log)|*.log'
            $dlg.FileName = "FreeMixKit-Diagnostics-$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $picked = $dlg.ShowDialog()
            if (-not $picked) { return $null }
            $SavePath = $dlg.FileName
        }
        catch {
            $msg = "Could not open the save dialog: $($_.Exception.Message)"
            if ($global:sync -and $global:sync.FmkWindow) {
                [void](Show-FmkConfirmDialog -Message $msg -Title 'Share Logs' -ConfirmText 'OK' -CancelText 'Close')
            }
            else {
                Write-Host $msg
            }
            return $null
        }
    }

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add('FreeMixKit diagnostic log')
    [void]$lines.Add("Generated: $((Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))")
    try { [void]$lines.Add("App version: v$AppVersion") } catch { }
    try { [void]$lines.Add("OS: $($SysInfo.OS)") } catch { }
    [void]$lines.Add("Tasks recorded: $(@($Results).Count)")
    [void]$lines.Add('')
    [void]$lines.Add('---- Task results ----')
    if (@($Results).Count -eq 0) {
        [void]$lines.Add('(no tasks recorded this session)')
    }
    else {
        foreach ($r in @($Results)) {
            try {
                [void]$lines.Add(("[{0}] Module={1} Status={2} Duration={3}s Message={4}" -f
                        $r.Timestamp, $r.Module, $r.Status, $r.DurationSec, $r.Message))
            }
            catch { }
        }
    }
    [void]$lines.Add('')
    [void]$lines.Add('---- Session log tail (sanitized) ----')
    if (@($SessionLog).Count -eq 0) {
        [void]$lines.Add('(session log unavailable)')
    }
    else {
        foreach ($line in @($SessionLog)) { [void]$lines.Add([string]$line) }
    }

    $body = Protect-FmkLogSecrets -Text ($lines -join "`r`n")
    try {
        [System.IO.File]::WriteAllText($SavePath, $body, (New-Object System.Text.UTF8Encoding($false)))
    }
    catch {
        $msg = "Could not write the log file: $($_.Exception.Message)"
        if ($global:sync -and $global:sync.FmkWindow) {
            [void](Show-FmkConfirmDialog -Message $msg -Title 'Share Logs' -ConfirmText 'OK' -CancelText 'Close')
        }
        else {
            Write-Host $msg
        }
        return $null
    }

    Write-Log "Diagnostics saved to: $SavePath" 'Success'
    if ($global:sync -and $global:sync.FmkWindow) {
        $confirm = Show-FmkConfirmDialog -Message "Saved to:`n$SavePath`n`nAttach it manually to a new GitHub issue." -Title 'Share Logs' -ConfirmText 'Open Issues' -CancelText 'Close'
        if ($confirm) { [void](Open-FmkIssuesPage) }
    }
    return $SavePath
}
