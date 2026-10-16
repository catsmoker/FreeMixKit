# ==============================================================================
# MODULE REGISTRY
# ==============================================================================

$Modules = [ordered]@{}
$ModuleMeta = [ordered]@{}
$ModuleResults = New-Object System.Collections.Generic.List[object]

function Register-Module(
    [string]$Key,
    [string]$Label,
    [string]$Description,
    [scriptblock]$Action,
    [string]$Risk = "None",
    [hashtable]$Options = @{}
) {
    $Modules[$Key] = $Action
    $ModuleMeta[$Key] = [ordered]@{
        Key             = $Key
        Label           = $Label
        Description     = $Description
        Risk            = $Risk
        RequiresNetwork = $false
        ConfirmMessage  = $null
        Verify          = $null
        RollbackHint    = ""
        # UI placement. Category drives the sidebar; Group drives the section
        # heading within a page. Both are optional so a module can be registered
        # without appearing in the GUI.
        Category        = "Other"
        Group           = ""
        # Key into $FmkLogos (src/assets/logos.ps1) for the Software grid. A
        # module with no entry -- or no Logo key at all -- still draws, falling
        # back to the label's initials, so this never has to be set.
        Logo            = ""
        # Button caption. "Run" suits a maintenance task; a module that installs
        # something reads better as "Install", and one that uninstalls or opens a
        # page reads better as its own verb.
        ActionLabel     = "Run"
        # Risk metadata. Only set on entries that carry a Low/Medium/High
        # rating; ordinary tools stay Risk="None" with these empty. Kept here
        # (not in UI code) so contributors update ratings without touching UI.
        # RiskReason is the short tooltip line; the Risk* dimensions feed the
        # detailed tooltip and the risk-help dialog. RiskEvidence is a URL or
        # short reference; RiskReviewed is yyyy-MM-dd. Never fabricate audits.
        RiskReason      = ""
        RiskTechnical   = ""
        RiskSource      = ""
        RiskPrivacy     = ""
        RiskLicense     = ""
        RiskPrivileges  = ""
        RiskEvidence    = ""
        RiskReviewed    = ""
    }

    foreach ($k in $Options.Keys) {
        $ModuleMeta[$Key][$k] = $Options[$k]
    }
}

function Get-ModuleMetaValue([string]$ActionKey, [string]$Name, $DefaultValue) {
    if ($ModuleMeta.Contains($ActionKey) -and $ModuleMeta[$ActionKey].Contains($Name) -and $null -ne $ModuleMeta[$ActionKey][$Name]) {
        return $ModuleMeta[$ActionKey][$Name]
    }
    return $DefaultValue
}

function Add-ModuleResult([string]$Module, [string]$Status, [string]$Message, [timespan]$Duration, [string]$RollbackHint = "", [string]$PreviousState = "") {
    $ModuleResults.Add([pscustomobject]@{
            Timestamp     = Get-Date
            Module        = $Module
            Status        = $Status
            DurationSec   = [Math]::Round($Duration.TotalSeconds, 2)
            Message       = $Message
            RollbackHint  = $RollbackHint
            PreviousState = $PreviousState
        })
}

function Export-SessionResults {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $resultsPath = Join-Path $DataLogs "FreeMixKit_ModuleResults_$stamp.csv"

    if ($ModuleResults.Count -gt 0) {
        $ModuleResults | Export-Csv -Path $resultsPath -NoTypeInformation -Encoding UTF8
        Write-Log "Module results exported to: $resultsPath" "Success"
    }
    else {
        Write-Log "No module executions to export for this session." "Info"
    }

}

# Last progress record surfaced to the status line, so Write-Progress ticks
# update the status instead of spamming the log. Reset per module run below.
$Script:FmkLastProgress = @{ Activity = ''; Percent = -2 }

function Format-FmkStreamRecord {
    <#
        Classifies one merged-stream record into how the Output panel should
        show it. Pure function: no UI, no logging, safe to unit test.
        Returns @{ Kind; Type; Text; Percent; Activity } where Type is the
        Write-Log color key. Progress records carry no log text at all.
    #>
    param($Record)

    if ($null -eq $Record) { return @{ Kind = 'Empty'; Type = 'Info'; Text = ''; Percent = -1; Activity = '' } }

    if ($Record -is [System.Management.Automation.ErrorRecord]) {
        $detail = ''
        try { $detail = [string]$Record.Exception.Message } catch { }
        if (-not $detail) {
            try { $detail = $Record.ToString() } catch { $detail = 'Unknown error' }
        }
        $category = ''
        try { $category = [string]$Record.CategoryInfo.Category } catch { }
        $text = if ($category) { "ERROR [$category]: $detail" } else { "ERROR: $detail" }
        return @{ Kind = 'Error'; Type = 'Error'; Text = $text; Percent = -1; Activity = '' }
    }

    if ($Record -is [System.Management.Automation.WarningRecord]) {
        return @{ Kind = 'Warning'; Type = 'Warn'; Text = "WARNING: $($Record.Message)"; Percent = -1; Activity = '' }
    }

    if ($Record -is [System.Management.Automation.VerboseRecord]) {
        return @{ Kind = 'Verbose'; Type = 'Info'; Text = "VERBOSE: $($Record.Message)"; Percent = -1; Activity = '' }
    }

    if ($Record -is [System.Management.Automation.DebugRecord]) {
        return @{ Kind = 'Debug'; Type = 'Info'; Text = "DEBUG: $($Record.Message)"; Percent = -1; Activity = '' }
    }

    if ($Record -is [System.Management.Automation.InformationRecord]) {
        $infoText = ''
        try { $infoText = [string]$Record.MessageData } catch { $infoText = $Record.ToString() }
        return @{ Kind = 'Information'; Type = 'Info'; Text = $infoText; Percent = -1; Activity = '' }
    }

    if ($Record -is [System.Management.Automation.ProgressRecord]) {
        $pct = -1
        try { $pct = [int]$Record.PercentComplete } catch { $pct = -1 }
        $activity = ''
        try { $activity = [string]$Record.Activity } catch { }
        return @{ Kind = 'Progress'; Type = 'Info'; Text = ''; Percent = $pct; Activity = $activity }
    }

    # Ordinary success output. Strings pass through untouched; anything else
    # is formatted the way the console would render one object. Trailing
    # whitespace (line breaks, padding) is trimmed, but leading indentation
    # is preserved verbatim.
    $body = ''
    if ($Record -is [string]) { $body = $Record }
    else {
        try { $body = (@($Record) | Out-String) } catch { $body = [string]$Record }
    }
    if ($null -eq $body) { $body = '' }
    $body = ([string]$body).TrimEnd()
    return @{ Kind = 'Output'; Type = 'Info'; Text = $body; Percent = -1; Activity = '' }
}

function Write-FmkStreamRecord {
    <#
        Sends one merged-stream record to the Output panel, preserving the
        stream distinction (errors red, warnings amber, verbose/debug tagged).
        Progress records update the status line instead -- throttled to a new
        activity or a >=5 point move, so a chatty Write-Progress loop cannot
        flood either surface. Reached from the background runspace; both sinks
        marshal to the UI thread through the existing delegates.
    #>
    param($Record)

    $fmt = Format-FmkStreamRecord -Record $Record
    if ($fmt.Kind -eq 'Empty') { return }
    if ($fmt.Kind -eq 'Progress') {
        $last = $Script:FmkLastProgress
        $show = ($fmt.Activity -ne $last.Activity) -or ([Math]::Abs($fmt.Percent - $last.Percent) -ge 5)
        if ($show) {
            $Script:FmkLastProgress = @{ Activity = $fmt.Activity; Percent = $fmt.Percent }
            $statusText = $fmt.Activity
            if ($fmt.Percent -ge 0) { $statusText += " ($($fmt.Percent)%)" }
            if (-not $statusText) { $statusText = 'Working...' }
            Set-FmkStatus -Text $statusText
        }
        return
    }
    if ($fmt.Text -eq '') { return }
    Write-Log $fmt.Text $fmt.Type
}

function Invoke-ModuleAction([string]$ActionKey) {
    if (-not $Modules.Contains($ActionKey)) {
        throw "Unknown action '$ActionKey'."
    }

    $label = Get-ModuleMetaValue -ActionKey $ActionKey -Name "Label" -DefaultValue $ActionKey
    $requiresNetwork = [bool](Get-ModuleMetaValue -ActionKey $ActionKey -Name "RequiresNetwork" -DefaultValue $false)
    $confirmMessage = Get-ModuleMetaValue -ActionKey $ActionKey -Name "ConfirmMessage" -DefaultValue $null
    $verifyBlock = Get-ModuleMetaValue -ActionKey $ActionKey -Name "Verify" -DefaultValue $null
    $rollbackHint = Get-ModuleMetaValue -ActionKey $ActionKey -Name "RollbackHint" -DefaultValue ""

    if ($confirmMessage) {
        if (-not (Confirm-Action $confirmMessage)) {
            return [pscustomobject]@{
                Status        = "Canceled"
                Message       = "Canceled by user"
                Duration      = [timespan]::Zero
                RollbackHint  = $rollbackHint
                PreviousState = ""
            }
        }
    }

    if ($requiresNetwork -and -not (Test-NetworkConnectivity)) {
        throw "Network is required, but connectivity check failed."
    }

    $Script:ModuleExecutionContext = @{}
    $Script:FmkLastProgress = @{ Activity = ''; Percent = -2 }
    $start = Get-Date
    $hadLastExitCode = Test-Path variable:global:LASTEXITCODE
    $oldLastExitCode = if ($hadLastExitCode) { $global:LASTEXITCODE } else { 0 }
    $global:LASTEXITCODE = 0

    try {
        # Merged-stream execution: every record the module emits -- success
        # output, errors, warnings, verbose/debug/information, progress -- is
        # classified and streamed to the Output panel as it happens instead of
        # being discarded. Terminating errors still throw straight through to
        # the catch below, and the constructed result object (not stray output)
        # remains the return value, so the module contract is unchanged.
        & $Modules[$ActionKey] *>&1 | ForEach-Object { Write-FmkStreamRecord -Record $_ }

        $duration = (Get-Date) - $start

        $effectiveExitCode = if (Test-Path variable:global:LASTEXITCODE) { $global:LASTEXITCODE } else { 0 }
        if ($effectiveExitCode -ne 0) {
            throw "Module returned non-zero exit code: $effectiveExitCode"
        }

        if ($verifyBlock) {
            $ok = & $verifyBlock
            if (-not $ok) {
                throw "Post-check failed for module '$label'."
            }
        }

        $previousState = ""
        if ($Script:ModuleExecutionContext.ContainsKey("PreviousState")) {
            $previousState = [string]$Script:ModuleExecutionContext["PreviousState"]
        }

        return [pscustomobject]@{
            Status        = "Success"
            Message       = "Completed"
            Duration      = $duration
            RollbackHint  = $rollbackHint
            PreviousState = $previousState
        }
    }
    finally {
        if ($hadLastExitCode) {
            $global:LASTEXITCODE = $oldLastExitCode
        }
        else {
            Remove-Variable -Scope Global -Name LASTEXITCODE -ErrorAction SilentlyContinue
        }
    }
}
