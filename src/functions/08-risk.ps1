# ==============================================================================
# RISK ASSESSMENT
# ==============================================================================
# Consistent risk-rating criteria for applicable tools. Ratings explain the
# potential consequences of using a tool (technical safety, system changes,
# privacy, source trustworthiness, licensing or service-policy concerns).
# They are guidance, not a guarantee of safety.
#
# Levels: None (untagged, no meaningful warning), Low, Medium, High.
# Low, Medium, and High each get a badge in their own color; None draws
# nothing. Untagged stays untagged -- Low is never assigned just to show
# a badge.
#
# Dimensions stored per entry (see Register-Module defaults in 02-registry):
#   RiskTechnical  - could it damage the system, compromise security, lose data?
#   RiskSource     - is the download source authentic and trustworthy?
#   RiskPrivacy    - does it collect, expose, or transmit sensitive info?
#   RiskLicense    - could use violate terms, license, or agreements?
#   RiskPrivileges - does it need admin / make substantial system changes?
# Licensing alone never makes High; High needs technical-danger evidence.

$FmkRiskHelpTitle = 'Risk ratings'

$FmkRiskHelpText = @'
Risk ratings explain the potential consequences of using a tool. They consider technical safety, system changes, privacy, source trustworthiness, and licensing or service-policy concerns. They are guidance, not a guarantee of safety.

LOW: Limited or well-understood risk. Well-understood functionality, limited and predictable system changes, credible source with reasonable maintenance, no identified significant security concern. Licensing or service-policy concerns may still exist even when technical risk is low.

MEDIUM: Meaningful caveats or elevated risk. Significant system modifications, admin operations with meaningful consequences, changes that could break software or affect stability, unclear maintenance or limited verification, substantial privacy implications, or uncertain compatibility or licensing consequences.

HIGH: Significant risk or serious consequences. Unverified executables from suspicious sources, tools that disable security controls, destructive operations without safeguards, credential theft or covert surveillance, high-impact modifications that could compromise integrity, or serious well-supported evidence of malicious behavior. A licensing concern alone never makes High.

Untagged entries need no warning: ordinary installs from verified sources with predictable behavior.
'@

$FmkRiskLevelOrder = @{ None = 0; Low = 1; Medium = 2; High = 3 }

function Get-FmkRiskTooltip {
    <#
        Multi-line tooltip for a badge or tile: level, short reason, then any
        dimensions that are set, plus evidence and review date. Returns "" for
        untagged entries so callers can skip the tooltip.
    #>
    param([Parameter(Mandatory)]$Meta)

    $risk = [string]$Meta.Risk
    if (-not $risk -or $risk -eq 'None' -or $risk -eq 'Normal') { return '' }

    $lines = @("Risk: $risk")
    if ($Meta.RiskReason) { $lines += [string]$Meta.RiskReason }
    if ($Meta.RiskTechnical) { $lines += "Technical: $($Meta.RiskTechnical)" }
    if ($Meta.RiskSource) { $lines += "Source: $($Meta.RiskSource)" }
    if ($Meta.RiskPrivacy) { $lines += "Privacy: $($Meta.RiskPrivacy)" }
    if ($Meta.RiskLicense) { $lines += "Policy: $($Meta.RiskLicense)" }
    if ($Meta.RiskPrivileges) { $lines += "Privileges: $($Meta.RiskPrivileges)" }
    if ($Meta.RiskEvidence) { $lines += "Ref: $($Meta.RiskEvidence)" }
    if ($Meta.RiskReviewed) { $lines += "Reviewed: $($Meta.RiskReviewed)" }
    return ($lines -join "`n")
}

function Test-FmkRiskBadgeVisible {
    <#
        Whether a risk level draws a badge. Pure rule, shared by the badge
        builder so the condition lives in exactly one place: Low, Medium, and
        High are badged; None (and anything unknown) is not.
    #>
    param([string]$Risk)
    return ($Risk -eq 'Low' -or $Risk -eq 'Medium' -or $Risk -eq 'High')
}

function Get-FmkRiskShort {
    <# One-line "LEVEL - reason" for queue summaries and log lines. #>
    param([Parameter(Mandatory)]$Meta)
    if ($Meta.RiskReason) { return "$($Meta.Risk): $($Meta.RiskReason)" }
    return [string]$Meta.Risk
}

function Test-FmkRiskNeedsConfirm {
    <# True for High entries: queue must require explicit confirmation. #>
    param([Parameter(Mandatory)]$Meta)
    return ([string]$Meta.Risk -eq 'High')
}

function Show-FmkRiskHelp {
    <#
        Risk-rating explainer dialog. Runs on the UI thread via the shared
        confirm dialog so no new window chrome is needed.
    #>
    $text = $FmkRiskHelpText
    if ($global:sync -and $global:sync.FmkWindow) {
        [void](Show-FmkConfirmDialog -Message $text -Title $FmkRiskHelpTitle -ConfirmText 'OK' -CancelText 'Close')
        return
    }
    Write-Host $text
}
