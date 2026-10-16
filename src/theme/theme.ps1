# ==============================================================================
# UI THEME TOKENS
# ==============================================================================
# Palette keys here must match the percent-wrapped placeholders in
# src/xaml/ui.xaml. Every brush in that file is declared as Color="%%FmkBg%%" and
# so on; Get-FmkThemedXaml substitutes the hex values into the XAML text before it
# is handed to XamlReader.Parse, which is what makes this the single place the
# colors live.
#
# Why the substitution happens before Parse rather than as a ResourceDictionary
# write afterwards: see the comment above the palette in ui.xaml. Short version
# -- writing a brush under a key a live {DynamicResource} points at throws, and
# the brushes are frozen once resolved, so they cannot be recolored in place
# either. There is no safe way to swap the palette on a window that is already up.
#
# Values are taken from WinUtil's own theme table,
#   config/themes.json   (github.com/ChrisTitusTech/winutil)
# and each key names the WinUtil key it was copied from. Keys marked "derived"
# have no upstream equivalent: WinUtil paints its window and its panels one flat
# colour and separates them with borders, which does not carry over to a
# card-based layout, so those sit between the real values.
#
# Metrics come from the same file and live in src/xaml/ui.xaml: base font Arial
# 12, header Consolas 16, corner radius 2 on buttons (WinUtil is nearly square),
# 5 on cards and text boxes, button height 25, header row 50.

$FmkThemes = @{
    Dark  = [ordered]@{
        # --- surfaces (WinUtil config/themes.json, Dark) ----------------------
        FmkBg          = '#232629'   # MainBackgroundColor
        FmkBgAlt       = '#1D2024'   # sidebar/header panel -- derived, WinUtil
        #                              paints these the same as the background
        FmkCard        = '#2A2E33'   # card surface -- between MainBackgroundColor
        #                              and AppInstallHighlightedColor
        FmkCardHover   = '#3C3C3C'   # AppInstallHighlightedColor
        FmkBorder      = '#2F373D'   # BorderColor
        # --- text ------------------------------------------------------------
        FmkFg          = '#F7F7F7'   # MainForegroundColor
        FmkFgMuted     = '#A9B1B9'   # derived
        FmkLabel       = '#5BDCFF'   # LabelboxForegroundColor
        # --- actions ---------------------------------------------------------
        FmkAccent      = '#1E3747'   # ButtonBackgroundColor
        FmkAccentHover = '#3B4252'   # ButtonBackgroundMouseoverColor
        FmkSelect      = '#5E81AC'   # ButtonBackgroundSelectedColor
        FmkAccentFg    = '#F7F7F7'   # ButtonForegroundColor
        # --- log pane --------------------------------------------------------
        FmkLogBg       = '#1B1E21'   # derived
        # --- busy overlay ----------------------------------------------------
        FmkOverlay     = '#B31D2024'   # BgAlt at ~70% opacity; dims the body
        #                              while an exclusive task runs
        # --- risk ------------------------------------------------------------
        FmkRiskLow     = '#6EFF72'   # ProgressBarForegroundColor
        FmkRiskMed     = '#FFC83D'   # ProgressBarWarningColor
        FmkRiskHigh    = '#FF6B6B'   # ProgressBarErrorColor
    }
    Light = [ordered]@{
        FmkBg          = '#F7F7F7'   # MainBackgroundColor
        FmkBgAlt       = '#EFEFEF'
        FmkCard        = '#FFFFFF'
        FmkCardHover   = '#CFCFCF'   # AppInstallHighlightedColor
        FmkBorder      = '#C4C8CC'   # BorderColor is #232629 at 0.2 opacity;
        #                              pre-blended so it stays one flat token
        FmkFg          = '#232629'   # MainForegroundColor
        FmkFgMuted     = '#5A5F66'
        FmkLabel       = '#1F6F8B'   # LabelboxForegroundColor is #5BDCFF, which
        #                              is unreadable on white; darkened instead
        FmkAccent      = '#F5F5F5'   # ButtonBackgroundColor
        FmkAccentHover = '#C2C2C2'   # ButtonBackgroundMouseoverColor
        FmkSelect      = '#D8DDE2'
        FmkAccentFg    = '#232629'   # ButtonForegroundColor
        FmkLogBg       = '#FFFFFF'
        FmkOverlay     = '#B3EFEFEF'   # BgAlt at ~70% opacity; dims the body
        FmkRiskLow     = '#1A7F37'
        FmkRiskMed     = '#B36A00'   # ProgressBarWarningColor
        FmkRiskHigh    = '#D13438'   # ProgressBarErrorColor
    }
}

function Get-FmkThemedXaml {
    <#
        Returns the window XAML with every percent-wrapped placeholder replaced
        by the hex value for the requested theme. Call this immediately before
        XamlReader.Parse; the window is then born with its final colors and
        nothing is ever written into Window.Resources at runtime.

        Both runtime approaches were tried against this window and both fail:

          * Writing a brush under a key that a live {DynamicResource} points at
            makes WPF synchronously re-resolve that reference. The re-resolution
            throws
                '#FF1B1B1F' is not a valid value for property 'Background'
            for the Window's own Background/Foreground. Writing a key nothing
            references succeeds, so it is the invalidation that breaks, not the
            value being written.
          * Mutating the existing brush's Color instead does not dodge it: a
            brush is frozen as soon as a {DynamicResource} resolves it onto a
            live element, and a frozen brush's Color is read-only.

        A missing token is a build-time mistake, so this throws rather than
        letting XamlReader report it as a syntax error somewhere unrelated.
    #>
    param(
        [Parameter(Mandatory)][string]$Xaml,
        [Parameter(Mandatory)][ValidateSet('Dark', 'Light')][string]$Theme
    )

    $palette = $FmkThemes[$Theme]
    $result = $Xaml

    foreach ($key in $palette.Keys) {
        $result = $result.Replace("%%$key%%", $palette[$key])
    }

    $leftover = @([regex]::Matches($result, '%%([A-Za-z0-9_]+)%%') |
            ForEach-Object { $_.Groups[1].Value } |
            Select-Object -Unique)

    if ($leftover.Count -gt 0) {
        throw "ui.xaml uses palette token(s) the '$Theme' theme does not define: $($leftover -join ', ')"
    }

    return $result
}

function Get-FmkBrush {
    <# Convenience lookup for colors the registry supplies at runtime (risk badges). #>
    param([Parameter(Mandatory)][string]$Key)
    $window = $null
    if ($global:sync) { $window = $global:sync.FmkWindow }
    if ($window -and $window.Resources.Contains($Key)) {
        return $window.Resources[$Key]
    }
    return [System.Windows.Media.Brushes]::Gray
}
