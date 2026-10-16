#Requires -Version 5.1
<#
.SYNOPSIS
    FreeMixKit smoke tests: metadata, risk, selection, progress, support, build.
.DESCRIPTION
    Mostly static (parses sources and the built w.ps1 the way Compile.ps1
    does), plus a headless runtime section that dot-sources the built file in
    library mode and exercises the real selection, filter, progress, risk, and
    diagnostics functions with mock sync state and real (never shown) WPF
    buttons. No window is ever displayed. Run from the repo root:
        powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Fmk.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Root = Split-Path -Parent $PSScriptRoot
$Src = Join-Path $Root 'src'
$failures = New-Object System.Collections.ArrayList
$passes = 0

function Assert-Test {
    param([bool]$Condition, [string]$Name, [string]$Detail = '')
    if ($Condition) {
        $script:passes++
        Write-Host "  PASS $Name" -ForegroundColor Green
    }
    else {
        [void]$script:failures.Add($Name)
        Write-Host "  FAIL $Name $Detail" -ForegroundColor Red
    }
}

function Read-FileText([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { throw "Missing file: $Path" }
    return (Get-Content -LiteralPath $Path -Raw)
}

Write-Host 'FreeMixKit tests' -ForegroundColor Cyan

# 1. Build output exists and parses clean --------------------------------------
$built = Join-Path $Root 'w.ps1'
Assert-Test (Test-Path -LiteralPath $built) 'build output exists'
if (Test-Path -LiteralPath $built) {
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        (Resolve-Path -LiteralPath $built).Path, [ref]$null, [ref]$parseErrors)
    Assert-Test (@($parseErrors).Count -eq 0) 'built w.ps1 parses clean'
    $builtText = Read-FileText $built
    Assert-Test ($builtText -notmatch '#>>[A-Z]+<<#') 'no unreplaced build tokens'
    foreach ($fn in @('function Register-Module', 'function Invoke-FmkSoftwareQueue',
            'function Toggle-FmkSoftwareSelection', 'function Get-FmkRiskTooltip',
            'function Show-FmkRiskHelp', 'function Initialize-FmkUi',
            'function Confirm-FmkQueueRisks', 'function Apply-FmkFilter',
            'function Test-FmkRiskBadgeVisible', 'function Set-FmkProgress',
            'function Get-FmkProgressRequest', 'function New-FmkAboutPanel',
            'function Get-FmkIssuesUrl', 'function Open-FmkIssuesPage',
            'function Protect-FmkLogSecrets', 'function Export-FmkDiagnosticLog',
            'function Invoke-ExternalCommand', 'function Write-FmkStreamText',
            'function Format-FmkStreamRecord', 'function Write-FmkStreamRecord',
            'function Get-FmkObservedPercent', 'function Get-FmkQueueOverallPercent',
            'function Update-FmkObservedProgress', 'function Set-FmkTaskComplete',
            'function Test-FmkStreamEncoding', 'function Read-FmkStreamBytes',
            'function Convert-FmkBytesToText',
            'function Get-FmkSystemReportData', 'function Format-FmkSysTextReport',
            'function ConvertTo-FmkSysJsonReport', 'function Read-FmkCim',
            'function Get-FmkSysDisplayValue', 'function Write-FmkSysReportFile')) {
        Assert-Test ($builtText.Contains($fn)) "built file contains $fn"
    }
    # The report lives in the generated FILE, never in the UI: none of the
    # removed dashboard machinery may return.
    foreach ($gone in @('New-FmkSystemReportPanel', 'Start-FmkSystemReport',
            'Show-FmkSystemReport', 'Show-FmkReportReady',
            'Invoke-FmkReportRunspace', 'Copy-FmkSystemReport',
            'Export-FmkSystemReport', 'ActReportReady', 'NoAutoReport')) {
        Assert-Test (-not $builtText.Contains($gone)) "removed UI helper stays out: $gone"
    }
}

# 2. Sources parse clean --------------------------------------------------------
$allSources = @(Get-ChildItem -LiteralPath $Src -Filter '*.ps1' -File -Recurse)
$badParse = @()
foreach ($f in $allSources) {
    $errs = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$errs)
    if (@($errs).Count -gt 0) { $badParse += $f.Name }
}
Assert-Test ($badParse.Count -eq 0) 'all src files parse clean' ($badParse -join ', ')

# 3. Risk metadata --------------------------------------------------------------
$registry = Read-FileText (Join-Path $Src 'registry.ps1')
$allowed = @('None', 'Low', 'Medium', 'High')
$riskHits = @([regex]::Matches($registry, '(?ms)^Register-Module\s+"([^"]+)".*?\$Modules\["[^"]+"\] "(.*?)"\s*@\{(.*?)\n\}') |
    ForEach-Object { @{ Key = $_.Groups[1].Value; Risk = $_.Groups[2].Value; Block = $_.Groups[3].Value } })
Assert-Test ($riskHits.Count -gt 0) 'registry registrations found'
$badRisk = @($riskHits | Where-Object { $allowed -notcontains $_.Risk })
Assert-Test ($badRisk.Count -eq 0) 'risk values are None/Low/Medium/High'
$noReason = @($riskHits | Where-Object {
        ($_.Risk -eq 'Medium' -or $_.Risk -eq 'High') -and ($_.Block -notmatch '(?m)^\s*RiskReason\s*=')
    } | ForEach-Object { $_.Key })
Assert-Test ($noReason.Count -eq 0) 'Medium/High entries set RiskReason' ($noReason -join ', ')
$noConfirm = @($riskHits | Where-Object {
        ($_.Risk -eq 'High') -and ($_.Block -notmatch '(?m)^\s*ConfirmMessage\s*=')
    } | ForEach-Object { $_.Key })
Assert-Test ($noConfirm.Count -eq 0) 'High entries set ConfirmMessage' ($noConfirm -join ', ')

# Requested risk-tag states (explicit source of truth): Winhance Low,
# .NET and Notepad++ untagged (None), SpotX and Legcord Medium. Unrelated
# entries must not move: WinUtil stays Low, GenP stays High with its
# licensing-vs-technical split documented in its block.
function Get-RiskHit([string]$Key) {
    # -NoEnumerate: a bare return would unroll a single hit, turning
    # $result.Count into the hashtable key count instead of the hit count.
    $out = @($riskHits | Where-Object { $_.Key -eq $Key })
    Write-Output $out -NoEnumerate
}
$winhance = Get-RiskHit 'Winhance'
Assert-Test (($winhance.Count -eq 1) -and ($winhance[0].Risk -eq 'Low')) 'Winhance rated Low'
$dotnet = Get-RiskHit 'Dev_DotNet'
Assert-Test (($dotnet.Count -eq 1) -and ($dotnet[0].Risk -eq 'None')) '.NET untagged (None)'
Assert-Test (($dotnet.Count -eq 1) -and ($dotnet[0].Block -notmatch 'RiskReason')) '.NET carries no risk detail'
$npp = Get-RiskHit 'Dev_NotepadPlusPlus'
Assert-Test (($npp.Count -eq 1) -and ($npp[0].Risk -eq 'None')) 'Notepad++ untagged (None)'
Assert-Test (($npp.Count -eq 1) -and ($npp[0].Block -notmatch 'RiskReason')) 'Notepad++ carries no risk detail'
$spotx = Get-RiskHit 'SpotX'
Assert-Test (($spotx.Count -eq 1) -and ($spotx[0].Risk -eq 'Medium')) 'SpotX rated Medium'
$legcord = Get-RiskHit 'Legcord'
Assert-Test (($legcord.Count -eq 1) -and ($legcord[0].Risk -eq 'Medium')) 'Legcord rated Medium'
$winutil = Get-RiskHit 'WinUtil'
Assert-Test (($winutil.Count -eq 1) -and ($winutil[0].Risk -eq 'Low')) 'WinUtil unchanged (Low)'
$genp = Get-RiskHit 'AdobeGenP'
Assert-Test (($genp.Count -eq 1) -and ($genp[0].Risk -eq 'High')) 'GenP unchanged (High)'
if ($genp.Count -eq 1) {
    Assert-Test ($genp[0].Block -match 'licens' -and $genp[0].Block -match 'Technical') 'GenP separates licensing from technical risk'
}
# Ordinary tools stay untagged: at least one winget catalog entry is None.
$brave = Get-RiskHit 'App_Brave'
Assert-Test (($brave.Count -eq 1) -and ($brave[0].Risk -eq 'None')) 'ordinary catalog entries untagged (Brave=None)'

# 3b. NanaZip catalog replacement -------------------------------------------------
$nana = Get-RiskHit 'App_NanaZip'
Assert-Test (($nana.Count -eq 1) -and ($nana[0].Risk -eq 'None')) 'NanaZip registered untagged'
Assert-Test (($nana.Count -eq 1) -and ($nana[0].Block -match 'Logo\s*=\s*"App_NanaZip"')) 'NanaZip points at its own logo'
Assert-Test (@($riskHits | Where-Object { $_.Key -eq 'Dev_SevenZip' }).Count -eq 0) '7-Zip registration removed'
$moduleText = ''
foreach ($mp in @(Get-ChildItem -LiteralPath (Join-Path $Src 'modules') -Filter '*.ps1' -File)) {
    $moduleText += (Get-Content -LiteralPath $mp.FullName -Raw)
}
Assert-Test ($moduleText.Contains('$Modules["App_NanaZip"]')) 'NanaZip module defined'
Assert-Test ($moduleText.Contains('M2Team.NanaZip')) 'NanaZip uses verified Winget ID'
Assert-Test ($moduleText -notmatch '\$Modules\["Dev_SevenZip"\]') '7-Zip module removed'
Assert-Test ($registry -notmatch '7-Zip') 'no 7-Zip references left in registry'

# 4. Selection ordering (pure ArrayList semantics the queue relies on) ----------
$sel = New-Object System.Collections.ArrayList
foreach ($k in @('App_Brave', 'App_Telegram', 'App_VLC', 'App_LocalSend')) { [void]$sel.Add($k) }
[void]$sel.Remove('App_Telegram')
[void]$sel.Add('App_Telegram')
Assert-Test (($sel -join ',') -eq 'App_Brave,App_VLC,App_LocalSend,App_Telegram') 'selection order is insertion order'
# Snapshot isolation: later toggles must not mutate a running queue.
$snapshot = @($sel)
[void]$sel.Add('App_Steam')
Assert-Test (($snapshot -join ',') -eq 'App_Brave,App_VLC,App_LocalSend,App_Telegram') 'queue snapshot isolated from later toggles'

# 5. Search and filtering (same normalization as Apply-FmkFilter) ---------------
function Test-Match([string]$Label, [string]$Description, [string]$Term) {
    $hay = ("$Label $Description").ToLowerInvariant()
    $t = $Term.Trim().ToLowerInvariant()
    return (($t -eq '') -or $hay.Contains($t))
}
Assert-Test (Test-Match 'Brave Browser' 'Privacy-focused web browser.' 'brave') 'search exact'
Assert-Test (Test-Match 'Brave Browser' 'Privacy-focused web browser.' 'BRAV') 'search case-insensitive'
Assert-Test (Test-Match 'Telegram Desktop' 'Desktop messaging app.' 'gram') 'search partial'
Assert-Test (-not (Test-Match 'VLC Media Player' 'Plays almost any media format.' 'brave')) 'search non-match'
Assert-Test (Test-Match 'VLC Media Player' 'Plays almost any media format.' '') 'search empty matches all'

# 6. Icon-resource resolution ----------------------------------------------------
$assetText = Read-FileText (Join-Path $Src 'assets\logos.ps1')
$logoKeys = @([regex]::Matches($assetText, "(?m)^\s*'([A-Za-z0-9_]+)'\s*=\s*@\{") |
    ForEach-Object { $_.Groups[1].Value })
$referenced = @([regex]::Matches($registry, '(?m)^\s*Logo\s*=\s*"([^"]+)"') |
    ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
$missing = @($referenced | Where-Object { $logoKeys -notcontains $_ })
Assert-Test ($missing.Count -eq 0) 'every Logo resolves to an asset entry' ($missing -join ', ')
# No letter/initial substitutes in the Software grid except documented Mono lows.
$monoRefs = @($referenced | Where-Object {
        $m = [regex]::Match($assetText, "(?ms)'$([regex]::Escape($_))'\s*=\s*@\{([^}]*)\}")
        $m.Success -and $m.Groups[2].Value -match 'Mono'
    })
Assert-Test ($monoRefs.Count -le 1) 'grid has no initial-letter substitutes (except documented Mono)' ($monoRefs -join ', ')

# 7. Categories consolidated ------------------------------------------------------
$groups = @([regex]::Matches($registry, '(?m)^\s*Group\s*=\s*"([^"]+)"') |
    ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
$expectedGroups = @('Browsers', 'Messaging', 'Media', 'File Sharing & Downloads',
    'Development', 'Video & Image Tools', 'Gaming', 'System & Maintenance', 'Productivity')
foreach ($g in $expectedGroups) {
    Assert-Test ($groups -contains $g) "category group present: $g"
}

# 8. Window sizing -----------------------------------------------------------------
$xaml = Read-FileText (Join-Path $Src 'xaml\ui.xaml')
$wm = [regex]::Match($xaml, 'Height="(\d+)" Width="(\d+)"')
Assert-Test ($wm.Success) 'window dimensions declared'
if ($wm.Success) {
    $h = [int]$wm.Groups[1].Value; $w = [int]$wm.Groups[2].Value
    Assert-Test (($w -ge 1140 -and $w -le 1240) -and ($h -ge 780 -and $h -le 840)) 'window slightly larger, not excessive' "${w}x${h}"
}
$bootText = Read-FileText (Join-Path $Src 'functions\07-boot.ps1')
Assert-Test ($bootText.Contains('WorkArea')) 'window clamped to work area'
Assert-Test ($xaml.Contains('InstallSelectedButton')) 'Install Selected button in XAML'
Assert-Test ($xaml.Contains('x:Name="ProgressPercent"')) 'visible percentage label in XAML'
Assert-Test ($xaml.Contains('x:Name="BusyOverlay"')) 'loading overlay in XAML'
Assert-Test ($xaml.Contains('x:Name="BusyOverlayText"')) 'overlay status text in XAML'
Assert-Test ($xaml.Contains('x:Name="BusyOverlaySpinner"')) 'overlay spinner in XAML'
Assert-Test ($xaml.Contains('x:Name="TaskCancelButton"')) 'overlay Cancel button in XAML'
Assert-Test ($xaml -notmatch 'x:Name="QueueCancelButton"') 'superseded action-bar Cancel removed'
Assert-Test ($xaml.Contains('x:Name="SideBar"') -and $xaml.Contains('x:Name="MainContent"')) 'lockable sidebar/content regions named'
Assert-Test ($xaml.Contains('x:Name="page_About"') -and $xaml.Contains('x:Name="cards_About"')) 'About page in XAML'
Assert-Test ($xaml -notmatch 'x:Name="page_Links"') 'old Links page gone from XAML'
Assert-Test ($xaml.Contains('RiskHelpButton') -or $xaml.Contains('GlobalRiskHelpButton')) 'risk help reachable from UI'
# Selection/focus decoupling: the tile border must be driven by the selected
# trigger alone. A blanket IsKeyboardFocused recolor would always win and
# paint stale highlights on focused-but-deselected tiles.
$tileStyle = ''
$tm = [regex]::Match($xaml, '(?s)<Style x:Key="TileButton".*?(<Style x:Key="TileFocusVisual"|<!-- Category chips)')
if ($tm.Success) { $tileStyle = $tm.Value }
Assert-Test ($tileStyle -ne '') 'TileButton style found'
Assert-Test ($tileStyle -notmatch '<Trigger Property="IsKeyboardFocused"') 'no focus recolor inside TileButton triggers'
Assert-Test ($xaml.Contains('x:Key="TileFocusVisual"')) 'tile focus adorner style in XAML'
$cardsText = Read-FileText (Join-Path $Src 'functions\06-cards.ps1')
Assert-Test ($cardsText.Contains("FindResource('TileFocusVisual')")) 'tiles use the focus adorner'
Assert-Test ($cardsText.Contains("'About'") -and $cardsText -notmatch "'Links'") 'sidebar metadata renamed Links to About'

# 9. Queue safety properties (static) ----------------------------------------------
$queueText = Read-FileText (Join-Path $Src 'functions\09-queue.ps1')
Assert-Test ($queueText.Contains('QueueRunning')) 'queue has re-entrancy guard'
Assert-Test ($queueText.Contains('QueueCancel')) 'queue supports cancel'
Assert-Test ($queueText.Contains('Confirm-FmkQueueRisks')) 'queue reviews Medium/High before install'
Assert-Test ($queueText -notmatch 'Bypass.*UAC|Disable.*UAC|Set-ItemProperty.*EnableLUA') 'queue never bypasses security controls'

# 9b. Window/taskbar icon ------------------------------------------------------------
$iconAsset = Join-Path $Src 'assets\icon.ps1'
Assert-Test (Test-Path -LiteralPath $iconAsset) 'vendored window icon asset exists'
if (Test-Path -LiteralPath $iconAsset) {
    $iconText = Read-FileText $iconAsset
    $b64m = [regex]::Match($iconText, "\`$FmkAppIconB64 = '([A-Za-z0-9+/=]+)'")
    Assert-Test ($b64m.Success) 'vendored icon carries base64 ICO bytes'
    if ($b64m.Success) {
        $icoBytes = [System.Convert]::FromBase64String($b64m.Groups[1].Value)
        $isIco = ($icoBytes.Length -gt 6) -and ($icoBytes[0] -eq 0) -and ($icoBytes[1] -eq 0) -and
            ([System.BitConverter]::ToUInt16($icoBytes, 2) -eq 1) -and
            ([System.BitConverter]::ToUInt16($icoBytes, 4) -ge 1)
        Assert-Test $isIco 'vendored icon is a valid multi-image ICO'
    }
}
$templateText = Read-FileText (Join-Path $Src 'w.template.ps1')
Assert-Test ($templateText.Contains('FmkAppIconB64')) 'template falls back to vendored icon'
Assert-Test ($bootText.Contains('BitmapCacheOption')) 'window icon decode releases the file'
if (Test-Path -LiteralPath $built) {
    Assert-Test ($builtText.Contains('$FmkAppIconB64')) 'built file embeds the window icon'
}

# 9c. Source encoding (About-mojibake regression) ---------------------------------------
# A UTF-8 bullet once shipped double-encoded because the build read BOM-less
# sources with the system ANSI codepage. Sources stay ASCII-only so no
# BOM, host codepage, script parser, or font has to agree on anything.
$nonAsciiFiles = @()
foreach ($sf in @(Get-ChildItem -LiteralPath $Src -Filter '*.ps1' -File -Recurse)) {
    $sb = [System.IO.File]::ReadAllBytes($sf.FullName)
    foreach ($byte in $sb) {
        if ($byte -gt 127) { $nonAsciiFiles += $sf.Name; break }
    }
}
Assert-Test ($nonAsciiFiles.Count -eq 0) 'sources are ASCII-only' ($nonAsciiFiles -join ', ')
if (Test-Path -LiteralPath $built) {
    # Byte-level mojibake scan, written ASCII-only on purpose: embedding the
    # mojibake characters themselves would make THIS file non-ASCII and
    # reintroduce the host-codepage parsing hazard under test. C3 A2 E2 82
    # AC is the old double-encoded bullet signature.
    $builtBytes = [System.IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $built).Path)
    $hasBadBullet = $false
    for ($i = 0; $i -le $builtBytes.Length - 6; $i++) {
        if ($builtBytes[$i] -eq 0xC3 -and $builtBytes[$i + 1] -eq 0xA2 -and
            $builtBytes[$i + 2] -eq 0xE2 -and $builtBytes[$i + 3] -eq 0x82 -and
            $builtBytes[$i + 4] -eq 0xAC) { $hasBadBullet = $true; break }
    }
    Assert-Test (-not $hasBadBullet) 'built file carries no double-encoded bullet'
    Assert-Test ($builtText.Contains('Software catalog: search')) 'feature text intact'
    Assert-Test ($builtText.Contains('"- $f"')) 'feature bullets use ASCII hyphens'
}

# 10. Headless runtime: real functions, mock UI ---------------------------------------
# Dot-sources the BUILT file in library mode (no window, no elevation) and
# drives the actual selection, filter, progress, risk, and diagnostics code.
# Buttons are real WPF objects that are never shown; sync state is a plain
# hashtable. Skipped loudly if the environment cannot do even that.
$runtimeAvailable = $false
try {
    . (Join-Path $Root 'w.ps1') -FmkLibraryMode
    Set-StrictMode -Version Latest
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase -ErrorAction Stop
    $probe = New-Object System.Windows.Controls.Button -ErrorAction Stop
    $runtimeAvailable = $true
}
catch {
    Write-Host ("  SKIP runtime section unavailable: " + $_.Exception.Message) -ForegroundColor Yellow
}

if ($runtimeAvailable) {
    function New-MockTile([string]$Key, [string]$Haystack) {
        $t = New-Object System.Windows.Controls.Button
        $t.Uid = $Key
        $t.DataContext = $Haystack
        $t.Tag = $null
        return $t
    }
    function New-MockSync {
        return @{
            SoftwareSelection    = New-Object System.Collections.ArrayList
            QueueRunning         = $false
            Busy                 = $false
            BusyLabel            = $null
            CardSections         = @{}
            InstallSelectedButton = $null
            ClearSelectionButton = $null
            SelectionStatus      = $null
            ActiveCategory       = 'Software'
            SoftwareGroup        = $null
            SoftwareCount        = $null
            FilterTerms          = @{}
            FmkWindow            = $null
            EmptyPageNotice      = $null
            SearchBox            = $null
        }
    }

    # Badge visibility rule: Low, Medium, High shown; None/unknown hidden.
    Assert-Test (Test-FmkRiskBadgeVisible -Risk 'Low') 'badge visible for Low'
    Assert-Test (Test-FmkRiskBadgeVisible -Risk 'Medium') 'badge visible for Medium'
    Assert-Test (Test-FmkRiskBadgeVisible -Risk 'High') 'badge visible for High'
    Assert-Test (-not (Test-FmkRiskBadgeVisible -Risk 'None')) 'badge hidden when untagged'
    Assert-Test (-not (Test-FmkRiskBadgeVisible -Risk '')) 'badge hidden for empty risk'
    Assert-Test (-not (Test-FmkRiskBadgeVisible -Risk 'low ')) 'badge hidden for unknown value'

    # Selection: first click selects, second deselects, visuals agree.
    $msync = New-MockSync
    $tileA = New-MockTile 'App_VLC' 'vlc media player plays almost any media format'
    $tileB = New-MockTile 'App_Brave' 'brave browser privacy-focused web browser'
    $mockHead = [pscustomobject]@{ Visibility = $null }
    $msync.CardSections = @{ Software = @(@{ Heading = $mockHead; Body = ([pscustomobject]@{ Children = @($tileA, $tileB) }) }) }
    [void](Toggle-FmkSoftwareSelection -ActionKey 'App_VLC' -Sync $msync)
    Assert-Test ((@(Get-FmkSoftwareSelection -Sync $msync) -join ',') -eq 'App_VLC') 'first click selects'
    Assert-Test ($tileA.Tag -eq 'selected') 'selected border marker set on first click'
    [void](Toggle-FmkSoftwareSelection -ActionKey 'App_VLC' -Sync $msync)
    Assert-Test ((@(Get-FmkSoftwareSelection -Sync $msync).Count) -eq 0) 'second click deselects'
    Assert-Test ($null -eq $tileA.Tag) 'border marker cleared immediately on deselect'
    # Multi-select keeps insertion order; switching cards leaves no stale mark.
    [void](Toggle-FmkSoftwareSelection -ActionKey 'App_VLC' -Sync $msync)
    [void](Toggle-FmkSoftwareSelection -ActionKey 'App_Brave' -Sync $msync)
    [void](Toggle-FmkSoftwareSelection -ActionKey 'App_VLC' -Sync $msync)
    Assert-Test ((@(Get-FmkSoftwareSelection -Sync $msync) -join ',') -eq 'App_Brave') 'toggling another card leaves no stale selection'
    Assert-Test (($null -eq $tileA.Tag) -and ($tileB.Tag -eq 'selected')) 'visuals agree after switching cards'
    [void](Clear-FmkSoftwareSelection -Sync $msync)
    Assert-Test ((@(Get-FmkSoftwareSelection -Sync $msync).Count) -eq 0 -and ($null -eq $tileB.Tag)) 'clear resets state and visuals'
    # Busy/queue locks selection controls.
    $msync.Busy = $true
    [void](Toggle-FmkSoftwareSelection -ActionKey 'App_VLC' -Sync $msync)
    Assert-Test ((@(Get-FmkSoftwareSelection -Sync $msync).Count) -eq 0) 'selection locked while busy'
    $msync.Busy = $false

    # Filtering preserves selection visuals and hides only non-matches.
    [void](Toggle-FmkSoftwareSelection -ActionKey 'App_Brave' -Sync $msync)
    $msync.SearchBox = [pscustomobject]@{ Text = 'vlc' }
    $msync.EmptyPageNotice = [pscustomobject]@{ Text = ''; Visibility = $null }
    $mockWindow = New-Object PSObject
    $mockWindow | Add-Member -MemberType ScriptMethod -Name FindName -Value { param($n) return $null }
    $msync.FmkWindow = $mockWindow
    Apply-FmkFilter -Sync $msync
    Assert-Test ($tileB.Tag -eq 'selected') 'selection marker survives filtering'
    Assert-Test (($tileA.Visibility -eq [System.Windows.Visibility]::Visible) -and ($tileB.Visibility -eq [System.Windows.Visibility]::Collapsed)) 'filter shows matches, hides the rest'
    $msync.SearchBox = [pscustomobject]@{ Text = '' }
    Apply-FmkFilter -Sync $msync
    Assert-Test (($tileA.Visibility -eq [System.Windows.Visibility]::Visible) -and ($tileB.Visibility -eq [System.Windows.Visibility]::Visible)) 'clearing search restores all cards'

    # Progress normalization: measured numbers only, never invented.
    $r50 = Get-FmkProgressRequest -Percent 50
    Assert-Test (($r50.Mode -eq 'Determinate') -and ($r50.Value -eq 50)) 'measured percent stays determinate'
    $rHi = Get-FmkProgressRequest -Percent 150
    Assert-Test (($rHi.Mode -eq 'Determinate') -and ($rHi.Value -eq 100)) 'percent clamped to 100'
    $rLo = Get-FmkProgressRequest -Percent -5
    Assert-Test (($rLo.Mode -eq 'Determinate') -and ($rLo.Value -eq 0)) 'percent clamped to 0'
    Assert-Test ((Get-FmkProgressRequest -Percent $null).Mode -eq 'Indeterminate') 'missing percent is indeterminate'
    Assert-Test ((Get-FmkProgressRequest -Percent 'abc').Mode -eq 'Indeterminate') 'unparseable percent is indeterminate'
    Assert-Test ((Get-FmkProgressRequest -Percent ([double]::NaN)).Mode -eq 'Indeterminate') 'NaN percent is indeterminate'

    # About/support: official URL, redaction, export behavior.
    Assert-Test ((Get-FmkIssuesUrl) -eq 'https://github.com/catsmoker/FreeMixKit/issues') 'issues URL is the official page'
    $repoHits = @(Select-String -LiteralPath (Join-Path $Src 'w.template.ps1') -Pattern 'github.com/catsmoker/FreeMixKit' -SimpleMatch)
    Assert-Test ($repoHits.Count -eq 1) 'repo URL defined once in the template' ($repoHits.Count)
    $dupIssues = @(Get-ChildItem -LiteralPath $Src -Filter '*.ps1' -File -Recurse | Select-String -Pattern 'FreeMixKit/issues' -SimpleMatch)
    Assert-Test ($dupIssues.Count -eq 0) 'issues URL derived, never duplicated' ($dupIssues.Count)
    $dirty = 'user password=hunter2 and api_key: ABC123 ok' + "`n" + 'Authorization: Bearer xyzzy-token-99' + "`n" + 'Installing VLC...'
    $clean = Protect-FmkLogSecrets -Text $dirty
    Assert-Test (($clean -notmatch 'hunter2') -and ($clean -notmatch 'ABC123') -and ($clean -notmatch 'xyzzy-token-99')) 'secrets redacted from diagnostics'
    Assert-Test (($clean -match '\[REDACTED\]') -and ($clean -match 'Installing VLC')) 'structure preserved after redaction'
    $pem = '-----BEGIN RSA PRIVATE KEY-----' + "`n" + 'MIIBOgIBAAJBAK3' + "`n" + '-----END RSA PRIVATE KEY-----'
    Assert-Test ((Protect-FmkLogSecrets -Text $pem) -notmatch 'MIIBOgIBAAJBAK3') 'private key body redacted'
    Assert-Test ((Protect-FmkLogSecrets -Text '') -eq '') 'empty input sanitizes to empty'
    $emptyExport = Export-FmkDiagnosticLog -SessionLog @() -Results @() -SavePath ''
    Assert-Test ($null -eq $emptyExport) 'no logs yields a message, not an empty report'
    $tmpLog = Join-Path ([System.IO.Path]::GetTempPath()) 'fmk-test-diag.txt'
    if (Test-Path -LiteralPath $tmpLog) { Remove-Item -LiteralPath $tmpLog -Force }
    $sampleResults = @([pscustomobject]@{ Timestamp = '2026-10-16 12:00:00'; Module = 'App_VLC'; Status = 'Success'; DurationSec = 1.5; Message = 'Completed' })
    $saved = Export-FmkDiagnosticLog -SessionLog @('[12:00:01] Installing VLC', 'note password=hunter2') -Results $sampleResults -SavePath $tmpLog
    Assert-Test (($saved -eq $tmpLog) -and (Test-Path -LiteralPath $tmpLog)) 'diagnostics export writes the chosen file'
    if (Test-Path -LiteralPath $tmpLog) {
        $body = Get-Content -LiteralPath $tmpLog -Raw
        Assert-Test (($body -match 'App_VLC') -and ($body -match 'FreeMixKit diagnostic log')) 'export carries task context'
        Assert-Test (($body -notmatch 'hunter2') -and ($body -match '\[REDACTED\]')) 'export is sanitized'
        Remove-Item -LiteralPath $tmpLog -Force
    }

    # Observed percentage: parsed only from genuine output text, never made up.
    Assert-Test ((Get-FmkObservedPercent -Text 'Downloading 42%') -eq 42) 'percent parsed from output'
    Assert-Test ((Get-FmkObservedPercent -Text 'done 100%') -eq 100) 'percent caps at 100'
    Assert-Test ((Get-FmkObservedPercent -Text '10% .. 42%') -eq 42) 'last percentage on the line wins'
    Assert-Test ((Get-FmkObservedPercent -Text '150%') -eq 100) 'overshoot clamped, not rejected'
    Assert-Test ($null -eq (Get-FmkObservedPercent -Text 'no numbers here')) 'no percentage yields null'
    Assert-Test ($null -eq (Get-FmkObservedPercent -Text '')) 'empty line yields null'
    # Queue overall math: completed tasks plus the current measured fraction.
    Assert-Test ((Get-FmkQueueOverallPercent -Done 2 -Total 5 -ItemPercent 50) -eq 50) 'overall blends completed and current'
    Assert-Test ((Get-FmkQueueOverallPercent -Done 2 -Total 5 -ItemPercent 0) -eq 40) 'overall floor is completed-only'
    Assert-Test ((Get-FmkQueueOverallPercent -Done 0 -Total 1 -ItemPercent 100) -eq 100) 'single-item completion reaches 100'
    Assert-Test ((Get-FmkQueueOverallPercent -Done 0 -Total 0 -ItemPercent 50) -eq 0) 'zero total never divides'
    # Muted lifecycle vs full-color tool output.
    Assert-Test ($FmkLogColorKeys['Muted'] -eq 'FmkFgMuted') 'muted lifecycle color registered'
    $mutedSites = @(Get-ChildItem -LiteralPath (Join-Path $Src 'functions') -Filter '*.ps1' -File -Recurse |
        Select-String -Pattern "'Muted'" -SimpleMatch)
    Assert-Test ($mutedSites.Count -ge 4) 'lifecycle sites log muted' ($mutedSites.Count)
    # Forwarding honors queue context, or parks nothing headless.
    $usync = @{ FmkWindow = 'fake-window'; QueueRunning = $true; QueueTotal = 5; QueueDone = 2; ObservedItemSent = -1 }
    $savedSync = $global:sync
    try {
        $global:sync = $usync
        Update-FmkObservedProgress -ItemPercent 50
        Assert-Test (($usync.PendingProgressMode -eq 'Determinate') -and ($usync.PendingProgressValue -eq 50)) 'observed percent forwarded as overall'
        $usync.QueueRunning = $false
        Update-FmkObservedProgress -ItemPercent 7
        Assert-Test (($usync.PendingProgressMode -eq 'Determinate') -and ($usync.PendingProgressValue -eq 7)) 'lone task forwards its own percent'
    }
    finally { $global:sync = $savedSync }

    # Task manager on a real parsed window (never shown): overlay, bar,
    # percentage, lock, cancel, and terminal states.
    $wiredOk = $false
    $tsync = $null
    try {
        $wxm = [regex]::Match($builtText, "(?s)\`$FmkXaml = @'\r\n(.*?)\r\n'@")
        $wthemed = Get-FmkThemedXaml -Xaml $wxm.Groups[1].Value -Theme 'Dark'
        $wwindow = [System.Windows.Markup.XamlReader]::Parse($wthemed)
        $tsync = New-FmkSyncState -Window $wwindow -Xaml $wxm.Groups[1].Value
        $global:sync = $tsync
        $wiredOk = $true
    }
    catch {
        Write-Host ('  SKIP task-manager behavior: ' + $_.Exception.Message) -ForegroundColor Yellow
    }
    if ($wiredOk) {
        try {
            $vis = [System.Windows.Visibility]
            $tsync.PendingBusy = $true
            $tsync.PendingBusyText = 'Installing X (2 of 5)'
            $tsync.PendingProgressMode = 'Determinate'
            $tsync.PendingProgressValue = 40
            Set-FmkBusyNow
            Assert-Test ($tsync.BusyOverlay.Visibility -eq $vis::Visible) 'overlay shows while busy'
            Assert-Test (($tsync.ProgressBar.Visibility -eq $vis::Visible) -and ($tsync.ProgressBar.Value -eq 40) -and (-not $tsync.ProgressBar.IsIndeterminate)) 'determinate bar reflects measured value'
            Assert-Test ($tsync.ProgressPercent.Text -eq '40%') 'percentage number visible'
            Assert-Test ($tsync.BusyOverlayText.Text -eq 'Installing X (2 of 5)') 'determinate overlay shows no waiting filler'
            Assert-Test ((-not $tsync.SideBar.IsEnabled) -and (-not $tsync.MainContent.IsEnabled)) 'sidebar and pages locked'
            $tsync.PendingProgressMode = 'Indeterminate'
            Set-FmkBusyNow
            Assert-Test ($tsync.ProgressBar.IsIndeterminate -and ($tsync.ProgressPercent.Text -eq '')) 'indeterminate shows no number'
            Assert-Test ($tsync.BusyOverlayText.Text -match 'waiting for installer progress information') 'indeterminate names the missing evidence'
            $tsync.QueueRunning = $true
            Set-FmkBusyNow
            Assert-Test ($tsync.TaskCancelButton.Visibility -eq $vis::Visible) 'cancel offered for the queue'
            $tsync.QueueRunning = $false
            Set-FmkBusyNow
            Assert-Test ($tsync.TaskCancelButton.Visibility -eq $vis::Collapsed) 'no cancel for lone tasks'
            $tsync.PendingBusy = $false
            $tsync.PendingProgressMode = 'Complete'
            Set-FmkBusyNow
            Assert-Test (($tsync.BusyOverlay.Visibility -eq $vis::Collapsed) -and ($tsync.ProgressBar.Value -eq 100) -and ($tsync.ProgressPercent.Text -eq '100%')) 'success completes the bar'
            Assert-Test ($tsync.SideBar.IsEnabled -and $tsync.MainContent.IsEnabled) 'lock released after success'
            $tsync.PendingProgressMode = 'Failed'
            $tsync.PendingProgressValue = 42
            Set-FmkBusyNow
            Assert-Test (($tsync.ProgressBar.Value -eq 42) -and ($tsync.ProgressPercent.Text -eq '42%')) 'failure freezes at last measured value'
            Assert-Test ($tsync.ProgressBar.Foreground -eq $tsync.FmkWindow.Resources['FmkRiskHigh']) 'failure shown in red, not 100%'
            $tsync.PendingProgressMode = 'None'
            $tsync.PendingProgressValue = 0
            Set-FmkBusyNow
            Assert-Test (($tsync.ProgressBar.Visibility -eq $vis::Collapsed) -and ($tsync.ProgressPercent.Text -eq '')) 'idle hides bar and number'
        }
        finally { $global:sync = $null }
    }

    # About tab preservation: existing links intact under the new name.
    $aboutKeys = @($riskHits | Where-Object {
            ($_.Block -match '(?m)^\s*Category\s*=\s*"About"') } | ForEach-Object { $_.Key })
    Assert-Test (($aboutKeys -contains 'OpenGitHub') -and ($aboutKeys -contains 'OpenWebsite') -and ($aboutKeys -contains 'AddShortcut')) 'About tab preserves repository, website, shortcut'
    Assert-Test (@($riskHits | Where-Object { $_.Block -match '(?m)^\s*Category\s*=\s*"Links"' }).Count -eq 0) 'no registrations left on Links'

    # Output pipeline: the real execution path, observed via the session log.
    # Trailing blank from the file's final newline must not count as a line,
    # or every slice starts one line late and drops the first new entry.
    function Remove-TrailingBlanks([array]$Lines) {
        $n = $Lines.Count
        while ($n -gt 0 -and $Lines[$n - 1] -eq '') { $n-- }
        if ($n -eq 0) { return @() }
        return @($Lines[0..($n - 1)])
    }
    function Get-SessionLineCount {
        try {
            $rawText = [System.IO.File]::ReadAllText($SessionLogPath, [System.Text.Encoding]::UTF8)
            return @(Remove-TrailingBlanks @($rawText -split "`r?`n")).Count
        }
        catch { return 0 }
    }
    function Read-NewSessionLines([int]$BeforeCount) {
        # Plain enumerated return on purpose: callers join the result, and a
        # nested array would stringify to "System.Object[]". (The -NoEnumerate
        # trick used by Get-RiskHit exists because that caller needs .Count.)
        try {
            $rawText = [System.IO.File]::ReadAllText($SessionLogPath, [System.Text.Encoding]::UTF8)
            $all = @(Remove-TrailingBlanks @($rawText -split "`r?`n"))
            if ($all.Count -le $BeforeCount) { return @() }
            return @($all[$BeforeCount..($all.Count - 1)])
        }
        catch { return @() }
    }
    $logWritable = (Get-SessionLineCount) -gt 0 -or (Test-Path -LiteralPath $SessionLogPath)
    if (-not $logWritable) {
        Write-Host '  SKIP output-pipeline tests: session log unavailable' -ForegroundColor Yellow
    }
    else {
        # 0. Real winget download through the product path (download-only: the
        # only side effect is a temp-dir file, removed afterwards).
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            $dlDir = Join-Path ([System.IO.Path]::GetTempPath()) 'fmk-test-dl'
            if (-not (Test-Path -LiteralPath $dlDir)) { [void](New-Item -ItemType Directory -Path $dlDir) }
            $dc0 = Get-SessionLineCount
            [void](Invoke-ExternalCommand -FilePath 'winget.exe' -Arguments @('download', '--id', 'LocalSend.LocalSend', '--exact', '--accept-package-agreements', '--accept-source-agreements', '--download-directory', $dlDir))
            $td = @(Read-NewSessionLines $dc0) -join "`n"
            Assert-Test ($td -match 'Found LocalSend') 'real download shows discovery output'
            Assert-Test ($td -match 'Downloading https?://') 'real download shows source URL'
            Assert-Test ($td -match 'Successfully verified installer hash') 'real download shows verification'
            Assert-Test ($td -match 'Installer downloaded:') 'real download shows result path'
            Remove-Item -LiteralPath $dlDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        else {
            Write-Host '  SKIP real download check: winget not installed' -ForegroundColor Yellow
        }

        # 1. Multiple stdout lines arrive in order.
        $c0 = Get-SessionLineCount
        [void](Invoke-ExternalCommand -FilePath 'cmd.exe' -Arguments @('/c', 'echo alpha-line-1 & echo alpha-line-2 & echo alpha-line-3'))
        $t1 = @(Read-NewSessionLines $c0) -join "`n"
        Assert-Test ($t1 -match 'alpha-line-1[\s\S]*alpha-line-2[\s\S]*alpha-line-3') 'process stdout streams in order'
        Assert-Test ($t1 -notmatch 'Stream capture degraded') 'clean run emits no capture warning'

        # 2. stderr is captured and a nonzero exit throws with its code.
        $c0 = Get-SessionLineCount
        $threw = $false
        $throwMsg = ''
        try {
            [void](Invoke-ExternalCommand -FilePath 'cmd.exe' -Arguments @('/c', 'echo out-marker & echo err-marker 1>&2 & exit 3'))
        }
        catch { $threw = $true; $throwMsg = $_.Exception.Message }
        $t2 = @(Read-NewSessionLines $c0) -join "`n"
        Assert-Test $threw 'nonzero exit throws'
        Assert-Test ($throwMsg -match 'exit 3') 'exit code reported, not just started'
        Assert-Test (($t2 -match 'out-marker') -and ($t2 -match 'err-marker')) 'stdout and stderr both captured'

        # 3. PowerShell warning/error/success streams through a real module run.
        $Modules['__TestStreams__'] = {
            $VerbosePreference = 'Continue'
            Write-Output 'stream-line-ok'
            Write-Warning 'stream-warn-marker'
            Write-Verbose 'stream-verbose-marker'
            Write-Information 'stream-info-marker' -InformationAction Continue
        }
        $c0 = Get-SessionLineCount
        $sres = Invoke-ModuleAction -ActionKey '__TestStreams__'
        $t3 = @(Read-NewSessionLines $c0) -join "`n"
        Assert-Test ($sres.Status -eq 'Success') 'warning-carrying module still succeeds'
        Assert-Test ($t3 -match 'stream-line-ok') 'success output streamed, not discarded'
        Assert-Test ($t3 -match 'WARNING: stream-warn-marker') 'warning stream preserved with tag'
        Assert-Test ($t3 -match 'VERBOSE: stream-verbose-marker') 'emitted verbose stream captured'
        Assert-Test ($t3 -match 'stream-info-marker') 'information stream captured'
        $Modules['__TestErr__'] = {
            Write-Output 'before-err-line'
            Write-Error 'stream-err-marker'
        }
        $c0 = Get-SessionLineCount
        # Terminating errors propagate out of Invoke-ModuleAction by design;
        # the runspace and single-task callers convert them into Failed
        # results. The lines already streamed must survive regardless.
        $ethrew = $false
        $emsg = ''
        try { [void](Invoke-ModuleAction -ActionKey '__TestErr__') }
        catch { $ethrew = $true; $emsg = $_.Exception.Message }
        $t3e = @(Read-NewSessionLines $c0) -join "`n"
        Assert-Test $ethrew 'terminating error propagates to the caller'
        Assert-Test ($emsg -match 'stream-err-marker') 'error text carried, not swallowed'
        Assert-Test ($t3e -match 'before-err-line') 'output before the error is kept'
        $Modules.Remove('__TestStreams__')
        $Modules.Remove('__TestErr__')

        # 4. Unicode survives the pipeline. Compared with an exact in-process
        # string (built from char codes, so this file stays ASCII-only): the
        # RTL/text-capture layer of some hosts renders non-ASCII as boxes,
        # which must not be mistaken for pipeline corruption.
        $uniExpect = 'uni-h' + [char]0xE9 + 'llo-' + [char]0x2713 + '-done'
        $uniCmd = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; '$uniExpect'"
        $uniB64 = [System.Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes('& {' + $uniCmd + '}'))
        $c0 = Get-SessionLineCount
        [void](Invoke-ExternalCommand -FilePath 'powershell.exe' -Arguments @('-NoProfile', '-EncodedCommand', $uniB64))
        $t4 = @(Read-NewSessionLines $c0) -join "`n"
        Assert-Test ($t4.Contains($uniExpect)) 'unicode text byte-identical end to end'

        # 5. Incremental output arrives during the run, not only at the end. The
        # session log stamps every line at arrival time (1s resolution), so a
        # 12 x 500ms tick train must span several seconds in the record; an
        # end-of-run dump would stamp every line (near-)identically. Anchored
        # matches: tick-probe-1 is a substring of 10/11/12.
        $tickCmd = '1..12 | ForEach-Object { Write-Output ("tick-probe-$_"); Start-Sleep -Milliseconds 500 }'
        $tickB64 = [System.Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes('& {' + $tickCmd + '}'))
        $c0 = Get-SessionLineCount
        [void](Invoke-ExternalCommand -FilePath 'powershell.exe' -Arguments @('-NoProfile', '-EncodedCommand', $tickB64))
        $t5 = @(Read-NewSessionLines $c0)
        $firstTick = @($t5 | Where-Object { $_ -match 'tick-probe-1\s*$' } | Select-Object -First 1)
        $lastTick = @($t5 | Where-Object { $_ -match 'tick-probe-12\s*$' } | Select-Object -First 1)
        Assert-Test (($firstTick.Count -eq 1) -and ($lastTick.Count -eq 1)) 'all incremental lines recorded'
        if (($firstTick.Count -eq 1) -and ($lastTick.Count -eq 1)) {
            # Session lines are stamped "[yyyy-MM-dd HH:mm:ss] [TYPE] text".
            $m1 = [regex]::Match($firstTick[0], '(\d\d):(\d\d):(\d\d)\]')
            $m2 = [regex]::Match($lastTick[0], '(\d\d):(\d\d):(\d\d)\]')
            Assert-Test ($m1.Success -and $m2.Success) 'arrival timestamps parsable'
            $s1 = [int]$m1.Groups[1].Value * 3600 + [int]$m1.Groups[2].Value * 60 + [int]$m1.Groups[3].Value
            $s2 = [int]$m2.Groups[1].Value * 3600 + [int]$m2.Groups[2].Value * 60 + [int]$m2.Groups[3].Value
            $spanSecs = $s2 - $s1
            if ($spanSecs -lt 0) { $spanSecs += 86400 }
            Assert-Test ($spanSecs -ge 3) 'first/last lines separated in time (streamed)' ("span=${spanSecs}s")
        }

        # 6. Volume completes without hanging or losing lines. Kept at a size
        # that proves multi-tick capture without flooding the host console.
        Write-Host '  (bulk run starts)' -ForegroundColor DarkGray
        $bulkStart = Get-Date
        $c0 = Get-SessionLineCount
        [void](Invoke-ExternalCommand -FilePath 'cmd.exe' -Arguments @('/c', 'for /L %i in (1,1,250) do @echo bulk-line-%i'))
        $bulkSecs = ((Get-Date) - $bulkStart).TotalSeconds
        Write-Host ("  (bulk run done in " + [Math]::Round($bulkSecs, 1) + "s)") -ForegroundColor DarkGray
        $t6 = @(Read-NewSessionLines $c0) -join "`n"
        $bulkHits = @([regex]::Matches($t6, 'bulk-line-\d+')).Count
        Assert-Test ($bulkHits -ge 240) 'bulk output volume completes' ("got $bulkHits/250 in $([Math]::Round($bulkSecs, 1))s")
        Assert-Test (($t6 -match 'bulk-line-1\b[\s\S]*bulk-line-250')) 'bulk output keeps order'
        # Burst cap holds back overflow in order instead of dropping it.
        $oldCap = $FmkStreamPollLines
        $FmkStreamPollLines = 5
        try {
            $capState = @{ Fragment = ''; LogType = 'Info' }
            $twelve = @('s1', 's2', 's3', 's4', 's5', 's6', 's7', 's8') -join "`n"
            $cc0 = Get-SessionLineCount
            Write-FmkStreamText -State $capState -Text $twelve
            $ct1 = @(Read-NewSessionLines $cc0) -join "`n"
            Assert-Test (($ct1 -match 's1') -and ($ct1 -notmatch 's8')) 'burst cap paces output across ticks'
            Write-FmkStreamText -State $capState -Text '' -Final:$true
            $ct2 = @(Read-NewSessionLines $cc0) -join "`n"
            Assert-Test ($ct2 -match 's8') 'held-back segments flush in order'
        }
        finally { $FmkStreamPollLines = $oldCap }

        # 6b. Stream encoding: BOM and NUL-interleaved prefixes select UTF-16,
        # plain bytes stay UTF-8, short prefixes wait for more evidence.
        $bom16 = Test-FmkStreamEncoding -Lead ([byte[]]@(0xFF, 0xFE, 0x41, 0x00))
        Assert-Test (($bom16.Enc -eq 'utf16') -and ($bom16.Skip -eq 2)) 'UTF-16 BOM detected and skipped'
        $bom8 = Test-FmkStreamEncoding -Lead ([byte[]]@(0xEF, 0xBB, 0xBF, 0x41))
        Assert-Test (($bom8.Enc -eq 'utf8') -and ($bom8.Skip -eq 3)) 'UTF-8 BOM detected and skipped'
        $leText = Test-FmkStreamEncoding -Lead ([byte[]]@(0x41, 0x00, 0x42, 0x00))
        Assert-Test (($leText.Enc -eq 'utf16') -and ($leText.Skip -eq 0)) 'NUL-interleaved prefix selects UTF-16'
        $plainText = Test-FmkStreamEncoding -Lead ([byte[]]@(0x41, 0x42, 0x43, 0x44))
        Assert-Test (($plainText.Enc -eq 'utf8') -and ($plainText.Skip -eq 0)) 'plain bytes stay UTF-8'
        $cjkUtf8 = Test-FmkStreamEncoding -Lead ([byte[]]@(0xE4, 0xB8, 0xAD, 0xE6))
        Assert-Test ($cjkUtf8.Enc -eq 'utf8') 'multibyte UTF-8 not mistaken for UTF-16'
        Assert-Test ($null -eq (Test-FmkStreamEncoding -Lead ([byte[]]@(0x41, 0x00)))) 'short prefix waits for evidence'
        $shortFinal = Test-FmkStreamEncoding -Lead ([byte[]]@(0x41, 0x00)) -Final
        Assert-Test ($shortFinal.Enc -eq 'utf8') 'end of stream falls back to UTF-8'

        # 6c. Real UTF-16 process end to end: cmd /u emits UTF-16LE, which must
        # arrive as exact text lines with no NUL interleaving and no CJK
        # mojibake from a single-byte misread.
        $c0 = Get-SessionLineCount
        [void](Invoke-ExternalCommand -FilePath 'cmd.exe' -Arguments @('/u', '/c', 'echo u16-alpha-line & echo u16-err-line 1>&2'))
        $tu = @(Read-NewSessionLines $c0) -join "`n"
        Assert-Test ($tu -match 'u16-alpha-line') 'UTF-16 stdout decoded exactly'
        Assert-Test ($tu -match 'u16-err-line') 'UTF-16 stderr decoded exactly'
        Assert-Test ($tu -notmatch [string][char]0) 'no NUL bytes leak into logged text'

        # 7. Progress records update status, never the log.
        $Modules['__TestProg__'] = {
            Write-Progress -Activity 'prog-activity-marker' -Status 'half' -PercentComplete 50
            Write-Progress -Activity 'prog-activity-marker' -Status 'done' -PercentComplete 100
            Write-Progress -Activity 'prog-activity-marker' -Completed
        }
        $c0 = Get-SessionLineCount
        $pres = Invoke-ModuleAction -ActionKey '__TestProg__'
        $t7 = @(Read-NewSessionLines $c0) -join "`n"
        Assert-Test ($pres.Status -eq 'Success') 'progress-carrying module succeeds'
        Assert-Test ($t7 -notmatch 'prog-activity-marker') 'progress ticks do not spam the log'
        $Modules.Remove('__TestProg__')

        # Stream-record mapping for every PowerShell stream.
        try { Write-Error 'errx-marker' -ErrorAction Stop } catch { $capErr = $_ }
        $fe = Format-FmkStreamRecord -Record $capErr
        Assert-Test (($fe.Kind -eq 'Error') -and ($fe.Type -eq 'Error') -and ($fe.Text -match 'errx-marker')) 'error record mapped'
        $wr = $null
        Write-Warning 'warnx-marker' -WarningVariable wr 3>$null
        $fw = Format-FmkStreamRecord -Record (@($wr)[0])
        Assert-Test (($fw.Kind -eq 'Warning') -and ($fw.Type -eq 'Warn') -and ($fw.Text -match 'warnx-marker')) 'warning record mapped'
        $fv = Format-FmkStreamRecord -Record (New-Object System.Management.Automation.VerboseRecord('verbx-marker'))
        Assert-Test (($fv.Kind -eq 'Verbose') -and ($fv.Text -match 'verbx-marker')) 'verbose record mapped'
        $fd = Format-FmkStreamRecord -Record (New-Object System.Management.Automation.DebugRecord('debugx-marker'))
        Assert-Test (($fd.Kind -eq 'Debug') -and ($fd.Text -match 'debugx-marker')) 'debug record mapped'
        $fi = Format-FmkStreamRecord -Record (New-Object System.Management.Automation.InformationRecord('infox-marker', 'tagx'))
        Assert-Test (($fi.Kind -eq 'Information') -and ($fi.Text -match 'infox-marker')) 'information record mapped'
        $pr = New-Object System.Management.Automation.ProgressRecord(1, 'act-marker', 'stat-marker')
        $pr.PercentComplete = 42
        $fp = Format-FmkStreamRecord -Record $pr
        Assert-Test (($fp.Kind -eq 'Progress') -and ($fp.Percent -eq 42) -and ($fp.Activity -eq 'act-marker')) 'progress record mapped'
        $fs = Format-FmkStreamRecord -Record 'plain-line'
        Assert-Test (($fs.Kind -eq 'Output') -and ($fs.Text -eq 'plain-line')) 'plain output passes through'
        $find = Format-FmkStreamRecord -Record "`t  indented-line  "
        Assert-Test ($find.Text -eq "`t  indented-line") 'indentation kept, trailing space trimmed'

        # One real, side-effect-free FreeMixKit-adjacent operation: winget itself.
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            $c0 = Get-SessionLineCount
            [void](Invoke-ExternalCommand -FilePath 'winget.exe' -Arguments @('--version'))
            $tw = @(Read-NewSessionLines $c0) -join "`n"
            Assert-Test ($tw -match 'v?\d+\.\d+') 'real winget output shown, not just status'
        }
        else {
            Write-Host '  SKIP real winget check: winget not installed' -ForegroundColor Yellow
        }

        # 10b. System report on mocked data (deterministic; real hardware is
        # verified separately on a live machine, see the final report).
        Assert-Test ((Format-FmkBytes 0) -eq '0 B') 'zero bytes stay zero'
        Assert-Test (((Format-FmkBytes 1536) -eq '1.5 KB') -or ((Format-FmkBytes 1536) -eq '1,5 KB')) 'bytes to KB'
        Assert-Test ((Format-FmkBytes ([double]2 * 1GB)) -eq '2.0 GB') 'bytes to GB'
        Assert-Test ((Format-FmkBytes -1) -eq 'Unavailable') 'negative bytes unavailable'
        Assert-Test ((Format-FmkBytes 'abc') -eq 'Unavailable') 'garbage bytes unavailable'
        Assert-Test ((Format-FmkGigahertz 3600) -eq '3.60 GHz') 'megahertz to gigahertz'
        Assert-Test ((Format-FmkGigahertz 0) -eq 'Unavailable') 'zero frequency unavailable, not zero'
        Assert-Test ((Format-FmkWmiDate '20241015123000.000000-240') -match '^\d{4}-\d\d-\d\d \d\d:\d\d$') 'WMI date parsed to local time'
        Assert-Test ((Format-FmkWmiDate 'garbage') -eq 'Unavailable') 'bad WMI date unavailable'
        Assert-Test ((ConvertFrom-FmkWmiDate (Get-Date)) -is [datetime]) 'DateTime passes through'
        $bootAgo = (Get-Date).AddDays(-2).AddHours(-3).AddMinutes(-5)
        Assert-Test ((Format-FmkUptime $bootAgo) -match '2 days') 'uptime days shown'
        Assert-Test ((Format-FmkUptime ((Get-Date).AddHours(1))) -eq 'Unavailable') 'future boot unavailable'
        Assert-Test ((Get-FmkMemoryTypeName 0x1A) -eq 'DDR4') 'DDR4 code mapped'
        Assert-Test ((Get-FmkMemoryTypeName 0x22) -eq 'DDR5') 'DDR5 code mapped'
        Assert-Test ((Get-FmkMemoryTypeName 0xFF) -match 'Unknown') 'unknown memory code keeps its code'
        Assert-Test ((Get-FmkCpuArchName 9) -eq 'x64') 'arch code mapped'
        Assert-Test ((Get-FmkCpuArchName 12) -eq 'ARM64') 'ARM64 code mapped'
        Assert-Test ((Get-FmkOpStatusName 2) -eq 'OK') 'operational status OK mapped'
        Assert-Test ((Get-FmkOpStatusName 5) -eq 'Predictive failure') 'predictive failure mapped'
        Assert-Test ((ConvertFrom-FmkWslLine "D`0e`0f`0a`0u`0l`0t`0") -eq 'Default') 'WSL UTF-16 line cleaned'
        Assert-Test ((ConvertFrom-FmkWslLine '  plain  ') -eq 'plain') 'WSL plain line trimmed'

        $privRow = @{ Label = 'Disk 0 serial number'; Value = 'SERIAL-123'; Sensitivity = 'Private' }
        Assert-Test ((Get-FmkSysDisplayValue $privRow) -eq '[hidden]') 'private hidden on screen'
        Assert-Test ((Get-FmkSysDisplayValue $privRow -ForExport:$true) -eq '[withheld]') 'private withheld from default export'
        Assert-Test ((Get-FmkSysDisplayValue $privRow -IncludeSensitive:$true) -eq 'SERIAL-123') 'private shown when included'
        $idRow = @{ Label = 'Computer name (potentially sensitive)'; Value = 'PC-1'; Sensitivity = 'Identifying' }
        Assert-Test ((Get-FmkSysDisplayValue $idRow) -eq 'PC-1') 'identifying shown on screen'
        Assert-Test ((Get-FmkSysDisplayValue $idRow -ForExport:$true) -eq '[withheld]') 'identifying withheld from default export'
        Assert-Test ((Get-FmkSysDisplayValue $idRow -IncludeSensitive:$true -ForExport:$true) -eq 'PC-1') 'identifying included when asked'

        $mockSec = New-FmkSysSection -Key 'T' -Title 'Test Section'
        Add-FmkSysRow $mockSec 'Present' 'yes'
        Add-FmkSysRow $mockSec 'Zero stays' 0
        Add-FmkSysRow $mockSec 'Missing' $null
        Add-FmkSysRow $mockSec 'Secret' 's3cr3t' -Sensitivity 'Private'
        Add-FmkSysWarning $mockSec 'first notice'
        Add-FmkSysWarning $mockSec 'second notice'
        $mockVals = @($mockSec.Rows | ForEach-Object { $_.Value })
        Assert-Test ($mockVals -contains 'yes') 'row keeps value'
        Assert-Test ($mockVals -contains '0') 'row keeps numeric zero'
        Assert-Test ($mockVals -contains 'Unavailable') 'row maps missing to Unavailable'
        Assert-Test ((Find-FmkSysRow $mockSec 'Present') -eq 'yes') 'row lookup finds value'
        Assert-Test ((Find-FmkSysRow $mockSec 'Nope') -eq '') 'row lookup misses cleanly'

        $mockStorage = New-FmkSysSection -Key 'Storage' -Title 'Storage'
        Add-FmkSysRow $mockStorage 'Scope note' 'Physical disks and logical volumes are listed separately below.'
        Add-FmkSysRow $mockStorage 'Disk 0' 'MOCK-NVME-1 -- 512.0 GB, NVMe / SSD (boot disk)'
        Add-FmkSysRow $mockStorage 'Disk 0 serial number' 'NVME-SERIAL-1' -Sensitivity 'Private'
        Add-FmkSysRow $mockStorage 'Disk 0 volume C:' '(Windows) NTFS Local fixed disk 100.0 GB free of 200.0 GB (50% free) (system drive)'
        Add-FmkSysRow $mockStorage 'Disk 1' 'MOCK-HDD-2 -- 1.0 TB, SATA / HDD'
        Add-FmkSysRow $mockStorage 'Disk 1 volume D:' '(Data) NTFS Local fixed disk 0 B free of 500.0 GB (0% free)'
        Add-FmkSysWarning $mockStorage 'SMART status unavailable (often requires administrator rights).'
        $mockReport = @{
            GeneratedAt = (Get-Date '2026-01-02 03:04:05')
            IsAdmin     = $false
            Sections    = @($mockStorage)
        }
        $mockText = Format-FmkSysTextReport -Report $mockReport -IncludeSensitive:$false
        Assert-Test (($mockText -match 'MOCK-NVME-1') -and ($mockText -match 'MOCK-HDD-2')) 'text covers every physical disk'
        Assert-Test (($mockText -match 'Disk 0 volume C:') -and ($mockText -match 'Disk 1 volume D:')) 'volumes attach to their disks'
        Assert-Test ($mockText -match '0 B free') 'zero free space stays zero'
        Assert-Test (($mockText -match '\[hidden\]') -and ($mockText -notmatch 'NVME-SERIAL-1')) 'default text redacts serials'
        $mockExport = Format-FmkSysTextReport -Report $mockReport -IncludeSensitive:$false -ForExport:$true
        Assert-Test (($mockExport -match '\[withheld\]') -and ($mockExport -notmatch 'NVME-SERIAL-1')) 'export file redacts serials'
        Assert-Test ($mockText -match 'SMART status unavailable') 'section warnings render'
        $mockFull = Format-FmkSysTextReport -Report $mockReport -IncludeSensitive:$true
        Assert-Test ($mockFull -match 'NVME-SERIAL-1') 'included text restores serials'
        $mockJson = (ConvertTo-FmkSysJsonReport -Report $mockReport -IncludeSensitive:$false | ConvertTo-Json -Depth 6)
        Assert-Test (($mockJson -match 'MOCK-HDD-2') -and ($mockJson -match '\[withheld\]') -and ($mockJson -notmatch 'NVME-SERIAL-1')) 'JSON redacts like text'
        $mockJsonFull = (ConvertTo-FmkSysJsonReport -Report $mockReport -IncludeSensitive:$true | ConvertTo-Json -Depth 6)
        Assert-Test ($mockJsonFull -match 'NVME-SERIAL-1') 'JSON includes when asked'

        # 10c. Report file encoding: UTF-8 with BOM round-trips exactly, so no
        # reader can mistake the bytes for UTF-16 (the reported CJK mojibake).
        # Unicode sample built from char codes to keep this file ASCII-only.
        $encSample = 'Free Mix Kit System Report' + "`r`n" +
            'Audio device 1: Intel' + [char]0xAE + ' Smart Sound' + "`r`n" +
            'Monitor 1: ' + [char]0x4E2D + [char]0x6587 + ' U278' + "`r`n" +
            'Note: 100% -- (ok), price: 1,299 EUR' + "`r`n"
        $encPath = Join-Path ([System.IO.Path]::GetTempPath()) 'fmk-enc-test.txt'
        if (Test-Path -LiteralPath $encPath) { Remove-Item -LiteralPath $encPath -Force }
        Write-FmkSysReportFile -Path $encPath -Text $encSample
        $encBytes = [System.IO.File]::ReadAllBytes($encPath)
        Assert-Test (($encBytes.Length -gt 3) -and ($encBytes[0] -eq 0xEF) -and ($encBytes[1] -eq 0xBB) -and ($encBytes[2] -eq 0xBF)) 'report file carries a UTF-8 BOM'
        $encBack = [System.IO.File]::ReadAllText($encPath, [System.Text.Encoding]::UTF8)
        Assert-Test ($encBack -eq $encSample) 'UTF-8 round-trip preserves text exactly'
        Assert-Test ($encBack.Contains('Free Mix Kit System Report')) 'ASCII heading intact'
        Assert-Test ($encBack.Contains([string][char]0xAE)) 'registered mark preserved'
        Assert-Test ($encBack.Contains([string][char]0x4E2D)) 'CJK text preserved'
        $noBomBody = $encBytes[3..($encBytes.Length - 1)]
        $misread = [System.Text.Encoding]::Unicode.GetString($noBomBody)
        Assert-Test ($misread -notmatch 'Free Mix Kit System Report') 'sanity: UTF-16 misread differs (why the BOM matters)'
        Remove-Item -LiteralPath $encPath -Force -ErrorAction SilentlyContinue
    }
}

# 11. Build validation ---------------------------------------------------------------
$verify = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root 'Compile.ps1') -Verify 2>&1
Assert-Test ($LASTEXITCODE -eq 0) 'Compile.ps1 -Verify passes'

# 11. PSScriptAnalyzer (optional; skipped when not installed) -------------------------
$pssa = Get-Module -ListAvailable -Name PSScriptAnalyzer | Select-Object -First 1
if ($pssa) {
    $findings = Invoke-ScriptAnalyzer -Path (Join-Path $Src 'functions\09-queue.ps1') -Severity Error, Warning |
        Where-Object { $_.RuleName -notin @('PSAvoidUsingWriteHost') }
    Assert-Test (@($findings).Count -eq 0) 'PSScriptAnalyzer clean on queue module'
}
else {
    Write-Host '  SKIP PSScriptAnalyzer not installed' -ForegroundColor Yellow
}

Write-Host ''
Write-Host "Passed: $passes; Failed: $($failures.Count)" -ForegroundColor $(if ($failures.Count -eq 0) { 'Green' } else { 'Red' })
if ($failures.Count -gt 0) {
    Write-Host 'Failures:' -ForegroundColor Red
    foreach ($f in $failures) { Write-Host "  - $f" -ForegroundColor Red }
    exit 1
}
