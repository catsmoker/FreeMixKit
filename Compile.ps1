#Requires -Version 5.1
<#
.SYNOPSIS
    Builds the distributable w.ps1 from the sources under src/.

.DESCRIPTION
    Concatenates the template, helper functions, theme, modules, registry, and
    XAML into a single self-contained script at the repository root. The root
    w.ps1 is what the README's `irm ... | iex` one-liner serves, so it must stay
    a single file.

.PARAMETER Verify
    Run all build validations and report, but do not write w.ps1.

.PARAMETER OutputPath
    Override the output file. Defaults to w.ps1 beside this script.

.EXAMPLE
    .\Compile.ps1
    .\Compile.ps1 -Verify
#>
[CmdletBinding()]
param(
    [switch]$Verify,
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Root = $PSScriptRoot
$SrcDir = Join-Path $Root 'src'

if (-not $OutputPath) { $OutputPath = Join-Path $Root 'w.ps1' }

# ------------------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------------------

function Read-SourceFile {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Missing source file: $Path"
    }
    # Explicit UTF-8, never the host default codepage: Get-Content -Raw decodes
    # a BOM-less file with the system ANSI codepage, which once double-encoded
    # a UTF-8 bullet into mojibake in the shipped script. .NET honors a BOM
    # when one is present and otherwise decodes plain UTF-8, matching the
    # UTF-8-without-BOM output written below.
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
}

function Get-SourceSet {
    <#
        Returns matching .ps1 files in deterministic (name) order.

        Callers must wrap the result in @(): PowerShell unrolls an array on the
        way out of a function, so a directory holding exactly one file would
        otherwise yield a bare FileInfo (no .Count) and an empty directory would
        yield $null. Only wrapping at the call site keeps .Count reliable.
    #>
    param(
        [Parameter(Mandatory)][string]$Directory,
        [string]$Filter = '*.ps1'
    )
    if (-not (Test-Path -LiteralPath $Directory)) { return @() }
    return @(Get-ChildItem -LiteralPath $Directory -Filter $Filter -File | Sort-Object Name | ForEach-Object { $_.FullName })
}

function Join-SourceBlocks {
    <# Labels each file with a source banner so the built file stays navigable. #>
    param([string[]]$Paths)

    $sb = New-Object System.Text.StringBuilder
    foreach ($p in $Paths) {
        $name = Split-Path -Leaf $p
        $rule = '# ' + ('-' * 76)
        [void]$sb.AppendLine($rule)
        [void]$sb.AppendLine("# source: $name")
        [void]$sb.AppendLine($rule)
        [void]$sb.AppendLine((Read-SourceFile -Path $p).TrimEnd())
        [void]$sb.AppendLine()
    }
    return $sb.ToString()
}

function Get-BuildStamp {
    $sha = 'no-git'
    try {
        $git = Get-Command git -ErrorAction SilentlyContinue
        if ($git) {
            $out = & git -C $Root rev-parse --short HEAD 2>$null
            if ($LASTEXITCODE -eq 0 -and $out) { $sha = $out.Trim() }
        }
    }
    catch { }
    return "$sha @ $((Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))"
}

function New-XamlBlock {
    <#
        Wraps the XAML in a single-quoted here-string. The XAML contains $ and "
        freely, so a double-quoted here-string would interpolate them; and any
        line beginning with '@ would terminate the here-string early and corrupt
        the build. Validate for that, and refuse rather than emit a broken file.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Xaml)

    if ([string]::IsNullOrWhiteSpace($Xaml)) { return '' }

    $offending = @()
    $lineNo = 0
    foreach ($line in ($Xaml -split "`r?`n")) {
        $lineNo++
        if ($line.TrimStart().StartsWith("'@")) {
            $offending += $lineNo
        }
    }
    if ($offending.Count -gt 0) {
        throw ("XAML contains a line starting with '@ at line(s) $($offending -join ', '). " +
            "This would terminate the embedding here-string early. Reformat those lines " +
            "(e.g. move the '@ off the start of the line) and rebuild.")
    }

    # NOTE: the opening delimiter is built with an escaped quote inside a
    # single-quoted string, and newlines are appended as real characters. Do not
    # put `r`n inside a single-quoted literal here -- it would stay literal text.
    $nl = "`r`n"
    $open = '$FmkXaml = @'''   # -> $FmkXaml = @'
    $close = "'@"               # -> '@
    return ($open + $nl + $Xaml.TrimEnd() + $nl + $close + $nl)
}

# ------------------------------------------------------------------------------
# Gather
# ------------------------------------------------------------------------------

Write-Host 'FreeMixKit build' -ForegroundColor Cyan
Write-Host "  root: $Root"

$templatePath = Join-Path $SrcDir 'w.template.ps1'
$template = Read-SourceFile -Path $templatePath

$functionPaths = @(Get-SourceSet -Directory (Join-Path $SrcDir 'functions'))
$themePaths = @(Get-SourceSet -Directory (Join-Path $SrcDir 'theme'))
$assetPaths = @(Get-SourceSet -Directory (Join-Path $SrcDir 'assets'))
$modulePaths = @(Get-SourceSet -Directory (Join-Path $SrcDir 'modules'))

$registryPath = Join-Path $SrcDir 'registry.ps1'
$xamlPath = Join-Path $SrcDir 'xaml\ui.xaml'

$registry = if (Test-Path -LiteralPath $registryPath) { Read-SourceFile -Path $registryPath } else { '' }
$xaml = if (Test-Path -LiteralPath $xamlPath) { Read-SourceFile -Path $xamlPath } else { '' }

# Concatenated only so the palette-token check below has the key names to hand.
$themeText = ''
foreach ($p in $themePaths) { $themeText += (Read-SourceFile -Path $p) }

# Concatenated for the logo-key check.
$assetText = ''
foreach ($p in $assetPaths) { $assetText += (Read-SourceFile -Path $p) }

Write-Host "  functions: $($functionPaths.Count) file(s)"
Write-Host "  theme:     $($themePaths.Count) file(s)"
Write-Host "  assets:    $($assetPaths.Count) file(s)"
Write-Host "  modules:   $($modulePaths.Count) file(s)"

if ($functionPaths.Count -eq 0) { throw "No helper functions found under $SrcDir\functions" }
if ($assetPaths.Count -eq 0) { throw "No vendored assets found under $SrcDir\assets" }
if ($modulePaths.Count -eq 0) { throw "No modules found under $SrcDir\modules" }
if (-not $registry) { throw "Missing registry: $registryPath" }

# ------------------------------------------------------------------------------
# Validate
# ------------------------------------------------------------------------------

$problems = @()

# Parse every source file with the real PowerShell parser rather than counting
# braces -- a naive brace count trips over braces inside strings and comments.
foreach ($p in @($templatePath) + @($functionPaths) + @($themePaths) + @($assetPaths) + @($modulePaths) + @($registryPath)) {
    if (-not (Test-Path -LiteralPath $p)) { continue }
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        (Resolve-Path -LiteralPath $p).Path, [ref]$null, [ref]$parseErrors)
    foreach ($e in @($parseErrors)) {
        $problems += "$(Split-Path -Leaf $p):$($e.Extent.StartLineNumber) $($e.Message)"
    }
}

# Assigning to a read-only automatic variable parses perfectly and then dies at
# runtime with
#     Cannot overwrite variable Host because it is read-only or constant.
# reported against the built w.ps1, which says nothing about which source file
# to fix. `$host` is the easy one to trip over -- PowerShell variable names are
# case-insensitive, so a local $host for "the panel these cards go in" collides
# with $Host silently until that function first runs.
$readOnlyAutoVars = '(?i)\$(Host|HOME|PID|PSHOME|PWD|Matches|MyInvocation|PSScriptRoot|PSCommandPath|ExecutionContext|ShellId|StackTrace|PSVersionTable|PSItem|PSCmdlet|PSBoundParameters|NestedPromptLevel|IsCoreCLR|IsLinux|IsMacOS|IsWindows|PSEdition|PSCulture|PSUICulture|ConsoleFileName|PSSenderInfo|PSDebugContext|true|false|null)\s*=[^=]'

foreach ($p in @($templatePath) + @($functionPaths) + @($themePaths) + @($assetPaths) + @($modulePaths) + @($registryPath)) {
    if (-not (Test-Path -LiteralPath $p)) { continue }
    $lineNo = 0
    foreach ($line in @(Get-Content -LiteralPath $p)) {
        $lineNo++
        if ($line -match $readOnlyAutoVars) {
            $problems += ('{0}:{1} assigns to the read-only automatic variable ${2}' -f
                (Split-Path -Leaf $p), $lineNo, $Matches[1])
        }
    }
}

# Every Category the registry assigns must have a matching nav page in the XAML,
# otherwise modules silently vanish from the GUI.
$xamlBlock = New-XamlBlock -Xaml $xaml

if ($xaml) {
    $declared = @([regex]::Matches($xaml, 'x:Name="page_([A-Za-z0-9]+)"') |
            ForEach-Object { $_.Groups[1].Value })
    $used = @([regex]::Matches($registry, 'Category\s*=\s*"([^"]+)"') |
            ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)

    if ($declared.Count -eq 0) {
        $problems += "XAML is present but declares no x:Name=""page_*"" elements"
    }
    foreach ($c in $used) {
        if ($declared -notcontains $c) {
            $problems += "registry uses Category '$c' but the XAML has no page_$c"
        }
    }

    # The palette is baked into the XAML text by Get-FmkThemedXaml, so a token
    # with no matching key is a startup crash rather than a build error unless
    # it is caught here.
    $tokens = @([regex]::Matches($xaml, '%%([A-Za-z0-9_]+)%%') |
            ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
    $paletteKeys = @([regex]::Matches($themeText, '(?m)^\s*(Fmk[A-Za-z0-9_]*)\s*=') |
            ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)

    if ($tokens.Count -eq 0) {
        $problems += 'ui.xaml has no %%Token%% placeholders -- the palette is still hardcoded?'
    }
    foreach ($t in $tokens) {
        if ($paletteKeys -notcontains $t) {
            $problems += "ui.xaml uses %%$t%% but no theme defines a key named $t"
        }
    }

    Write-Host "  palette:   $($tokens.Count) token(s) in XAML, $($paletteKeys.Count) key(s) in theme"

    # ui.xaml is XML before it is XAML, and the failure mode is nasty: an illegal
    # construct (a "--" inside a comment, a bare "&") builds a working w.ps1 that
    # throws from XamlReader.Parse at startup, after the console has already been
    # hidden. Catch it here instead, where the line number is still useful.
    try {
        $null = [xml]$xaml
    }
    catch {
        $problems += "ui.xaml is not well-formed XML: $($_.Exception.Message)"
    }
}

# --- vendored assets ----------------------------------------------------------
# Layer data is written by a generator, but it is checked here rather than
# trusted: a malformed entry still parses as PowerShell, and the only symptom
# would be a tile quietly drawing nothing.
if ($assetText) {
    $logoEntries = @([regex]::Matches($assetText, "(?m)^\s*'([A-Za-z0-9_]+)'\s*=\s*@\{(.+)\}\s*$") |
            ForEach-Object { @{ Key = $_.Groups[1].Value; Body = $_.Groups[2].Value.Trim() } })

    if ($logoEntries.Count -eq 0) {
        $problems += 'assets: found no logo entries (expected a $FmkLogos hashtable)'
    }

    foreach ($entry in $logoEntries) {
        $isMark = $entry.Body -match "^Color\s*=\s*'#[0-9A-Fa-f]{6}'\s*;\s*Data\s*=\s*'.+'\s*$"
        $isMono = $entry.Body -match "^Mono\s*=\s*'[^']{1,4}'\s*$"
        $isPng = $entry.Body -match "^Png\s*=\s*'[A-Za-z0-9+/=]+'\s*$"
        $isShapes = $entry.Body -match '^Shapes\s*=\s*@\(.+\)\s*$'
        if (-not ($isMark -or $isMono -or $isPng -or $isShapes)) {
            $problems += "assets: entry '$($entry.Key)' is not a brand mark (Color + Data), raster icon (Png), shape set (Shapes), or monogram (Mono)"
        }
        if ($isPng) {
            try { [void][Convert]::FromBase64String(([regex]::Match($entry.Body, "Png\s*=\s*'([^']+)'")).Groups[1].Value) }
            catch { $problems += "assets: entry '$($entry.Key)' has a Png value that is not valid base64" }
        }
    }

    # A Logo key that resolves to nothing is not an error -- the grid falls back
    # to initials -- but it is always a typo, and a typo that renders as a
    # different-looking tile is easier to catch here than in a screenshot.
    $logoKeys = @($logoEntries | ForEach-Object { $_.Key })
    $referenced = @([regex]::Matches($registry, "(?m)^\s*Logo\s*=\s*""([^""]+)""") |
            ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
    foreach ($name in $referenced) {
        if ($logoKeys -notcontains $name) {
            $problems += "registry sets Logo = '$name' but no logo entry defines that key"
        }
    }

    Write-Host "  logos:     $($logoKeys.Count) entry(s), $($referenced.Count) referenced by the registry"
}

# --- registry / module wiring -------------------------------------------------
# Register-Module takes its action as a [scriptblock] argument, and $null binds
# to that parameter happily. So a registration pointing at a $Modules key nothing
# defines -- a typo, or a module file that was deleted -- produces a card that
# renders, clicks, and then fails with "Unknown action" the first time a user
# tries it. Catch that here instead.
$moduleText = ''
foreach ($p in $modulePaths) { $moduleText += (Read-SourceFile -Path $p) }

$definedModules = @([regex]::Matches($moduleText, '(?m)^\s*\$Modules\["([^"]+)"\]\s*=') |
        ForEach-Object { $_.Groups[1].Value })
$registeredKeys = @()
foreach ($m in [regex]::Matches($registry, '(?m)^Register-Module\s+"([^"]+)"[^\r\n]*?\$Modules\["([^"]+)"\]')) {
    $registeredKeys += $m.Groups[1].Value
    if ($definedModules -notcontains $m.Groups[2].Value) {
        $problems += "registry registers '$($m.Groups[1].Value)' with `$Modules['$($m.Groups[2].Value)'], which no module defines"
    }
}

$duplicateKeys = @($registeredKeys | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
if ($duplicateKeys.Count -gt 0) {
    $problems += "registry registers the same key twice: $($duplicateKeys -join ', ')"
}

$unregistered = @($definedModules | Where-Object { $registeredKeys -notcontains $_ })
if ($unregistered.Count -gt 0) {
    Write-Host "  note:      module(s) defined but not registered (reachable only from the library): $($unregistered -join ', ')" -ForegroundColor DarkGray
}

if ($registeredKeys.Count -gt 0) {
    Write-Host "  registry:  $($registeredKeys.Count) registration(s), $($definedModules.Count) module(s) defined"
}

# --- risk metadata ------------------------------------------------------------
# Risk must be None/Low/Medium/High. Medium/High must carry a RiskReason so the
# badge tooltip is never an bare label. High must also carry a ConfirmMessage:
# single-run High modules confirm via Invoke-ModuleAction, and the Software
# queue holds its own combined review, so an unexplained High is a bug.
$allowedRisks = @('None', 'Low', 'Medium', 'High')
foreach ($m in [regex]::Matches($registry, '(?ms)^Register-Module\s+"([^"]+)".*?\$Modules\["[^"]+"\] "(.*?)"\s*@\{(.*?)\n\}')) {
    $regKey = $m.Groups[1].Value
    $risk = $m.Groups[2].Value
    $block = $m.Groups[3].Value
    if ($allowedRisks -notcontains $risk) {
        $problems += "registry '$regKey' uses Risk '$risk' (expected None, Low, Medium, or High)"
        continue
    }
    if (($risk -eq 'Medium' -or $risk -eq 'High') -and ($block -notmatch '(?m)^\s*RiskReason\s*=')) {
        $problems += "registry '$regKey' is Risk '$risk' but sets no RiskReason"
    }
    if (($risk -eq 'High') -and ($block -notmatch '(?m)^\s*ConfirmMessage\s*=')) {
        $problems += "registry '$regKey' is Risk 'High' but sets no ConfirmMessage"
    }
}

if ($problems.Count -gt 0) {
    Write-Host ''
    Write-Host 'Build validation failed:' -ForegroundColor Red
    foreach ($p in $problems) { Write-Host "  - $p" -ForegroundColor Red }
    exit 1
}

Write-Host '  validation: ok' -ForegroundColor Green

if ($Verify) {
    Write-Host ''
    Write-Host 'Verify only; no output written.' -ForegroundColor Yellow
    exit 0
}

# ------------------------------------------------------------------------------
# Assemble
# ------------------------------------------------------------------------------
# Tokens are replaced with .Replace() (literal), never -replace (regex), because
# the substituted text contains $ and \ that a regex replacement would mangle.

$stamp = Get-BuildStamp
$buildHeader = @"
# ==============================================================================
# FreeMixKit -- BUILT FILE, DO NOT EDIT DIRECTLY
# ------------------------------------------------------------------------------
# Generated by Compile.ps1 on $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
# Build: $stamp
#
# Source of truth is the src/ directory:
#   src/w.template.ps1   entry point and boot sequence
#   src/functions/*.ps1  helper + registry functions
#   src/theme/*.ps1      UI color tokens
#   src/assets/*.ps1     vendored icon path data
#   src/modules/*.ps1    module scriptblocks
#   src/registry.ps1     module registration (labels, risk, UI placement)
#   src/xaml/ui.xaml     WPF window definition
#
# Edit those, then run:  .\Compile.ps1
# ==============================================================================
"@

$output = $template
$output = $output.Replace('#>>BUILD<<#', $buildHeader)
$output = $output.Replace('#>>FUNCTIONS<<#', (Join-SourceBlocks -Paths $functionPaths).TrimEnd())
$output = $output.Replace('#>>THEME<<#', (Join-SourceBlocks -Paths $themePaths).TrimEnd())
$output = $output.Replace('#>>ASSETS<<#', (Join-SourceBlocks -Paths $assetPaths).TrimEnd())
$output = $output.Replace('#>>MODULES<<#', (Join-SourceBlocks -Paths $modulePaths).TrimEnd())
$output = $output.Replace('#>>REGISTRY<<#', $registry.TrimEnd())
$output = $output.Replace('#>>XAML<<#', $xamlBlock.TrimEnd())

# Any token left unreplaced means the template and this script disagree.
$leftover = [regex]::Matches($output, '#>>[A-Z]+<<#')
if ($leftover.Count -gt 0) {
    $names = ($leftover | ForEach-Object { $_.Value } | Select-Object -Unique) -join ', '
    throw "Unreplaced build token(s) in output: $names"
}

# Normalise to CRLF and write UTF-8 without BOM (a BOM trips up some hosts).
$output = ($output -split "`r?`n") -join "`r`n"
if (-not $output.EndsWith("`r`n")) { $output += "`r`n" }

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($OutputPath, $output, $utf8NoBom)

$sizeKb = [Math]::Round((Get-Item -LiteralPath $OutputPath).Length / 1KB, 1)
$lineCount = (Get-Content -LiteralPath $OutputPath | Measure-Object -Line).Lines

Write-Host ''
Write-Host "Wrote $OutputPath" -ForegroundColor Green
Write-Host "  $lineCount lines, $sizeKb KB"
Write-Host "  build: $stamp"
