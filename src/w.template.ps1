<#
.SYNOPSIS
    Standalone system utility suite with a WPF desktop interface.

.NOTES
    Author: catsmoker
    Privileges: Administrator Required

    THIS FILE IS A BUILD TEMPLATE. Do not run it directly.
    Edit the files under src/ and run .\Compile.ps1 to regenerate ..\w.ps1
#>
param(
    # Dot-source this file for its functions, modules, and registry only --
    # no console, no logging, no WPF. Used by the background runspaces in
    # src/functions/04-runspace.ps1 so they get a complete environment from a
    # single source of truth instead of having everything marshalled by hand.
    [switch]$FmkLibraryMode
)

#>>BUILD<<#

# ==============================================================================
# CRASH HANDLER
# ==============================================================================
trap {
    # Library mode runs inside a background runspace. An interactive prompt there
    # would hang forever with no window to show it, so rethrow instead and let
    # the caller surface the real failure.
    if ($FmkLibraryMode) { throw }

    if (Get-Command Show-FmkConsole -ErrorAction SilentlyContinue) { Show-FmkConsole }
    try { [Console]::ResetColor() } catch { }
    Write-Host ""
    Write-Host "FreeMixKit crashed." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press ENTER to close..." -ForegroundColor Gray
    Read-Host | Out-Null
    Exit 1
}

# ==============================================================================
# 1. SETUP
# ==============================================================================

$ScriptUrl = "https://raw.githubusercontent.com/catsmoker/FreeMixKit/main/w.ps1"
# Single source of truth for the project repository. Issue links and the
# repository shortcut derive from this; never duplicate a guessed URL.
$FmkRepoUrl = "https://github.com/catsmoker/FreeMixKit"
$MSGameBarFixScriptUrl = "https://raw.githubusercontent.com/ajw0/ms-gamebar-fix/refs/heads/main/ms-gamebar-fix.ps1"
$WinUtilLatestScriptUrl = "https://christitus.com/win"
$WinUtilLegacyWindows10ScriptUrl = "https://github.com/ChrisTitusTech/winutil/releases/download/25.10.06/Winutil.ps1"
$AppVersion = "7"
$DataRoot = "C:\FreeMixKit"
$DataTemp = Join-Path $DataRoot "temp"
$DataLogs = Join-Path $DataRoot "logs"
$DataDownloads = Join-Path $DataRoot "downloads"
$SessionLogPath = Join-Path $DataLogs "FreeMixKit_Session.log"

# Cleared up front so every `if ($global:sync ...)` guard in the UI code has a
# defined value to test. Library mode leaves it null for good.
$global:sync = $null

foreach ($p in @($DataRoot, $DataTemp, $DataLogs, $DataDownloads)) {
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
}

# Deliberately NOT -Force. Library mode runs again for every module, and -Force
# would truncate the live session log each time.
if (-not (Test-Path -LiteralPath $SessionLogPath)) {
    New-Item -ItemType File -Path $SessionLogPath | Out-Null
}

$dataItem = Get-Item $DataRoot -Force
if (-not ($dataItem.Attributes -band [IO.FileAttributes]::Hidden)) {
    $dataItem.Attributes = $dataItem.Attributes -bor [IO.FileAttributes]::Hidden
}

# ==============================================================================
# 1b. RELAUNCH (ELEVATION + STA)
# ==============================================================================
# WPF needs an elevated process AND a single-threaded apartment. Windows
# PowerShell 5.1 is STA by default; PowerShell 7 is MTA and
# [Windows.Markup.XamlReader]::Parse throws on a non-STA thread. Both conditions
# are resolved in ONE relaunch rather than one restart each.

if (-not $FmkLibraryMode) {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $needsAdmin = -not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $needsSta = [Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA

    # The env var is the loop guard, and it is inherited by the child process.
    # That matters for the `irm | iex` path, where there is no command line left
    # to pass a switch on. Start-Process copies the current environment, so
    # setting it here is enough.
    if (($needsAdmin -or $needsSta) -and -not $env:FMK_RELAUNCHED) {

        $shellExe = if ($PSVersionTable.PSEdition -eq "Core") { "pwsh.exe" } else { "powershell.exe" }

        # Probe rather than assume that this host accepts -STA. If it does not,
        # the child would die before printing anything and the user would see
        # nothing at all. Windows PowerShell is always present and always STA,
        # so it is the fallback.
        if ($needsSta) {
            try {
                $probe = Start-Process $shellExe -ArgumentList '-NoProfile', '-STA', '-Command', 'exit 0' `
                    -PassThru -Wait -WindowStyle Hidden -ErrorAction Stop
                if ($probe.ExitCode -ne 0) { $shellExe = "powershell.exe" }
            }
            catch {
                $shellExe = "powershell.exe"
            }
        }

        $argList = @("-NoProfile", "-ExecutionPolicy", "Bypass")
        if ($needsSta) { $argList += "-STA" }
        if ($PSCommandPath) {
            $argList += @("-File", "`"$PSCommandPath`"")
        }
        else {
            # Running from `irm | iex`: there is no local file to point at.
            $argList += @("-Command", "irm $ScriptUrl | iex")
        }

        try {
            $env:FMK_RELAUNCHED = '1'
            if ($needsAdmin) {
                Start-Process $shellExe -ArgumentList $argList -Verb RunAs -ErrorAction Stop | Out-Null
            }
            else {
                Start-Process $shellExe -ArgumentList $argList -ErrorAction Stop | Out-Null
            }
            Exit
        }
        catch {
            Write-Host "Administrator rights are required to run FreeMixKit." -ForegroundColor Red
            Write-Host "Elevation was canceled or failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Press ENTER to close..." -ForegroundColor Gray
            Read-Host | Out-Null
            Exit 1
        }
    }

    # A relaunch already happened and we are still on the wrong apartment. Fail
    # loudly here rather than letting XamlReader::Parse throw something opaque.
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
        Write-Host "FreeMixKit needs a single-threaded apartment (STA) to show its window," -ForegroundColor Red
        Write-Host "and this host could not be relaunched into one." -ForegroundColor Red
        Write-Host "Try:  powershell -STA -File .\w.ps1" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Press ENTER to close..." -ForegroundColor Gray
        Read-Host | Out-Null
        Exit 1
    }
}

# ==============================================================================
# 1c. STRICT MODE
# ==============================================================================
# Version 1.0, not Latest. Latest (2.0+) throws on property access against a
# missing key or a $null, and the UI code is full of legitimate ones:
# FindName lookups, optional ModuleMeta fields, and keys absent from $sync
# before boot finishes. 1.0 keeps the uninitialized-variable checking that is
# actually useful and drops the rest.
Set-StrictMode -Version 1.0
$ErrorActionPreference = "Stop"

# ==============================================================================
# 2. STATIC SYSTEM INFO
# ==============================================================================
$SysInfo = @{
    OS  = (Get-CimInstance Win32_OperatingSystem).Caption
    CPU = (Get-CimInstance Win32_Processor).Name
    RAM = "{0:N1} GB" -f ((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB)
}

# ==============================================================================
# 3. HELPER FUNCTIONS
# ==============================================================================

#>>FUNCTIONS<<#

# ==============================================================================
# 3b. UI THEME TOKENS
# ==============================================================================

#>>THEME<<#

# ==============================================================================
# 3c. VENDORED ASSETS
# ==============================================================================
# Brand marks for the Software grid, as SVG path data. Vendored rather than
# fetched so the built script stays a single self-contained file that draws its
# icons with no network access and no image files.

#>>ASSETS<<#

# ==============================================================================
# 4. MODULE LIBRARY
# ==============================================================================

#>>MODULES<<#

# ==============================================================================
# 5. MODULE REGISTRATION
# ==============================================================================

#>>REGISTRY<<#

# ==============================================================================
# 5b. LIBRARY BOUNDARY
# ==============================================================================
# Everything above is the library: paths, helpers, theme, modules, registry.
# Everything below is UI only, and is skipped entirely when a background
# runspace dot-sources this file with -FmkLibraryMode.
#
# This is a guard rather than an early `return` on purpose. Dot-sourcing runs a
# script in the caller's scope, so a `return` here would also exit the scope
# that did the dot-sourcing -- aborting the runspace bootstrap. The guard is
# unambiguous.

if (-not $FmkLibraryMode) {

    # ==========================================================================
    # 6. UI DEFINITION
    # ==========================================================================

    #>>XAML<<#

    # ==========================================================================
    # 7. LAUNCH
    # ==========================================================================

    try {
        # Sibling file first (normal repo launch). The `irm | iex` path has no
        # directory to look in, so fall back to the vendored icon bytes
        # (src/assets/icon.ps1), materialized to temp: the window -- and hence
        # the taskbar button -- then always shows the project logo.
        $iconPath = ''
        if ($PSCommandPath) {
            $iconCandidate = Join-Path (Split-Path -Parent $PSCommandPath) 'freemixkit_icon.ico'
            if (Test-Path -LiteralPath $iconCandidate) { $iconPath = $iconCandidate }
        }
        if (-not $iconPath) {
            $iconVar = Get-Variable -Name 'FmkAppIconB64' -ValueOnly -ErrorAction SilentlyContinue
            if ($iconVar) {
                try {
                    $embeddedIcon = Join-Path $DataTemp 'freemixkit_icon.ico'
                    [System.IO.File]::WriteAllBytes($embeddedIcon, [System.Convert]::FromBase64String([string]$iconVar))
                    $iconPath = $embeddedIcon
                }
                catch { $iconPath = '' }
            }
        }

        Initialize-FmkUi -Xaml $FmkXaml -Version $AppVersion -IconPath $iconPath
    }
    catch {
        Show-FmkConsole
        try { [Console]::ResetColor() } catch { }
        Write-Host ""
        Write-Host "FreeMixKit crashed." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "Press ENTER to close..." -ForegroundColor Gray
        Read-Host | Out-Null
        Exit 1
    }

}
