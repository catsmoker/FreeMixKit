<#
.SYNOPSIS
    FreeMixKit v5.8
    Standalone system utility suite.

.NOTES
    Author: catsmoker
    Privileges: Administrator Required
#>

# ==============================================================================
# 1. SETUP & ADMIN CHECK
# ==============================================================================

$ScriptUrl = "https://raw.githubusercontent.com/catsmoker/FreeMixKit/main/w.ps1"
$MSGameBarFixScriptUrl = "https://raw.githubusercontent.com/ajw0/ms-gamebar-fix/refs/heads/main/ms-gamebar-fix.ps1"
$AppVersion = "5.8"
$DataRoot = "C:\\FreeMixKit"
$DataTemp = Join-Path $DataRoot "temp"
$DataLogs = Join-Path $DataRoot "logs"
$DataDownloads = Join-Path $DataRoot "downloads"
$SessionLogPath = Join-Path $DataLogs "FreeMixKit_Session.log"

foreach ($p in @($DataRoot, $DataTemp, $DataLogs, $DataDownloads)) {
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
}
$null = New-Item -ItemType File -Path $SessionLogPath -Force
$dataItem = Get-Item $DataRoot -Force
if (-not ($dataItem.Attributes -band [IO.FileAttributes]::Hidden)) {
    $dataItem.Attributes = $dataItem.Attributes -bor [IO.FileAttributes]::Hidden
}

$Host.UI.RawUI.WindowTitle = "FreeMixKit v$AppVersion"
try {
    # Force big window (120x40 is good for grid)
    $bufferSize = New-Object Management.Automation.Host.Size(120, 2000)
    $windowSize = New-Object Management.Automation.Host.Size(120, 40)
    if ($Host.UI.RawUI.BufferSize.Width -lt $windowSize.Width) { $Host.UI.RawUI.BufferSize = $bufferSize }
    $Host.UI.RawUI.WindowSize = $windowSize
    $Host.UI.RawUI.BufferSize = $bufferSize
}
catch { }

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $shellExe = if ($PSVersionTable.PSEdition -eq "Core") { "pwsh.exe" } else { "powershell.exe" }
    $argumentList = if ($PSCommandPath) {
        "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    }
    else {
        "-NoProfile -ExecutionPolicy Bypass -Command `"irm $ScriptUrl | iex`""
    }

    Start-Process $shellExe -ArgumentList $argumentList -Verb RunAs
    Exit
}

[Console]::BackgroundColor = "Black"
[Console]::ForegroundColor = "Green"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Clear-Host

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

function Write-Log($Message, $Type = "Info") {
    $c = switch ($Type) {
        "Info" { "White" }
        "Success" { "Cyan" }
        "Warn" { "Yellow" }
        "Error" { "Red" }
        default { "White" }
    }
    Write-Host " [$((Get-Date).ToString('HH:mm:ss'))] " -NoNewline -ForegroundColor DarkGray
    Write-Host $Message -ForegroundColor $c
    try {
        $tag = $Type.ToUpperInvariant()
        "[{0}] [{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $tag, $Message | Add-Content -Path $SessionLogPath -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch { }
}

function Confirm-Action([string]$Prompt) {
    $response = Read-Host "$Prompt (Y/N)"
    return $response -match '^[Yy]$'
}

function Ensure-WingetInstalled {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Log "Winget not found. Installing Winget..." "Warn"
        $bundlePath = Join-Path $DataTemp "winget.msixbundle"
        Invoke-WebRequest -Uri "https://aka.ms/getwinget" -OutFile $bundlePath
        Add-AppxPackage -Path $bundlePath
    }

    # Keep winget and its sources fresh
    try {
        Write-Log "Refreshing winget sources..." "Info"
        winget source update --disable-interactivity | Out-Null
    }
    catch {
        Write-Log "Winget source refresh failed: $($_.Exception.Message)" "Warn"
    }

    try {
        Write-Log "Checking winget client updates..." "Info"
        winget upgrade --id Microsoft.DesktopAppInstaller --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity | Out-Null
    }
    catch {
        Write-Log "Winget client upgrade skipped/failed: $($_.Exception.Message)" "Warn"
    }
}

function Test-WingetPackageInstalled([string]$Id) {
    Ensure-WingetInstalled
    $listOutput = winget list --id $Id --exact --source winget --accept-source-agreements 2>$null | Out-String
    return $listOutput -match [regex]::Escape($Id)
}

function Install-WingetPackage([string]$Id) {
    if (Test-WingetPackageInstalled -Id $Id) {
        Write-Log "Already installed: $Id"
        return
    }

    Write-Log "Installing package: $Id"
    Invoke-ExternalCommand -FilePath "winget.exe" -Arguments @(
        "install", "--id", $Id, "--exact", "-s", "winget",
        "--accept-package-agreements", "--accept-source-agreements", "--disable-interactivity"
    ) | Out-Null
}

function Test-NetworkConnectivity {
    try {
        return [bool](Test-Connection -ComputerName "1.1.1.1" -Count 1 -Quiet -ErrorAction Stop)
    }
    catch {
        return $false
    }
}

function Get-SpotifyExecutablePath {
    $candidates = @(
        "$env:APPDATA\\Spotify\\Spotify.exe",
        "$env:LOCALAPPDATA\\Microsoft\\WindowsApps\\Spotify.exe"
    )

    $msixRoot = Join-Path $env:LOCALAPPDATA "Microsoft\\WindowsApps"
    if (Test-Path $msixRoot) {
        $msixSpotify = Get-ChildItem -Path $msixRoot -Filter "Spotify*.exe" -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "spotify" } |
            Select-Object -First 1
        if ($msixSpotify) { $candidates += $msixSpotify.FullName }
    }

    foreach ($p in $candidates | Select-Object -Unique) {
        if (Test-Path $p) { return $p }
    }
    return $null
}

function Invoke-ExternalCommand {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter()][string[]]$Arguments = @()
    )

    $proc = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -NoNewWindow
    Wait-Process -Id $proc.Id -ErrorAction SilentlyContinue

    if ($proc.ExitCode -ne 0) {
        throw "Command failed (exit $($proc.ExitCode)): $FilePath $($Arguments -join ' ')"
    }

    return [pscustomobject]@{
        FilePath = $FilePath
        Arguments = ($Arguments -join " ")
        ExitCode = $proc.ExitCode
    }
}

function Get-ModuleMetaValue([string]$ActionKey, [string]$Name, $DefaultValue) {
    if ($ModuleMeta.Contains($ActionKey) -and $ModuleMeta[$ActionKey].Contains($Name) -and $null -ne $ModuleMeta[$ActionKey][$Name]) {
        return $ModuleMeta[$ActionKey][$Name]
    }
    return $DefaultValue
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
                Status = "Canceled"
                Message = "Canceled by user"
                Duration = [timespan]::Zero
                RollbackHint = $rollbackHint
                PreviousState = ""
            }
        }
    }

    if ($requiresNetwork -and -not (Test-NetworkConnectivity)) {
        throw "Network is required, but connectivity check failed."
    }

    $Script:ModuleExecutionContext = @{}
    $start = Get-Date
    $hadLastExitCode = Test-Path variable:global:LASTEXITCODE
    $oldLastExitCode = if ($hadLastExitCode) { $global:LASTEXITCODE } else { 0 }
    $global:LASTEXITCODE = 0

    try {
        & $Modules[$ActionKey] | Out-Null

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
            Status = "Success"
            Message = "Completed"
            Duration = $duration
            RollbackHint = $rollbackHint
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

# ==============================================================================
# 4. MODULE LIBRARY
# ==============================================================================

$Modules = [ordered]@{}
$ModuleMeta = [ordered]@{}
$ModuleResults = New-Object System.Collections.Generic.List[object]

function Register-Module(
    [string]$Key,
    [string]$Label,
    [string]$Description,
    [scriptblock]$Action,
    [string]$Risk = "Normal",
    [hashtable]$Options = @{}
) {
    $Modules[$Key] = $Action
    $ModuleMeta[$Key] = [ordered]@{
        Key         = $Key
        Label       = $Label
        Description = $Description
        Risk        = $Risk
        RequiresNetwork = $false
        ConfirmMessage  = $null
        Verify          = $null
        RollbackHint    = ""
    }

    foreach ($k in $Options.Keys) {
        $ModuleMeta[$Key][$k] = $Options[$k]
    }
}

function Add-ModuleResult([string]$Module, [string]$Status, [string]$Message, [timespan]$Duration, [string]$RollbackHint = "", [string]$PreviousState = "") {
    $ModuleResults.Add([pscustomobject]@{
            Timestamp   = Get-Date
            Module      = $Module
            Status      = $Status
            DurationSec = [Math]::Round($Duration.TotalSeconds, 2)
            Message     = $Message
            RollbackHint = $RollbackHint
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

# --- DEVELOPER ---
$Modules["DevChoice"] = {
    Write-Log "Starting Developer Environment Setup..." "Warn"
    Write-Log "Installing: VS Redists, .NET, Node, Python, Java, Tools, Bibata Cursor."
    
    # 1. Winget
    Ensure-WingetInstalled

    # 2. Packages
    # VC++ runtime coverage (x86 + x64):
    # Latest v14 (VS 2017–2026), VS 2015 (VC++ 14.0), 2013 (12.0), 2012 (11.0), 2010 (10.0), 2008 (9.0), 2005 (8.0)
    $vcRedists = @(
        @{ Version = "Latest v14 (VS 2017–2026)"; Arch = "x86"; Id = "Microsoft.VCRedist.2015+.x86" },
        @{ Version = "Latest v14 (VS 2017–2026)"; Arch = "x64"; Id = "Microsoft.VCRedist.2015+.x64" },
        @{ Version = "Visual Studio 2015 (VC++ 14.0)"; Arch = "x86"; Id = "Microsoft.VCRedist.2015+.x86" },
        @{ Version = "Visual Studio 2015 (VC++ 14.0)"; Arch = "x64"; Id = "Microsoft.VCRedist.2015+.x64" },
        @{ Version = "Visual Studio 2013 (VC++ 12.0)"; Arch = "x86"; Id = "Microsoft.VCRedist.2013.x86" },
        @{ Version = "Visual Studio 2013 (VC++ 12.0)"; Arch = "x64"; Id = "Microsoft.VCRedist.2013.x64" },
        @{ Version = "Visual Studio 2012 (VC++ 11.0)"; Arch = "x86"; Id = "Microsoft.VCRedist.2012.x86" },
        @{ Version = "Visual Studio 2012 (VC++ 11.0)"; Arch = "x64"; Id = "Microsoft.VCRedist.2012.x64" },
        @{ Version = "Visual Studio 2010 (VC++ 10.0)"; Arch = "x86"; Id = "Microsoft.VCRedist.2010.x86" },
        @{ Version = "Visual Studio 2010 (VC++ 10.0)"; Arch = "x64"; Id = "Microsoft.VCRedist.2010.x64" },
        @{ Version = "Visual Studio 2008 (VC++ 9.0)"; Arch = "x86"; Id = "Microsoft.VCRedist.2008.x86" },
        @{ Version = "Visual Studio 2008 (VC++ 9.0)"; Arch = "x64"; Id = "Microsoft.VCRedist.2008.x64" },
        @{ Version = "Visual Studio 2005 (VC++ 8.0)"; Arch = "x86"; Id = "Microsoft.VCRedist.2005.x86" },
        @{ Version = "Visual Studio 2005 (VC++ 8.0)"; Arch = "x64"; Id = "Microsoft.VCRedist.2005.x64" }
    )

    Write-Log "Queued VC++ Redistributables (x86/x64):" "Info"
    foreach ($r in $vcRedists) {
        if ($r.Id) { Write-Log " - $($r.Version) [$($r.Arch)] -> $($r.Id)" "Info" }
    }

    $packages = @(
        # prefer latest stable IDs that resolve on winget today
        "Microsoft.DotNet.SDK.10", "Microsoft.DotNet.Runtime.9", "Microsoft.DotNet.Runtime.8", "OpenJS.NodeJS.LTS", "Python.Python.3.14", "EclipseAdoptium.Temurin.25.JDK",
        "Microsoft.PowerShell", "Git.Git", "Gyan.FFmpeg", "M2Team.NanaZip", "Notepad++.Notepad++", "AdrienAllard.FileConverter"
    ) + ($vcRedists | Where-Object { $_ -and $_.Id } | ForEach-Object { $_.Id } | Sort-Object -Unique)

    foreach ($id in $packages) {
        try {
            Install-WingetPackage -Id $id
        }
        catch {
            Write-Log "Failed package '$id': $($_.Exception.Message)" "Error"
        }
    }

    # 2b. Game-related runtimes via winget
    $gamePackages = @("Microsoft.DirectX", "Microsoft.XNARedist")
    foreach ($gid in $gamePackages) {
        try {
            Write-Log "Installing game runtime: $gid" "Info"
            Install-WingetPackage -Id $gid
        }
        catch {
            Write-Log "Failed game runtime '$gid': $($_.Exception.Message)" "Error"
        }
    }

    Write-Log "Enabling .NET Framework 3.5 (Windows Feature)..." "Info"
    try {
        Start-Process "dism.exe" -ArgumentList "/Online /Enable-Feature /FeatureName:NetFx3 /All /NoRestart" -Wait -NoNewWindow
    }
    catch { Write-Log ".NET 3.5 enable failed: $($_.Exception.Message)" "Warn" }

    # 3. Notepad Fix
    Write-Log "Replacing Windows Notepad with Notepad++..." "Info"
    $nppPath = "C:\\Program Files\\Notepad++\\notepad++.exe"
    if (-not (Test-Path $nppPath)) {
        Write-Log "Notepad++ not found yet; skipping Notepad replacement." "Warn"
    }
    else {
        Get-AppxPackage *Microsoft.WindowsNotepad* | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        $reg = @"
Windows Registry Editor Version 5.00
[HKEY_CLASSES_ROOT\.txt]
@="txtfile"
"PerceivedType"="text"
"Content Type"="text/plain"
[HKEY_CLASSES_ROOT\.txt\ShellNew]
"NullFile"=""
[HKEY_CLASSES_ROOT\txtfile]
@="Text Document"
[HKEY_CLASSES_ROOT\txtfile\DefaultIcon]
@="$nppPath,0"
[HKEY_CLASSES_ROOT\txtfile\shell\open\command]
@="\""$nppPath\"" \"%1\""
[HKEY_CLASSES_ROOT\*\shell\NotepadPlusPlus]
@="Open with Notepad++"
"Icon"="\""$nppPath\"""
[HKEY_CLASSES_ROOT\*\shell\NotepadPlusPlus\command]
@="\""$nppPath\"" \"%1\""
"@
        $regPath = Join-Path $DataTemp "nppfix.reg"
        $reg | Out-File $regPath -Encoding ASCII -Force
        Start-Process reg.exe -Argument "import `"$regPath`"" -Wait -NoNewWindow
    }
    
    # 4. Bibata
    Write-Log "Installing Bibata Cursor..."
    try {
        $zip = "$DataDownloads\\Bibata.zip"; $dest = "$DataDownloads\\Bibata"
        Invoke-WebRequest "https://github.com/ful1e5/Bibata_Cursor/releases/download/v2.0.7/Bibata-Modern-Classic-Windows.zip" -OutFile $zip
        Expand-Archive $zip -Dest $dest -Force
        $regularInf = Get-ChildItem "$dest" -Recurse -Filter "*.inf" | Where-Object { $_.FullName -match "Regular" } | Select-Object -First 1
        if (-not $regularInf) { $regularInf = Get-ChildItem "$dest" -Recurse -Filter "*.inf" | Select-Object -First 1 }
        if ($regularInf) { Start-Process "RUNDLL32.EXE" -Arg "SETUPAPI.DLL,InstallHinfSection DefaultInstall 128 $($regularInf.FullName)" -Wait }
    }
    catch { Write-Log "Cursor Failed" "Error" }

    Write-Log "Done!" "Success"
}

# --- MAINTENANCE ---
$Modules["CleanSystem"] = {
    Write-Log "Cleaning..."
    Remove-Item "$env:TEMP\*" -Recurse -Force -EA SilentlyContinue
    Remove-Item "$DataTemp\*" -Recurse -Force -EA SilentlyContinue
    Remove-Item "C:\Windows\Temp\*" -Recurse -Force -EA SilentlyContinue
    Remove-Item "C:\Windows\Prefetch\*" -Recurse -Force -EA SilentlyContinue
    Clear-DnsClientCache
    Write-Log "Cleaned." "Success"
}
$Modules["SystemRepair"] = {
    Invoke-ExternalCommand -FilePath "sfc.exe" -Arguments @("/scannow") | Out-Null
    Invoke-ExternalCommand -FilePath "DISM.exe" -Arguments @("/Online", "/Cleanup-Image", "/RestoreHealth") | Out-Null
    Write-Log "Done." "Success"
}
$Modules["MalwareScan"] = { 
    $mrt = "$env:SystemRoot\System32\MRT.exe"
    if (!(Test-Path $mrt)) {
        Invoke-WebRequest "https://go.microsoft.com/fwlink/?LinkID=212732" -OutFile "$DataDownloads\\MRT.exe"
        $mrt = "$DataDownloads\\MRT.exe"
    }
    Start-Process $mrt -Wait

    $kvrt = Join-Path $DataDownloads "kvrt.exe"
    Invoke-WebRequest "https://devbuilds.s.kaspersky-labs.com/kvrt/latest/full/kvrt.exe" -OutFile $kvrt
    Start-Process $kvrt -Wait
}
$Modules["SystemReport"] = { 
    $f = "$env:USERPROFILE\Desktop\SysReport.txt"
    "OS: $($SysInfo.OS)`nCPU: $($SysInfo.CPU)`nRAM: $($SysInfo.RAM)" | Out-File $f
    Invoke-Item $f 
}

# --- APPS ---
$Modules["AdobeGenP"] = {
    Write-Log "Opening Creative Cloud..."
    Start-Process "https://www.adobe.com/download/creative-cloud"
    Write-Log "Opening GenP..."
    Start-Process "https://wiki.dbzer0.com/genp-guides/guide/#guide-2"
}
$Modules["WingetUpgrade"] = {
    Ensure-WingetInstalled
    Invoke-ExternalCommand -FilePath "winget.exe" -Arguments @(
        "upgrade", "--all", "--include-unknown",
        "--accept-source-agreements", "--accept-package-agreements"
    ) | Out-Null
}
$Modules["Spicetify"] = {
    # ===============================
    # Spotify → Spicetify Installer
    # ===============================

    Write-Log "Checking Spotify installation..." "Info"

    $spotifyExe = Get-SpotifyExecutablePath
    $currentUser = $env:USERNAME
    $asUser = {
        param([string]$Title, [string]$ScriptText)
        $runScript = Join-Path $DataTemp "RunAsUser_$Title.ps1"
        Set-Content -Path $runScript -Value $ScriptText -Force

        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$runScript`""
        $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(2)
        $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited

        Unregister-ScheduledTask -TaskName $Title -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName $Title -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
        Start-ScheduledTask -TaskName $Title
        Start-Sleep -Seconds 10
        Unregister-ScheduledTask -TaskName $Title -Confirm:$false -ErrorAction SilentlyContinue
    }

    if (-not (Test-Path $spotifyExe)) {

        Ensure-WingetInstalled

        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Write-Log "Installing Spotify using winget (standard user)..." "Info"
            $asUser.Invoke("WingetSpotify", @"
winget install --id Spotify.Spotify --exact -s winget --scope user --accept-source-agreements --accept-package-agreements --disable-interactivity
"@)
        }
        else {
            Write-Log "Winget not found. Using direct installer (standard user)..." "Info"

            $spotifyInstaller = "$DataDownloads\\SpotifySetup.exe"
            Invoke-WebRequest "https://download.scdn.co/SpotifySetup.exe" -OutFile $spotifyInstaller

            $asUser.Invoke("DirectSpotify", @"
Start-Process '$spotifyInstaller' -ArgumentList '/silent' -Wait
"@)
        }

        # Wait until Spotify exists
        for ($i = 0; $i -lt 20; $i++) {
            $spotifyExe = Get-SpotifyExecutablePath
            if ($spotifyExe) { break }
            Start-Sleep 2
        }

        if (-not $spotifyExe) {
            throw "Spotify installation failed or timed out."
        }
    }
    else {
        Write-Log "Spotify already installed." "Info"
    }

    Write-Log "Spotify detected at: $spotifyExe" "Success"

    # Run Spotify once so it initializes user data before Spicetify patches it
    Write-Log "Launching Spotify once to initialize (30s countdown)..." "Info"
    try {
        $proc = Start-Process -FilePath $spotifyExe -PassThru -WindowStyle Hidden
        for ($sec = 30; $sec -ge 1; $sec--) {
            $percent = [int](100 * (30 - $sec) / 30)
            Write-Progress -Activity "Spotify first-launch warmup" -Status "Waiting $sec s..." -PercentComplete $percent
            if ($proc.HasExited) { break }
            Start-Sleep -Seconds 1
        }
        Write-Progress -Activity "Spotify first-launch warmup" -Completed
        if ($proc -and -not $proc.HasExited) {
            Stop-Process -Id $proc.Id -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Log "Could not auto-launch/close Spotify: $($_.Exception.Message)" "Warn"
    }

    # ===============================
    # Prepare Spicetify Installer
    # ===============================

    Write-Log "Preparing Spicetify Installation..." "Info"

    $tempDir = $DataTemp
    $installScript = Join-Path $tempDir "Install-Spicetify.ps1"

    $scriptContent = @'
Write-Host "=== Spicetify Installer ===" -ForegroundColor Cyan
try {
    irm https://raw.githubusercontent.com/spicetify/marketplace/main/resources/install.ps1 | iex
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host "Press ENTER to close..."
Read-Host
'@

    Set-Content -Path $installScript -Value $scriptContent -Force

    # ===============================
    # Run as Standard User
    # ===============================

    $taskName = "FreeMixKit_Spicetify_User"
    $currentUser = $env:USERNAME

    Write-Log "Launching Spicetify installer as user: $currentUser..." "Info"

    $action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$installScript`""

    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(2)
    $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited

    try {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName

        Write-Log "Spicetify installer launched successfully." "Success"
    }
    catch {
        Write-Log "Failed to launch Spicetify installer." "Error"
    }

}

$Modules["Legcord"] = {
    try { 
        $u = ((Invoke-RestMethod "https://api.github.com/repos/Legcord/Legcord/releases/latest").assets | Where-Object name -match ".exe" | Select-Object -First 1).browser_download_url
        Invoke-WebRequest $u -OutFile "$DataDownloads\\legcord.exe"; Start-Process "$DataDownloads\\legcord.exe" -Wait
    }
    catch {
        Write-Log "Legcord install failed: $($_.Exception.Message)" "Error"
        throw
    }
}

# --- ACTIVATION ---
$Modules["MAS"] = { Invoke-RestMethod https://get.activated.win | Invoke-Expression }
$Modules["IAS"] = { Invoke-RestMethod https://coporton.com/ias | Invoke-Expression }

# --- UTILS ---
$Modules["WinUtil"] = { Invoke-RestMethod https://christitus.com/win | Invoke-Expression }
$Modules["FixResolution"] = { Invoke-WebRequest "https://www.monitortests.com/download/cru/cru-1.5.3.zip" -OutFile "$DataDownloads\\cru.zip"; Expand-Archive "$DataDownloads\\cru.zip" "$DataTemp\\CRU" -Force; Start-Process "$DataTemp\\CRU\\CRU.exe" -Wait; Start-Process "$DataTemp\\CRU\\restart64.exe" -Wait }
$Modules["MSGameBarFix"] = {
    $shellExe = if (Get-Command pwsh.exe -ErrorAction SilentlyContinue) { "pwsh.exe" } else { "powershell.exe" }
    $localScriptPath = Join-Path $DataTemp "ms-gamebar-fix.ps1"

    Write-Log "Downloading and launching ms-gamebar-fix..." "Info"
    Invoke-WebRequest -Uri $MSGameBarFixScriptUrl -OutFile $localScriptPath
    Start-Process -FilePath $shellExe -ArgumentList @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", ('"{0}"' -f $localScriptPath)
    ) -Wait
}
$Modules["OpenGitHub"] = { Start-Process "https://github.com/catsmoker/FreeMixKit" }
$Modules["OpenWebsite"] = { Start-Process "https://catsmoker.vercel.app/" }

$Modules["AddShortcut"] = {
    $iconUrl = "https://raw.githubusercontent.com/catsmoker/FreeMixKit/refs/heads/main/freemixkit_icon.ico"
    $iconPath = Join-Path $DataDownloads "freemixkit_icon.ico"
    try {
        Invoke-WebRequest $iconUrl -OutFile $iconPath -ErrorAction SilentlyContinue
    }
    catch {
        Write-Log "Icon download failed, continuing without icon: $($_.Exception.Message)" "Warn"
    }

    $shortcutShell = if (Get-Command pwsh.exe -ErrorAction SilentlyContinue) { "pwsh.exe" } else { "powershell.exe" }
    $s = (New-Object -ComObject WScript.Shell).CreateShortcut("$env:USERPROFILE\Desktop\FreeMixKit.lnk")
    $s.TargetPath = $shortcutShell
    $s.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"irm https://raw.githubusercontent.com/catsmoker/FreeMixKit/main/w.ps1 | iex`""
    if (Test-Path $iconPath) { $s.IconLocation = $iconPath }
    $s.Save()
    
    # 3. Set RunAsAdministrator (Byte Patching)
    try {
        $bytes = [System.IO.File]::ReadAllBytes("$env:USERPROFILE\Desktop\FreeMixKit.lnk")
        $bytes[0x15] = $bytes[0x15] -bor 0x20 # Bit 5 = RunAsAdmin
        [System.IO.File]::WriteAllBytes("$env:USERPROFILE\Desktop\FreeMixKit.lnk", $bytes)
    }
    catch {
        Write-Log "Failed to set shortcut RunAsAdmin flag: $($_.Exception.Message)" "Error"
        throw
    }
}

# Register metadata for modules (single source of truth for labels/descriptions/risk)
Register-Module "DevChoice" "DEV CHOICE (Full)" "Installs: VS Redists, .NET, Node.js, Python, Java, PowerShell, Git, FFmpeg, nanazip, Notepad++, File Converter, Bibata Cursor." $Modules["DevChoice"] "Medium" @{
    RequiresNetwork = $true
}
Register-Module "CleanSystem" "Clean System Junk" "Removes temp files, prefetch, and clears DNS cache." $Modules["CleanSystem"] "Low"
Register-Module "SystemRepair" "System Repair" "Runs SFC Scannow and DISM RestoreHealth." $Modules["SystemRepair"] "Low" @{
}
Register-Module "MalwareScan" "Malware Scan" "Runs MRT, then downloads and runs Kaspersky Virus Removal Tool (KVRT)." $Modules["MalwareScan"] "Low" @{
    RequiresNetwork = $true
}
Register-Module "SystemReport" "System Report" "Generates a text file with system specs on your desktop." $Modules["SystemReport"] "Low"
Register-Module "MAS" "MAS" "Runs MAS (Microsoft Activation Scripts) to activate Windows." $Modules["MAS"] "High"
Register-Module "IAS" "IAS" "Activates Internet Download Manager (IDM)." $Modules["IAS"] "High"
Register-Module "AdobeGenP" "Adobe GenP" "Downloads Creative Cloud and GenP activator." $Modules["AdobeGenP"] "High" @{
    RequiresNetwork = $true
}
Register-Module "WingetUpgrade" "Winget Upgrade" "Upgrades all installed software via Winget." $Modules["WingetUpgrade"] "Low" @{
    RequiresNetwork = $true
    Verify = { [bool](Get-Command winget -ErrorAction SilentlyContinue) }
}
Register-Module "Spicetify" "Spicetify" "Installs Spicetify for Spotify customization/ad-blocking." $Modules["Spicetify"] "Medium" @{
    RequiresNetwork = $true
}
Register-Module "Legcord" "Legcord" "Installs Legcord (BetterDiscord alternative)." $Modules["Legcord"] "Low" @{
    RequiresNetwork = $true
}
Register-Module "WinUtil" "WinUtil" "Launches Chris Titus Tech's Windows Utility." $Modules["WinUtil"] "Medium" @{
    RequiresNetwork = $true
}
Register-Module "FixResolution" "Fix Resolution" "Uses CRU to restart graphics driver and fix resolution." $Modules["FixResolution"] "Medium" @{
    RequiresNetwork = $true
}
Register-Module "MSGameBarFix" "ms-gamebar-fix" "Downloads and runs the Game Bar popup fixer from ajw0/ms-gamebar-fix." $Modules["MSGameBarFix"] "Medium" @{
    RequiresNetwork = $true
}
Register-Module "OpenGitHub" "GitHub Repository" "Opens the FreeMixKit GitHub page in your browser." $Modules["OpenGitHub"] "Low"
Register-Module "OpenWebsite" "Website" "Opens the author's website in your browser." $Modules["OpenWebsite"] "Low"
Register-Module "AddShortcut" "Add Shortcut" "Creates a shortcut for this script on the Desktop." $Modules["AddShortcut"] "Low"

# ==============================================================================
# 5. GRID MENU CONFIGURATION
# ==============================================================================

# Define Columns. Type: H=Header, I=Item
$Col1 = @(
    @{T = "H"; L = "[ DEVELOPER ]" }
    @{T = "I"; L = "DEV CHOICE"; A = "DevChoice"; D = "Installs: VS Redists, .NET, Node.js, Python, Java, PowerShell, Git, FFmpeg, nanazip, Notepad++, File Converter, Bibata Cursor." }
    @{T = "H"; L = "" }
    @{T = "H"; L = "[ MAINTENANCE ]" }
    @{T = "I"; L = "Clean System Junk"; A = "CleanSystem"; D = "Removes temp files, prefetch, and clears DNS cache." }
    @{T = "I"; L = "System Repair"; A = "SystemRepair"; D = "Runs SFC Scannow and DISM RestoreHealth." }
    @{T = "I"; L = "Malware Scan"; A = "MalwareScan"; D = "Runs MRT, then downloads and runs Kaspersky Virus Removal Tool (KVRT)." }
    @{T = "I"; L = "System Report"; A = "SystemReport"; D = "Generates a text file with system specs on your desktop." }
    @{T = "H"; L = "" }
    @{T = "H"; L = "[ ACTIVATION ]" }
    @{T = "I"; L = "MAS"; A = "MAS"; D = "Runs MAS (Microsoft Activation Scripts) to activate Windows." }
    @{T = "I"; L = "IAS"; A = "IAS"; D = "Activates Internet Download Manager (IDM)." }
)

$Col2 = @(
    @{T = "H"; L = "[ SOFTWARE ]" }
    @{T = "I"; L = "Adobe GenP"; A = "AdobeGenP"; D = "Downloads Creative Cloud and GenP activator." }
    @{T = "I"; L = "Winget Upgrade"; A = "WingetUpgrade"; D = "Upgrades all installed software via Winget." }
    @{T = "I"; L = "Spicetify"; A = "Spicetify"; D = "Installs Spicetify for Spotify customization/ad-blocking." }
    @{T = "I"; L = "Legcord"; A = "Legcord"; D = "Installs Legcord (BetterDiscord alternative)." }
    @{T = "H"; L = "" }
    @{T = "H"; L = "[ UTILITIES ]" }
    @{T = "I"; L = "WinUtil"; A = "WinUtil"; D = "Launches Chris Titus Tech's Windows Utility." }
    @{T = "I"; L = "Fix Resolution"; A = "FixResolution"; D = "Uses CRU to restart graphics driver and fix resolution." }
    @{T = "I"; L = "ms-gamebar-fix"; A = "MSGameBarFix"; D = "Downloads and runs the Game Bar popup fixer from ajw0/ms-gamebar-fix." }
    @{T = "H"; L = "" }
    @{T = "H"; L = "[ LINKS ]" }
    @{T = "I"; L = "GitHub Repository"; A = "OpenGitHub"; D = "Open the FreeMixKit GitHub page." }
    @{T = "I"; L = "Website"; A = "OpenWebsite"; D = "Open catsmoker's website." }
    @{T = "H"; L = "" }
    @{T = "H"; L = "[ EXIT ]" }
    @{T = "I"; L = "Add Shortcut"; A = "AddShortcut"; D = "Creates a shortcut for this script on the Desktop." }
    @{T = "I"; L = "Exit Application"; A = "EXIT"; D = "Closes the application." }
)

# Sync menu labels/descriptions from module metadata.
foreach ($column in @($Col1, $Col2)) {
    foreach ($item in $column) {
        if ($item.T -eq "I" -and $item.A -ne "EXIT" -and $ModuleMeta.Contains($item.A)) {
            $item.L = $ModuleMeta[$item.A].Label
            $item.D = $ModuleMeta[$item.A].Description
        }
    }
}

# Build Navigation Grid
# NavGrid is an array where item = {Col=0/1, Row=IndexInCol, Label, Action}
$NavItems = @()

# Process Col 1
for ($i = 0; $i -lt $Col1.Count; $i++) {
    if ($Col1[$i].T -eq "I") { $NavItems += @{C = 0; R = $i; L = $Col1[$i].L; A = $Col1[$i].A; D = $Col1[$i].D } }
}
# Process Col 2
for ($i = 0; $i -lt $Col2.Count; $i++) {
    if ($Col2[$i].T -eq "I") { $NavItems += @{C = 1; R = $i; L = $Col2[$i].L; A = $Col2[$i].A; D = $Col2[$i].D } }
}

for ($i = 0; $i -lt $NavItems.Count; $i++) {
    $NavItems[$i]["N"] = $i + 1
}

$SelIdx = 0 # Index in $NavItems
$NumberInput = ""
$LastActionLabel = "None"
$LastActionStatus = "N/A"
$LastActionMessage = "No action run yet."

function Move-Selection([string]$Direction, $NavItems, [ref]$SelIdx) {
    $curr = $NavItems[$SelIdx.Value]
    $target = $null
    switch ($Direction) {
        "Up"   { $target = $NavItems | Where-Object { $_.C -eq $curr.C -and $_.R -lt $curr.R } | Select-Object -Last 1 }
        "Down" { $target = $NavItems | Where-Object { $_.C -eq $curr.C -and $_.R -gt $curr.R } | Select-Object -First 1 }
        "Right" {
            if ($curr.C -eq 0) {
                $target = $NavItems | Where-Object { $_.C -eq 1 } | Sort-Object { [Math]::Abs($_.R - $curr.R) } | Select-Object -First 1
            }
        }
        "Left" {
            if ($curr.C -eq 1) {
                $target = $NavItems | Where-Object { $_.C -eq 0 } | Sort-Object { [Math]::Abs($_.R - $curr.R) } | Select-Object -First 1
            }
        }
    }
    if ($target) {
        $idx = [Array]::IndexOf($NavItems, $target)
        if ($idx -ge 0) { $SelIdx.Value = $idx }
    }
}

# ==============================================================================
# 6. RENDER LOOP
# ==============================================================================

Clear-Host
while ($true) {
    $uiWidth = [Math]::Max(100, $Host.UI.RawUI.WindowSize.Width)
    $line = "=" * ($uiWidth - 1)
    $dash = "-" * ($uiWidth - 1)
    $timeNow = Get-Date -Format "ddd HH:mm:ss"

    [Console]::SetCursorPosition(0, 0)
    Write-Host $line -F Blue
    Write-Host " FREEMIXKIT v$AppVersion" -NoNewline -F Cyan
    Write-Host " | $timeNow | ARROWS Navigate | NUMBER + ENTER Run | Q/ESC Exit" -F Gray
    Write-Host $line -F Blue
    Write-Host " OS: $($SysInfo.OS) | CPU: $($SysInfo.CPU) | RAM: $($SysInfo.RAM)" -F DarkGray
    Write-Host $dash -F Blue
    
    $startY = 5
    
    # RENDER COL 1
    $y = $startY
    $x = 2
    for ($i = 0; $i -lt $Col1.Count; $i++) {
        [Console]::SetCursorPosition($x, $y)
        $item = $Col1[$i]
        
        if ($item.T -eq "H") { 
            Write-Host $item.L -F DarkGray
        }
        else {
            # Check if selected
            $isSel = ($NavItems[$SelIdx].C -eq 0 -and $NavItems[$SelIdx].R -eq $i)
            $navItem = $NavItems | Where-Object { $_.C -eq 0 -and $_.R -eq $i } | Select-Object -First 1
            $label = "[{0}] {1}" -f $navItem.N, $item.L
            if ($label.Length -gt 48) { $label = $label.Substring(0, 45) + "..." }
            if ($isSel) { Write-Host " > $label " -B DarkCyan -F White }
            else { Write-Host "   $label " -F Green }
        }
        $y++
    }

    # RENDER COL 2
    $y = $startY
    $x = 60
    for ($i = 0; $i -lt $Col2.Count; $i++) {
        [Console]::SetCursorPosition($x, $y)
        $item = $Col2[$i]
        
        if ($item.T -eq "H") { 
            Write-Host $item.L -F DarkGray
        }
        else {
            # Check if selected
            $isSel = ($NavItems[$SelIdx].C -eq 1 -and $NavItems[$SelIdx].R -eq $i)
            $navItem = $NavItems | Where-Object { $_.C -eq 1 -and $_.R -eq $i } | Select-Object -First 1
            $label = "[{0}] {1}" -f $navItem.N, $item.L
            if ($label.Length -gt 48) { $label = $label.Substring(0, 45) + "..." }
            if ($isSel) { Write-Host " > $label " -B DarkCyan -F White }
            else { Write-Host "   $label " -F Green }
        }
        $y++
    }
    
    # Helper Stats area below
    $maxY = $startY + [Math]::Max($Col1.Count, $Col2.Count) + 1
    [Console]::SetCursorPosition(0, $maxY)
    Write-Host $line -F Blue

    $curr = $NavItems[$SelIdx]
    $currMeta = if ($curr.A -ne "EXIT" -and $ModuleMeta.Contains($curr.A)) { $ModuleMeta[$curr.A] } else { $null }
    $risk = if ($currMeta) { $currMeta.Risk } else { "N/A" }
    $riskColor = switch ($risk) {
        "High" { "Red" }
        "Medium" { "Yellow" }
        "Low" { "Green" }
        default { "DarkGray" }
    }

    for ($lineIdx = 1; $lineIdx -le 5; $lineIdx++) {
        [Console]::SetCursorPosition(0, $maxY + $lineIdx)
        Write-Host (" " * ($uiWidth - 1))
    }

    [Console]::SetCursorPosition(2, $maxY + 1)
    $infoText = $curr.D
    if ($infoText.Length -gt ($uiWidth - 10)) { $infoText = $infoText.Substring(0, $uiWidth - 13) + "..." }
    Write-Host "$($curr.L): " -NoNewline -F Cyan
    Write-Host $infoText -F Gray
    [Console]::SetCursorPosition(2, $maxY + 2)
    Write-Host "Risk: " -NoNewline -F Gray
    Write-Host $risk -F $riskColor
    [Console]::SetCursorPosition(0, $maxY + 4)
    Write-Host $dash -F Blue
    [Console]::SetCursorPosition(2, $maxY + 5)
    Write-Host "Input: $NumberInput" -F Gray
    [Console]::SetCursorPosition(2, $maxY + 6)
    Write-Host "Keys: Up/Down/W/S move | Left/Right/A/D switch column | Type number + Enter | Backspace clear | Q or Esc exit" -F DarkGray

    # INPUT HANDLING
    $k = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

    switch ($k.VirtualKeyCode) {
        {$_ -in 38,87} { Move-Selection "Up" $NavItems ([ref]$SelIdx) }
        {$_ -in 40,83} { Move-Selection "Down" $NavItems ([ref]$SelIdx) }
        {$_ -in 39,68} { Move-Selection "Right" $NavItems ([ref]$SelIdx) }
        {$_ -in 37,65} { Move-Selection "Left" $NavItems ([ref]$SelIdx) }
        13 {
            # ENTER
            if ($NumberInput) {
                $targetNumber = 0
                if ([int]::TryParse($NumberInput, [ref]$targetNumber)) {
                    $target = $NavItems | Where-Object { $_.N -eq $targetNumber } | Select-Object -First 1
                    if ($target) {
                        $SelIdx = $NavItems.IndexOf($target)
                        $curr = $NavItems[$SelIdx]
                    }
                    else {
                        $LastActionLabel = "Number Input"
                        $LastActionStatus = "Invalid"
                        $LastActionMessage = "No item matches number $NumberInput."
                        $NumberInput = ""
                        continue
                    }
                }
                $NumberInput = ""
            }

            $action = $curr.A
            if ($action -eq "EXIT") {
                Export-SessionResults
                Clear-Host
                exit
            }
            
            [Console]::SetCursorPosition(2, $maxY + 2)
            Write-Host "Executing: $($curr.L)..." -F Cyan
            if ($Modules.Contains($action)) {
                try {
                    $result = Invoke-ModuleAction -ActionKey $action
                    $status = $result.Status
                    $message = $result.Message
                    $duration = $result.Duration
                    $rollbackHint = $result.RollbackHint
                    $previousState = $result.PreviousState
                }
                catch {
                    $status = "Failed"
                    $message = $_.Exception.Message
                    $duration = [timespan]::Zero
                    $rollbackHint = Get-ModuleMetaValue -ActionKey $action -Name "RollbackHint" -DefaultValue ""
                    $previousState = ""
                    Write-Log "Error: $message" "Error"
                }

                Add-ModuleResult -Module $action -Status $status -Message $message -Duration $duration -RollbackHint $rollbackHint -PreviousState $previousState
                Write-Log "Result: $status in $([Math]::Round($duration.TotalSeconds, 2))s" $(if ($status -in @("Success", "Canceled")) { "Success" } else { "Error" })
                if ($rollbackHint) {
                    Write-Log "Rollback hint: $rollbackHint" "Warn"
                }
                if ($previousState) {
                    Write-Log "Previous state: $previousState" "Info"
                }

                $LastActionLabel = $curr.L
                $LastActionStatus = $status
                $LastActionMessage = $message
            }
            
            [Console]::SetCursorPosition(2, $maxY + 7)
            Write-Host "Press any key..." -F Gray
            $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
            Clear-Host
        }
        8 {
            if ($NumberInput.Length -gt 0) {
                $NumberInput = $NumberInput.Substring(0, $NumberInput.Length - 1)
            }
        }
        48 { if ($NumberInput.Length -lt 2) { $NumberInput += "0" } }
        49 { if ($NumberInput.Length -lt 2) { $NumberInput += "1" } }
        50 { if ($NumberInput.Length -lt 2) { $NumberInput += "2" } }
        51 { if ($NumberInput.Length -lt 2) { $NumberInput += "3" } }
        52 { if ($NumberInput.Length -lt 2) { $NumberInput += "4" } }
        53 { if ($NumberInput.Length -lt 2) { $NumberInput += "5" } }
        54 { if ($NumberInput.Length -lt 2) { $NumberInput += "6" } }
        55 { if ($NumberInput.Length -lt 2) { $NumberInput += "7" } }
        56 { if ($NumberInput.Length -lt 2) { $NumberInput += "8" } }
        57 { if ($NumberInput.Length -lt 2) { $NumberInput += "9" } }
        96 { if ($NumberInput.Length -lt 2) { $NumberInput += "0" } }
        97 { if ($NumberInput.Length -lt 2) { $NumberInput += "1" } }
        98 { if ($NumberInput.Length -lt 2) { $NumberInput += "2" } }
        99 { if ($NumberInput.Length -lt 2) { $NumberInput += "3" } }
        100 { if ($NumberInput.Length -lt 2) { $NumberInput += "4" } }
        101 { if ($NumberInput.Length -lt 2) { $NumberInput += "5" } }
        102 { if ($NumberInput.Length -lt 2) { $NumberInput += "6" } }
        103 { if ($NumberInput.Length -lt 2) { $NumberInput += "7" } }
        104 { if ($NumberInput.Length -lt 2) { $NumberInput += "8" } }
        105 { if ($NumberInput.Length -lt 2) { $NumberInput += "9" } }
        81 {
            Export-SessionResults
            Clear-Host
            exit
        }
        27 {
            Export-SessionResults
            Clear-Host
            exit
        }
    }
}
