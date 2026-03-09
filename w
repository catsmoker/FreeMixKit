<#
.SYNOPSIS
    FreeMixKit v5.8
    Standalone system utility suite.

.NOTES
    Author: catsmoker (Refactored by Assistant)
    Privileges: Administrator Required
#>

# ==============================================================================
# 1. SETUP & ADMIN CHECK
# ==============================================================================

$Host.UI.RawUI.WindowTitle = "FreeMixKit v5.8"
try {
    # Force big window (120x40 is good for grid)
    $bufferSize = New-Object Management.Automation.Host.Size(120, 2000)
    $windowSize = New-Object Management.Automation.Host.Size(120, 40)
    if ($Host.UI.RawUI.BufferSize.Width -lt $windowSize.Width) { $Host.UI.RawUI.BufferSize = $bufferSize }
    $Host.UI.RawUI.WindowSize = $windowSize
    $Host.UI.RawUI.BufferSize = $bufferSize
}
catch { Write-Host "Resize not supported." -ForegroundColor DarkGray }

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
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
}

function Confirm-Action([string]$Prompt) {
    $response = Read-Host "$Prompt (Y/N)"
    return $response -match '^[Yy]$'
}

function Ensure-WingetInstalled {
    if (Get-Command winget -ErrorAction SilentlyContinue) { return }
    Write-Log "Winget not found. Installing Winget..." "Warn"
    $bundlePath = Join-Path $env:TEMP "winget.msixbundle"
    Invoke-WebRequest -Uri "https://aka.ms/getwinget" -OutFile $bundlePath
    Add-AppxPackage -Path $bundlePath
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
    ) -TimeoutSec 1800 | Out-Null
}

function Test-NetworkConnectivity {
    try {
        return [bool](Test-Connection -ComputerName "1.1.1.1" -Count 1 -Quiet -ErrorAction Stop)
    }
    catch {
        return $false
    }
}

function Invoke-ExternalCommand {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter()][string[]]$Arguments = @(),
        [Parameter()][int]$TimeoutSec = 1800
    )

    $proc = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -NoNewWindow
    $finished = Wait-Process -Id $proc.Id -Timeout $TimeoutSec -ErrorAction SilentlyContinue
    if (-not $finished) {
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        throw "Command timed out after $TimeoutSec sec: $FilePath $($Arguments -join ' ')"
    }

    if ($proc.ExitCode -ne 0) {
        throw "Command failed (exit $($proc.ExitCode)): $FilePath $($Arguments -join ' ')"
    }

    return [pscustomobject]@{
        FilePath = $FilePath
        Arguments = ($Arguments -join " ")
        ExitCode = $proc.ExitCode
    }
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Operation,
        [Parameter()][int]$MaxAttempts = 1,
        [Parameter()][int]$DelaySec = 2,
        [Parameter()][string]$OperationName = "Operation"
    )

    $attempt = 1
    while ($attempt -le $MaxAttempts) {
        try {
            return & $Operation
        }
        catch {
            if ($attempt -ge $MaxAttempts) { throw }
            Write-Log "$OperationName failed on attempt $attempt/$MaxAttempts. Retrying in $DelaySec sec. Error: $($_.Exception.Message)" "Warn"
            Start-Sleep -Seconds $DelaySec
            $attempt++
        }
    }
}

function Get-ModuleMetaValue([string]$ActionKey, [string]$Name, $DefaultValue) {
    if ($ModuleMeta.Contains($ActionKey) -and $ModuleMeta[$ActionKey].Contains($Name) -and $null -ne $ModuleMeta[$ActionKey][$Name]) {
        return $ModuleMeta[$ActionKey][$Name]
    }
    return $DefaultValue
}

function Test-DnsServersMatch([string[]]$ExpectedServers) {
    $expected = $ExpectedServers | Sort-Object -Unique
    $upAdapters = Get-NetAdapter | Where-Object Status -eq Up
    if (-not $upAdapters) { return $false }

    foreach ($adapter in $upAdapters) {
        $dns = (Get-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4).ServerAddresses
        $actual = $dns | Sort-Object -Unique
        $sameCount = @($actual).Count -eq @($expected).Count
        $sameItems = -not (Compare-Object -ReferenceObject $expected -DifferenceObject $actual)
        if (-not ($sameCount -and $sameItems)) { return $false }
    }
    return $true
}

function Test-TelemetryValue([int]$ExpectedValue) {
    $paths = @(
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection"
    )

    foreach ($path in $paths) {
        if (-not (Test-Path $path)) { return $false }
        $value = (Get-ItemProperty -Path $path -Name "AllowTelemetry" -ErrorAction SilentlyContinue).AllowTelemetry
        if ($value -ne $ExpectedValue) { return $false }
    }
    return $true
}

function Invoke-ModuleAction([string]$ActionKey) {
    if (-not $Modules.Contains($ActionKey)) {
        throw "Unknown action '$ActionKey'."
    }

    $label = Get-ModuleMetaValue -ActionKey $ActionKey -Name "Label" -DefaultValue $ActionKey
    $requiresNetwork = [bool](Get-ModuleMetaValue -ActionKey $ActionKey -Name "RequiresNetwork" -DefaultValue $false)
    $confirmMessage = Get-ModuleMetaValue -ActionKey $ActionKey -Name "ConfirmMessage" -DefaultValue $null
    $timeoutSec = [int](Get-ModuleMetaValue -ActionKey $ActionKey -Name "TimeoutSec" -DefaultValue 1800)
    $retryCount = [int](Get-ModuleMetaValue -ActionKey $ActionKey -Name "RetryCount" -DefaultValue 1)
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
        Invoke-WithRetry -MaxAttempts $retryCount -DelaySec 3 -OperationName $label -Operation {
            & $Modules[$ActionKey]
        } | Out-Null

        $duration = (Get-Date) - $start
        if ($duration.TotalSeconds -gt $timeoutSec) {
            throw "Module exceeded timeout budget (${timeoutSec}s): $label"
        }

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
$Script:TranscriptPath = $null

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
        TimeoutSec      = 1800
        RetryCount      = 1
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
    $resultsPath = Join-Path $env:USERPROFILE "Desktop\FreeMixKit_ModuleResults_$stamp.csv"

    if ($ModuleResults.Count -gt 0) {
        $ModuleResults | Export-Csv -Path $resultsPath -NoTypeInformation -Encoding UTF8
        Write-Log "Module results exported to: $resultsPath" "Success"
    }
    else {
        Write-Log "No module executions to export for this session." "Info"
    }

    if ($Script:TranscriptPath) {
        try {
            Stop-Transcript | Out-Null
            Write-Log "Transcript saved to: $($Script:TranscriptPath)" "Success"
        }
        catch {
            Write-Log "Unable to stop transcript cleanly: $($_.Exception.Message)" "Warn"
        }
    }
}

try {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $Script:TranscriptPath = Join-Path $env:USERPROFILE "Desktop\FreeMixKit_Session_$stamp.log"
    Start-Transcript -Path $Script:TranscriptPath -ErrorAction Stop | Out-Null
    Write-Log "Session transcript started: $Script:TranscriptPath" "Info"
}
catch {
    $Script:TranscriptPath = $null
    Write-Log "Transcript unavailable: $($_.Exception.Message)" "Warn"
}

# --- DEVELOPER ---
$Modules["DevChoice"] = {
    Write-Log "Starting Developer Environment Setup..." "Warn"
    Write-Log "Installing: VS Redists, .NET, Node, Python, Java, Tools, Bibata Cursor."
    
    # 1. Winget
    Ensure-WingetInstalled

    # 2. Packages
    $packages = @(
        "Microsoft.DotNet.SDK.10", "Microsoft.DotNet.Runtime.10", "OpenJS.NodeJS.LTS", "Python.Python.3", "EclipseAdoptium.Temurin.21.JDK",
        "Microsoft.PowerShell", "Git.Git", "Gyan.FFmpeg", "7zip.7zip", "Notepad++.Notepad++", "AdrienAllard.FileConverter", "Google.GeminiCLI",
        "Microsoft.VCRedist.2005.x86", "Microsoft.VCRedist.2005.x64", "Microsoft.VCRedist.2008.x86", "Microsoft.VCRedist.2008.x64",
        "Microsoft.VCRedist.2010.x86", "Microsoft.VCRedist.2010.x64", "Microsoft.VCRedist.2012.x86", "Microsoft.VCRedist.2012.x64",
        "Microsoft.VCRedist.2013.x86", "Microsoft.VCRedist.2013.x64", "Microsoft.VCRedist.2015+.x86", "Microsoft.VCRedist.2015+.x64"
    )

    foreach ($id in $packages) {
        try {
            Install-WingetPackage -Id $id
        }
        catch {
            Write-Log "Failed package '$id': $($_.Exception.Message)" "Error"
        }
    }

    # 3. Notepad Fix
    Write-Log "Replacing Notepad with Notepad++..."
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
[HKEY_CLASSES_ROOT\txtfile\ShellNew]
"NullFile"=""
"@
    $reg | Out-File "$env:TEMP\nppfix.reg" -Encoding ASCII -Force
    Start-Process reg.exe -Argument "import `"$env:TEMP\nppfix.reg`"" -Wait -NoNewWindow
    
    # 4. Bibata
    Write-Log "Installing Bibata Cursor..."
    try {
        $zip = "$env:TEMP\Bibata.zip"; $dest = "$env:TEMP\Bibata"
        Invoke-WebRequest "https://github.com/ful1e5/Bibata_Cursor/releases/download/v2.0.7/Bibata-Modern-Classic-Windows.zip" -OutFile $zip
        Expand-Archive $zip -Dest $dest -Force
        $inf = Get-ChildItem "$dest" -Recurse -Filter "*.inf" | Select-Object -First 1
        if ($inf) { Start-Process "RUNDLL32.EXE" -Arg "SETUPAPI.DLL,InstallHinfSection DefaultInstall 128 $($inf.FullName)" -Wait }
    }
    catch { Write-Log "Cursor Failed" "Error" }

    Write-Log "Done!" "Success"
}

# --- MAINTENANCE ---
$Modules["CleanSystem"] = {
    Write-Log "Cleaning..."
    Remove-Item "$env:TEMP\*" -Recurse -Force -EA SilentlyContinue
    Remove-Item "C:\Windows\Temp\*" -Recurse -Force -EA SilentlyContinue
    Remove-Item "C:\Windows\Prefetch\*" -Recurse -Force -EA SilentlyContinue
    Clear-DnsClientCache
    Write-Log "Cleaned." "Success"
}
$Modules["SystemRepair"] = {
    Invoke-ExternalCommand -FilePath "sfc.exe" -Arguments @("/scannow") -TimeoutSec 7200 | Out-Null
    Invoke-ExternalCommand -FilePath "DISM.exe" -Arguments @("/Online", "/Cleanup-Image", "/RestoreHealth") -TimeoutSec 7200 | Out-Null
    Write-Log "Done." "Success"
}
$Modules["MalwareScan"] = { 
    $mrt = "$env:SystemRoot\System32\MRT.exe"
    if (!(Test-Path $mrt)) { Invoke-WebRequest "https://go.microsoft.com/fwlink/?LinkID=212732" -OutFile "$env:TEMP\MRT.exe"; $mrt = "$env:TEMP\MRT.exe" }
    Start-Process $mrt -Wait 
}
$Modules["MalwareScanAdv"] = {
    Write-Host "WARNING: DOWNLOADING 400MB+ (Tron Script)" -Bx Red -Fx White
    try {
        $l = (Invoke-WebRequest "https://bmrf.org/repos/tron/" -UseBasicParsing).Links.href | Where-Object { $_ -match "Tron v.+?\.exe" } | Select-Object -First 1
        if ($l) {
            Invoke-WebRequest "https://bmrf.org/repos/tron/$l" -OutFile "$env:TEMP\$l"
            Start-Process "$env:TEMP\$l" -Wait
        }
        else {
            throw "Unable to discover Tron executable link."
        }
    }
    catch {
        Write-Log "Error fetching Tron: $($_.Exception.Message)" "Error"
        throw
    }
}
$Modules["SystemReport"] = { 
    $f = "$env:USERPROFILE\Desktop\SysReport.txt"
    "OS: $($SysInfo.OS)`nCPU: $($SysInfo.CPU)`nRAM: $($SysInfo.RAM)" | Out-File $f
    Invoke-Item $f 
}

# --- APPS ---
$Modules["AdobeFree"] = {
    Write-Log "Opening Creative Cloud..."
    Start-Process "https://www.adobe.com/download/creative-cloud"
    Write-Log "Opening GenP..."
    Start-Process "https://gen.paramore.su"
}
$Modules["SoftwareUpdate"] = {
    Ensure-WingetInstalled
    Invoke-ExternalCommand -FilePath "winget.exe" -Arguments @(
        "upgrade", "--all", "--include-unknown",
        "--accept-source-agreements", "--accept-package-agreements"
    ) -TimeoutSec 7200 | Out-Null
}
$Modules["SpotifyPro"] = {
    # ===============================
    # Spotify → Spicetify Installer
    # ===============================

    Write-Log "Checking Spotify installation..." "Info"

    $spotifyExe = "$env:APPDATA\Spotify\Spotify.exe"

    if (-not (Test-Path $spotifyExe)) {

        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Write-Log "Installing Spotify using winget..." "Info"
            Install-WingetPackage -Id "Spotify.Spotify"
        }
        else {
            Write-Log "Winget not found. Using direct installer..." "Info"

            $spotifyInstaller = "$env:TEMP\SpotifySetup.exe"
            Invoke-WebRequest "https://download.scdn.co/SpotifySetup.exe" -OutFile $spotifyInstaller

            Start-Process $spotifyInstaller -ArgumentList "/silent" -Wait
        }

        # Wait until Spotify exists
        for ($i = 0; $i -lt 15; $i++) {
            if (Test-Path $spotifyExe) { break }
            Start-Sleep 2
        }

        if (-not (Test-Path $spotifyExe)) {
            Write-Log "Spotify installation failed or timed out." "Error"
            return
        }
    }
    else {
        Write-Log "Spotify already installed." "Info"
    }

    # ===============================
    # Prepare Spicetify Installer
    # ===============================

    Write-Log "Preparing Spicetify Installation..." "Info"

    $tempDir = $env:TEMP
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

$Modules["DiscordPro"] = {
    try { 
        $u = ((Invoke-RestMethod "https://api.github.com/repos/Legcord/Legcord/releases/latest").assets | Where-Object name -match ".exe" | Select-Object -First 1).browser_download_url
        Invoke-WebRequest $u -OutFile "$env:TEMP\legcord.exe"; Start-Process "$env:TEMP\legcord.exe" -Wait
    }
    catch {
        Write-Log "Legcord install failed: $($_.Exception.Message)" "Error"
        throw
    }
}

# --- ACTIVATION ---
$Modules["ActivateWindows"] = { Invoke-RestMethod https://get.activated.win | Invoke-Expression }
$Modules["ActivateIDM"] = { Invoke-RestMethod https://coporton.com/ias | Invoke-Expression }

# --- TWEAKS & NETWORK ---
$Modules["ToggleUpdates"] = {
    try {
        $svc = Get-Service wuauserv
        $isDisabled = $svc.StartType -eq "Disabled"
        $target = if ($isDisabled) { "Enable (Manual)" } else { "Disable" }
        $Script:ModuleExecutionContext["PreviousState"] = "StartupType=$($svc.StartType), Status=$($svc.Status)"
        $Script:ModuleExecutionContext["TargetStartupType"] = if ($isDisabled) { "Manual" } else { "Disabled" }
        Write-Log "Current Windows Update StartupType: $($svc.StartType)" "Info"

        if ($isDisabled) {
            Set-Service wuauserv -StartupType Manual
            Start-Service wuauserv -ErrorAction SilentlyContinue
        }
        else {
            Stop-Service wuauserv -Force
            Set-Service wuauserv -StartupType Disabled
        }
        Write-Log "Windows Update changed. Rollback: run 'Toggle Updates' again." "Success"
    }
    catch {
        Write-Log "Toggle Updates failed: $($_.Exception.Message)" "Error"
        throw
    }
}
$Modules["ToggleDefender"] = {
    try {
        $p = Get-MpPreference
        $currentDisabled = [bool]$p.DisableRealtimeMonitoring
        $targetDisabled = -not $currentDisabled
        $targetLabel = if ($targetDisabled) { "Disabled" } else { "Enabled" }
        $Script:ModuleExecutionContext["PreviousState"] = "DisableRealtimeMonitoring=$currentDisabled"
        $Script:ModuleExecutionContext["TargetDefenderDisabled"] = $targetDisabled
        Write-Log "Current Defender real-time monitoring: $(if ($currentDisabled) { "Disabled" } else { "Enabled" })" "Info"

        Set-MpPreference -DisableRealtimeMonitoring $targetDisabled
        Write-Log "Defender real-time monitoring set to $targetLabel. Rollback: run 'Toggle Defender' again." "Success"
    }
    catch {
        Write-Log "Toggle Defender failed: $($_.Exception.Message)" "Error"
        throw
    }
}
$Modules["DisableTelemetry"] = {
    $paths = @(
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection"
    )

    try {
        $values = foreach ($path in $paths) {
            if (Test-Path $path) {
                (Get-ItemProperty -Path $path -Name "AllowTelemetry" -ErrorAction SilentlyContinue).AllowTelemetry
            }
            else {
                $null
            }
        }

        $currentlyDisabled = ($values | Where-Object { $_ -ne $null -and $_ -ne 0 }).Count -eq 0
        $targetValue = if ($currentlyDisabled) { 1 } else { 0 }
        $targetLabel = if ($targetValue -eq 0) { "Disabled (0)" } else { "Enabled (1)" }
        $Script:ModuleExecutionContext["PreviousState"] = "AllowTelemetryValues=$($values -join ',')"
        $Script:ModuleExecutionContext["TargetTelemetryValue"] = $targetValue
        Write-Log "Current telemetry policy: $(if ($currentlyDisabled) { "Disabled" } else { "Enabled/Partial" })" "Info"

        foreach ($path in $paths) {
            New-Item -Path $path -Force -EA SilentlyContinue | Out-Null
            Set-ItemProperty -Path $path -Name "AllowTelemetry" -Value $targetValue -Force
        }
        Write-Log "Telemetry policy updated. Rollback: run 'Disable Telemetry' again." "Success"
    }
    catch {
        Write-Log "Telemetry policy change failed: $($_.Exception.Message)" "Error"
        throw
    }
}
$Modules["RegistryOptimize"] = { Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" "NtfsDisableLastAccessUpdate" 1; Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" "SystemResponsiveness" 0 }
$Modules["SetGoogleDNS"] = { Get-NetAdapter | Where-Object Status -eq Up | Set-DnsClientServerAddress -ServerAddresses "8.8.8.8", "8.8.4.4" }
$Modules["SetCloudflareDNS"] = { Get-NetAdapter | Where-Object Status -eq Up | Set-DnsClientServerAddress -ServerAddresses "1.1.1.1", "1.0.0.1" }
$Modules["ResetNetwork"] = {
    Get-NetAdapter | Where-Object Status -eq Up | Set-DnsClientServerAddress -ResetServerAddresses
    Invoke-ExternalCommand -FilePath "ipconfig.exe" -Arguments @("/flushdns") -TimeoutSec 120 | Out-Null
    Invoke-ExternalCommand -FilePath "netsh.exe" -Arguments @("winsock", "reset") -TimeoutSec 120 | Out-Null
}

# --- UTILS ---
$Modules["CTTUtility"] = { Invoke-RestMethod https://christitus.com/win | Invoke-Expression }
$Modules["RegistryBackup"] = { Start-Process reg.exe -Arg "export HKLM `"$env:USERPROFILE\Desktop\Backup.reg`" /y" -Wait }
$Modules["FixResolution"] = { Invoke-WebRequest "https://www.monitortests.com/download/cru/cru-1.5.3.zip" -OutFile "$env:TEMP\cru.zip"; Expand-Archive "$env:TEMP\cru.zip" "$env:TEMP\CRU" -Force; Start-Process "$env:TEMP\CRU\CRU.exe" -Wait; Start-Process "$env:TEMP\CRU\restart64.exe" -Wait }
$Modules["RemoveWindowsAI"] = {
    try {
        & ([scriptblock]::Create((Invoke-RestMethod "https://raw.githubusercontent.com/zoicware/RemoveWindowsAI/main/RemoveWindowsAi.ps1"))) -nonInteractive
    }
    catch {
        Write-Log "Remove Windows AI failed: $($_.Exception.Message)" "Error"
        throw
    }
}

$Modules["AddShortcut"] = {
    $iconUrl = "https://catsmoker.github.io/freemixkit_icon.ico"
    $iconPath = "$env:USERPROFILE\Pictures\freemixkit_icon.ico"
    try {
        Invoke-WebRequest $iconUrl -OutFile $iconPath -ErrorAction SilentlyContinue
    }
    catch {
        Write-Log "Icon download failed, continuing without icon: $($_.Exception.Message)" "Warn"
    }

    $s = (New-Object -ComObject WScript.Shell).CreateShortcut("$env:USERPROFILE\Desktop\FreeMixKit.lnk")
    $s.TargetPath = "powershell.exe"
    $s.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"irm https://catsmoker.github.io/w | iex`""
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
Register-Module "DevChoice" "DEV CHOICE (Full)" "Installs: VS Redists, .NET, Node.js, Python, Java, PowerShell, Git, FFmpeg, 7zip, Notepad++, File Converter, GeminiCLI, Bibata Cursor." $Modules["DevChoice"] "Medium" @{
    RequiresNetwork = $true
    RetryCount = 3
    TimeoutSec = 7200
}
Register-Module "CleanSystem" "Clean System Junk" "Removes temp files, prefetch, and clears DNS cache." $Modules["CleanSystem"] "Low"
Register-Module "SystemRepair" "System Repair" "Runs SFC Scannow and DISM RestoreHealth." $Modules["SystemRepair"] "Low" @{
    TimeoutSec = 14400
}
Register-Module "MalwareScan" "Malware Scan" "Runs the built-in Microsoft Malicious Software Removal Tool." $Modules["MalwareScan"] "Low"
Register-Module "MalwareScanAdv" "Malware Scan Adv" "Downloads and runs Tron Script (Heavy/Advanced deep clean)." $Modules["MalwareScanAdv"] "Medium" @{
    RequiresNetwork = $true
    RetryCount = 3
    ConfirmMessage = "This downloads a large advanced cleaner package. Continue?"
    TimeoutSec = 7200
}
Register-Module "SystemReport" "System Report" "Generates a text file with system specs on your desktop." $Modules["SystemReport"] "Low"
Register-Module "ActivateWindows" "Activate Windows" "Runs MAS (Microsoft Activation Scripts) to activate Windows." $Modules["ActivateWindows"] "High"
Register-Module "ActivateIDM" "Activate IDM" "Activates Internet Download Manager (IDM)." $Modules["ActivateIDM"] "High"
Register-Module "ToggleUpdates" "Toggle Updates" "Enables or Disables Windows Update service." $Modules["ToggleUpdates"] "High" @{
    ConfirmMessage = "This changes Windows Update service behavior. Continue?"
    RollbackHint = "Run Toggle Updates again to restore the previous mode."
    Verify = {
        $svc = Get-Service wuauserv
        $target = $Script:ModuleExecutionContext["TargetStartupType"]
        return ($null -ne $target) -and ($svc.StartType.ToString() -eq $target)
    }
}
Register-Module "ToggleDefender" "Toggle Defender" "Toggles Real-time monitoring for Windows Defender." $Modules["ToggleDefender"] "High" @{
    ConfirmMessage = "This changes Defender real-time protection. Continue?"
    RollbackHint = "Run Toggle Defender again to revert."
    Verify = {
        $prefs = Get-MpPreference
        $target = [bool]$Script:ModuleExecutionContext["TargetDefenderDisabled"]
        return ([bool]$prefs.DisableRealtimeMonitoring) -eq $target
    }
}
Register-Module "DisableTelemetry" "Disable Telemetry" "Toggles Windows data collection policies." $Modules["DisableTelemetry"] "High" @{
    ConfirmMessage = "This changes telemetry policy registry values. Continue?"
    RollbackHint = "Run Disable Telemetry again to toggle back."
    Verify = {
        $target = [int]$Script:ModuleExecutionContext["TargetTelemetryValue"]
        return Test-TelemetryValue -ExpectedValue $target
    }
}
Register-Module "RegistryOptimize" "Registry Optimize" "Tweaks NTFS access updates and System Responsiveness." $Modules["RegistryOptimize"] "Medium"
Register-Module "AdobeFree" "Adobe Free (GenP)" "Downloads Creative Cloud and GenP activator." $Modules["AdobeFree"] "High" @{
    RequiresNetwork = $true
    RetryCount = 3
}
Register-Module "SoftwareUpdate" "Software Update" "Upgrades all installed software via Winget." $Modules["SoftwareUpdate"] "Low" @{
    RequiresNetwork = $true
    RetryCount = 3
    TimeoutSec = 7200
    Verify = { [bool](Get-Command winget -ErrorAction SilentlyContinue) }
}
Register-Module "SpotifyPro" "Spotify Pro" "Installs Spicetify for Spotify customization/ad-blocking." $Modules["SpotifyPro"] "Medium" @{
    RequiresNetwork = $true
    RetryCount = 3
    TimeoutSec = 3600
}
Register-Module "DiscordPro" "Discord Pro" "Installs Legcord (BetterDiscord alternative)." $Modules["DiscordPro"] "Low" @{
    RequiresNetwork = $true
    RetryCount = 3
}
Register-Module "SetGoogleDNS" "Set Google DNS" "Sets DNS to 8.8.8.8 / 8.8.4.4." $Modules["SetGoogleDNS"] "Low" @{
    RollbackHint = "Run Reset Network or choose a different DNS profile."
    Verify = { Test-DnsServersMatch -ExpectedServers @("8.8.8.8", "8.8.4.4") }
}
Register-Module "SetCloudflareDNS" "Set Cloudflare DNS" "Sets DNS to 1.1.1.1 / 1.0.0.1." $Modules["SetCloudflareDNS"] "Low" @{
    RollbackHint = "Run Reset Network or choose a different DNS profile."
    Verify = { Test-DnsServersMatch -ExpectedServers @("1.1.1.1", "1.0.0.1") }
}
Register-Module "ResetNetwork" "Reset Network" "Resets DNS and Winsock settings." $Modules["ResetNetwork"] "Low"
Register-Module "CTTUtility" "CTT WinUtil" "Launches Chris Titus Tech's Windows Utility." $Modules["CTTUtility"] "Medium" @{
    RequiresNetwork = $true
    RetryCount = 3
}
Register-Module "RegistryBackup" "Registry Backup" "Backs up the HKLM registry hive to Desktop." $Modules["RegistryBackup"] "Low"
Register-Module "FixResolution" "Fix Resolution" "Uses CRU to restart graphics driver and fix resolution." $Modules["FixResolution"] "Medium" @{
    RequiresNetwork = $true
    RetryCount = 3
}
Register-Module "RemoveWindowsAI" "Remove Windows AI" "Removes Copilot and Recall features." $Modules["RemoveWindowsAI"] "High" @{
    RequiresNetwork = $true
    RetryCount = 3
    ConfirmMessage = "This applies system-level AI feature removals. Continue?"
}
Register-Module "AddShortcut" "Add Shortcut" "Creates a shortcut for this script on the Desktop." $Modules["AddShortcut"] "Low"

# ==============================================================================
# 5. GRID MENU CONFIGURATION
# ==============================================================================

# Define Columns. Type: H=Header, I=Item
$Col1 = @(
    @{T = "H"; L = "[ DEVELOPER ]" }
    @{T = "I"; L = "DEV CHOICE (Full)"; A = "DevChoice"; D = "Installs: VS Redists, .NET, Node.js, Python, Java, PowerShell, Git, FFmpeg, 7zip, Notepad++, File Converter, GeminiCLI, Bibata Cursor." }
    @{T = "H"; L = "" }
    @{T = "H"; L = "[ MAINTENANCE ]" }
    @{T = "I"; L = "Clean System Junk"; A = "CleanSystem"; D = "Removes temp files, prefetch, and clears DNS cache." }
    @{T = "I"; L = "System Repair"; A = "SystemRepair"; D = "Runs SFC Scannow and DISM RestoreHealth." }
    @{T = "I"; L = "Malware Scan"; A = "MalwareScan"; D = "Runs the built-in Microsoft Malicious Software Removal Tool." }
    @{T = "I"; L = "Malware Scan Adv"; A = "MalwareScanAdv"; D = "Downloads and runs Tron Script (Heavy/Advanced deep clean)." }
    @{T = "I"; L = "System Report"; A = "SystemReport"; D = "Generates a text file with system specs on your desktop." }
    @{T = "H"; L = "" }
    @{T = "H"; L = "[ ACTIVATION ]" }
    @{T = "I"; L = "Activate Windows"; A = "ActivateWindows"; D = "Runs MAS (Microsoft Activation Scripts) to activate Windows." }
    @{T = "I"; L = "Activate IDM"; A = "ActivateIDM"; D = "Activates Internet Download Manager (IDM)." }
    @{T = "H"; L = "" }
    @{T = "H"; L = "[ TWEAKS ]" }
    @{T = "I"; L = "Toggle Updates"; A = "ToggleUpdates"; D = "Enables or Disables Windows Update service." }
    @{T = "I"; L = "Toggle Defender"; A = "ToggleDefender"; D = "Toggles Real-time monitoring for Windows Defender." }
    @{T = "I"; L = "Disable Telemetry"; A = "DisableTelemetry"; D = "Disables Windows data collection policies." }
    @{T = "I"; L = "Registry Optimize"; A = "RegistryOptimize"; D = "Tweaks NTFS access updates and System Responsiveness." }
)

$Col2 = @(
    @{T = "H"; L = "[ SOFTWARE ]" }
    @{T = "I"; L = "Adobe Free (GenP)"; A = "AdobeFree"; D = "Downloads Creative Cloud and GenP activator." }
    @{T = "I"; L = "Software Update"; A = "SoftwareUpdate"; D = "Upgrades all installed software via Winget." }
    @{T = "I"; L = "Spotify Pro"; A = "SpotifyPro"; D = "Installs Spicetify for Spotify customization/ad-blocking." }
    @{T = "I"; L = "Discord Pro"; A = "DiscordPro"; D = "Installs Legcord (BetterDiscord alternative)." }
    @{T = "H"; L = "" }
    @{T = "H"; L = "[ NETWORK ]" }
    @{T = "I"; L = "Set Google DNS"; A = "SetGoogleDNS"; D = "Sets DNS to 8.8.8.8 / 8.8.4.4." }
    @{T = "I"; L = "Set Cloudflare DNS"; A = "SetCloudflareDNS"; D = "Sets DNS to 1.1.1.1 / 1.0.0.1." }
    @{T = "I"; L = "Reset Network"; A = "ResetNetwork"; D = "Resets DNS and Winsock settings." }
    @{T = "H"; L = "" }
    @{T = "H"; L = "[ UTILITIES ]" }
    @{T = "I"; L = "CTT WinUtil"; A = "CTTUtility"; D = "Launches Chris Titus Tech's Windows Utility." }
    @{T = "I"; L = "Registry Backup"; A = "RegistryBackup"; D = "Backs up the HKLM registry hive to Desktop." }
    @{T = "I"; L = "Fix Resolution"; A = "FixResolution"; D = "Uses CRU to restart graphics driver and fix resolution." }
    @{T = "H"; L = "" }
    @{T = "H"; L = "[ AI & EXIT ]" }
    @{T = "I"; L = "Remove Windows AI"; A = "RemoveWindowsAI"; D = "Removes Copilot and Recall features." }
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

$SelIdx = 0 # Index in $NavItems
$LastActionLabel = "None"
$LastActionStatus = "N/A"
$LastActionDuration = "0.00s"
$LastActionAt = "-"
$LastActionMessage = "No action run yet."

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
    Write-Host " FREEMIXKIT v5.7" -NoNewline -F Cyan
    Write-Host " | $timeNow | ARROWS Navigate | ENTER Run | Q/ESC Exit" -F Gray
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
            $label = $item.L
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
            $label = $item.L
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
    $needsNet = if ($currMeta) { [bool]$currMeta.RequiresNetwork } else { $false }
    $riskColor = switch ($risk) {
        "High" { "Red" }
        "Medium" { "Yellow" }
        "Low" { "Green" }
        default { "DarkGray" }
    }

    for ($lineIdx = 1; $lineIdx -le 6; $lineIdx++) {
        [Console]::SetCursorPosition(0, $maxY + $lineIdx)
        Write-Host (" " * ($uiWidth - 1))
    }

    [Console]::SetCursorPosition(2, $maxY + 1)
    $infoText = $curr.D
    if ($infoText.Length -gt ($uiWidth - 10)) { $infoText = $infoText.Substring(0, $uiWidth - 13) + "..." }
    Write-Host "INFO: $infoText" -F Yellow
    [Console]::SetCursorPosition(2, $maxY + 2)
    Write-Host "SELECTED: $($curr.L) | Risk: " -NoNewline -F Gray
    Write-Host $risk -NoNewline -F $riskColor
    Write-Host " | Network: $(if ($needsNet) { "Required" } else { "No" })" -F Gray
    [Console]::SetCursorPosition(2, $maxY + 3)
    Write-Host "LAST: $LastActionLabel | Status: $LastActionStatus | Duration: $LastActionDuration | At: $LastActionAt" -F DarkGray
    [Console]::SetCursorPosition(2, $maxY + 4)
    $lastMsg = $LastActionMessage
    if ($lastMsg.Length -gt ($uiWidth - 17)) { $lastMsg = $lastMsg.Substring(0, $uiWidth - 20) + "..." }
    Write-Host "LAST MESSAGE: $lastMsg" -F DarkGray
    [Console]::SetCursorPosition(2, $maxY + 5)
    Write-Host "KEYS: Up/Down move | Left/Right switch column | Enter execute | Q or Esc exit" -F DarkGray

    # INPUT HANDLING
    $k = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

    switch ($k.VirtualKeyCode) {
        38 {
            # UP
            # Find prev item in same column
            $prev = $NavItems | Where-Object { $_.C -eq $curr.C -and $_.R -lt $curr.R } | Select-Object -Last 1
            if ($prev) { $SelIdx = $NavItems.IndexOf($prev) }
        }
        40 {
            # DOWN
            # Find next item in same column
            $next = $NavItems | Where-Object { $_.C -eq $curr.C -and $_.R -gt $curr.R } | Select-Object -First 1
            if ($next) { $SelIdx = $NavItems.IndexOf($next) }
        }
        39 {
            # RIGHT
            if ($curr.C -eq 0) {
                # Jump to Col 1, similar Row Y
                $target = $NavItems | Where-Object { $_.C -eq 1 } | Sort-Object { [Math]::Abs($_.R - $curr.R) } | Select-Object -First 1
                if ($target) { $SelIdx = $NavItems.IndexOf($target) }
            }
        }
        37 {
            # LEFT
            if ($curr.C -eq 1) {
                # Jump to Col 0, similar Row Y
                $target = $NavItems | Where-Object { $_.C -eq 0 } | Sort-Object { [Math]::Abs($_.R - $curr.R) } | Select-Object -First 1
                if ($target) { $SelIdx = $NavItems.IndexOf($target) }
            }
        }
        13 {
            # ENTER
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
                $LastActionDuration = "{0:N2}s" -f $duration.TotalSeconds
                $LastActionAt = (Get-Date).ToString("HH:mm:ss")
                $LastActionMessage = $message
            }
            
            [Console]::SetCursorPosition(2, $maxY + 7)
            Write-Host "Press any key..." -F Gray
            $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
            Clear-Host
        }
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
