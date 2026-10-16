# ==============================================================================
# MODULE: Utilities, fixes, and links
# ==============================================================================

$Modules["WinUtil"] = {
    $shellExe = if (Get-Command pwsh.exe -ErrorAction SilentlyContinue) { "pwsh.exe" } else { "powershell.exe" }
    $isWindows11 = Test-IsWindows11
    $scriptUrl = if ($isWindows11) { $WinUtilLatestScriptUrl } else { $WinUtilLegacyWindows10ScriptUrl }
    $scriptFlavor = if ($isWindows11) { "latest Windows 11 build" } else { "pinned Windows 10-compatible build (25.10.06)" }
    $scriptFileName = if ($isWindows11) { "winutil-latest.ps1" } else { "winutil-25.10.06.ps1" }
    $scriptPath = Join-Path $DataTemp $scriptFileName

    try {
        if (-not $isWindows11) {
            Write-Log "Windows 10 detected. WinUtil dropped Windows 10 support on 2025-10-14; launching pinned build 25.10.06 instead." "Warn"
        }

        Write-Log "Downloading WinUtil ($scriptFlavor)..." "Info"
        Invoke-WebRequest -Uri $scriptUrl -OutFile $scriptPath

        Write-Log "Launching WinUtil..." "Info"
        Start-Process -FilePath $shellExe -ArgumentList @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", ('"{0}"' -f $scriptPath)
        ) -Wait
    }
    catch {
        Write-Log "WinUtil launch failed: $($_.Exception.Message)" "Error"
        throw
    }
}

$Modules["FixResolution"] = {
    $cruThreadUrl = "https://www.monitortests.com/forum/Thread-Custom-Resolution-Utility-CRU"
    $fallbackCruZipUrl = "https://www.monitortests.com/download/cru/cru-1.5.3.zip"
    $cruZipUrl = $fallbackCruZipUrl
    $cruZipPath = Join-Path $DataDownloads "cru.zip"
    $cruExtractPath = Join-Path $DataTemp "CRU"

    try {
        Write-Log "Resolving latest CRU download link..." "Info"
        $cruRequestParams = @{
            Uri = $cruThreadUrl
        }
        if ($PSVersionTable.PSEdition -ne "Core") {
            $cruRequestParams.UseBasicParsing = $true
        }
        $cruPage = Invoke-WebRequest @cruRequestParams
        $cruMatch = [regex]::Match($cruPage.Content, 'https://www\.monitortests\.com/download/cru/cru-[0-9][0-9A-Za-z\.\-]*\.zip')
        if ($cruMatch.Success) {
            $cruZipUrl = $cruMatch.Value
        }
    }
    catch {
        Write-Log "Could not resolve latest CRU link automatically. Using fallback package URL." "Warn"
    }

    Write-Log "Downloading CRU from: $cruZipUrl" "Info"
    Invoke-WebRequest -Uri $cruZipUrl -OutFile $cruZipPath
    Expand-Archive $cruZipPath $cruExtractPath -Force
    Start-Process (Join-Path $cruExtractPath "CRU.exe") -Wait
    Start-Process (Join-Path $cruExtractPath "restart64.exe") -Wait
}

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

$Modules["MemReduct"] = {
    $candidatePaths = @()
    if ($env:ProgramFiles) {
        $candidatePaths += Join-Path $env:ProgramFiles "Mem Reduct\memreduct.exe"
    }
    if (${env:ProgramFiles(x86)}) {
        $candidatePaths += Join-Path ${env:ProgramFiles(x86)} "Mem Reduct\memreduct.exe"
    }
    Install-And-LaunchWingetPackage -Id "Henry++.MemReduct" -ExecutableNames @("memreduct.exe") -CandidatePaths $candidatePaths
}

$Modules["WinMemoryCleaner"] = {
    Install-And-LaunchWingetPackage -Id "IgorMundstein.WinMemoryCleaner" -ExecutableNames @("WinMemoryCleaner.exe")
}

$Modules["OpenGitHub"] = { Start-Process $FmkRepoUrl }
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
