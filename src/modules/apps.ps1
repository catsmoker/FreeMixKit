# ==============================================================================
# MODULE: Third-party software (Adobe, Spotify, Windows tweaks, Discord)
# ==============================================================================

$Modules["AdobeGenP"] = {
    Ensure-WingetInstalled
    Write-Log "Automating Adobe Creative Cloud installation via Winget..." "Info"
    Install-WingetPackage -Id "Adobe.CreativeCloud"

    # --- GenP Automation (Multi-Source Dynamic) ---
    $guideUrl = "https://wiki.dbzer0.com/genp-guides/guide/"
    $internxtUrl = "https://wiki.dbzer0.com/genp-guides/guide/#download-directory"

    # 1. Official Guide Version (Source of Truth)
    Write-Log "Checking official GenP guide for latest recommended version..." "Info"
    $targetVersion = "unknown"
    try {
        $page = Invoke-WebRequest -Uri $guideUrl -UseBasicParsing -ErrorAction Stop
        if ($page.Content -match "GENP v([0-9\.]+)") {
            $targetVersion = "v" + $Matches[1]
            Write-Log "Official Recommended Version: $targetVersion" "Success"
        }
    }
    catch {
        Write-Log "Could not reach guide website. Cannot verify latest version." "Warn"
    }

    if ($targetVersion -eq "unknown") {
        Write-Log "Opening official guide for manual download..." "Info"
        Start-Process $internxtUrl
        return
    }

    # 2. Choice: Automated vs Manual
    Write-Host ""
    if (-not (Confirm-Action "Would you like to attempt an automated download from community mirrors (GitHub)?`n(Selecting 'N' will open the official guide for manual download instead)")) {
        Write-Log "Opening official guide for manual download..." "Info"
        Start-Process $internxtUrl
        return
    }

    # 3. Search Sources for matching version
    $sources = @(
        @{ Name = "catsmoker (Primary)"; Api = "https://api.github.com/repos/catsmoker/FreeMixKit/contents/"; Type = "Contents" },
        @{ Name = "Mojszli"; Api = "https://api.github.com/repos/Mojszli/GenP/releases/tags/tag1"; Type = "Release" },
        @{ Name = "TheMythologist"; Api = "https://api.github.com/repos/TheMythologist/GenP/releases"; Type = "Releases" }
    )

    $downloadUrl = $null
    $foundFileName = $null

    foreach ($source in $sources) {
        Write-Log "Searching source: $($source.Name)..." "Info"
        try {
            $data = Invoke-RestMethod -Uri $source.Api -ErrorAction Stop

            if ($source.Type -eq "Contents") {
                $file = $data | Where-Object { $_.name -match "GenP-$([regex]::Escape($targetVersion))\.exe$" } | Select-Object -First 1
                if ($file) {
                    $downloadUrl = $file.download_url
                    $foundFileName = $file.name
                }
            }
            elseif ($source.Type -eq "Release") {
                $asset = $data.assets | Where-Object { $_.name -match "GenP-$([regex]::Escape($targetVersion))\.exe$" } | Select-Object -First 1
                if ($asset) {
                    $downloadUrl = $asset.browser_download_url
                    $foundFileName = $asset.name
                }
            }
            elseif ($source.Type -eq "Releases") {
                foreach ($rel in $data) {
                    $asset = $rel.assets | Where-Object { $_.name -match "GenP-$([regex]::Escape($targetVersion))\.exe$" } | Select-Object -First 1
                    if ($asset) {
                        $downloadUrl = $asset.browser_download_url
                        $foundFileName = $asset.name
                        break
                    }
                }
            }

            if ($downloadUrl) {
                Write-Log "Found $targetVersion at $($source.Name)!" "Success"
                break
            }
        }
        catch {
            Write-Log "Source $($source.Name) search failed." "Warn"
        }
    }

    # 4. Download and Run
    if ($downloadUrl) {
        $genpPath = Join-Path $DataDownloads $foundFileName
        if (-not (Test-Path $genpPath)) {
            Write-Log "Downloading $foundFileName..." "Info"
            Invoke-WebRequest -Uri $downloadUrl -OutFile $genpPath
        }
        Write-Log "Launching GenP Patcher..." "Success"
        Start-Process $genpPath
    }
    else {
        Write-Log "Could not find an automated download for $targetVersion in any source." "Warn"
        Write-Log "Opening official guide for manual download..." "Info"
        Start-Process $internxtUrl
    }
}

$Modules["AdobeUninstall"] = {
    $ccCleanerUrl = "https://swupmf.adobe.com/webfeed/CleanerTool/win/AdobeCreativeCloudCleanerTool.exe"
    $ccCleanerPath = Join-Path $DataDownloads "AdobeCreativeCloudCleanerTool.exe"

    Write-Log "Downloading Adobe Creative Cloud Cleaner Tool..." "Info"
    Invoke-WebRequest -Uri $ccCleanerUrl -OutFile $ccCleanerPath

    Write-Log "Launching Adobe Creative Cloud Cleaner Tool..." "Info"
    Start-Process $ccCleanerPath
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
    # Spotify -> Spicetify Installer
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

            $spotifyInstaller = "$DataDownloads\SpotifySetup.exe"
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

$Modules["SpotX"] = {
    $taskName = "FreeMixKit_SpotX_User"
    $currentUser = $env:USERNAME
    $spotxScript = Join-Path $DataTemp "Run-SpotX.ps1"

    $scriptContent = @'
Write-Host "=== SpotX Installer ===" -ForegroundColor Cyan
try {
    iex "& { $(iwr -useb 'https://raw.githubusercontent.com/SpotX-Official/SpotX/refs/heads/main/run.ps1') } -confirm_uninstall_ms_spoti -confirm_spoti_recomended_over -podcasts_off -block_update_on -start_spoti -new_theme -adsections_off -lyrics_stat spotify"
}
catch {
    Write-Host "SpotX failed: $($_.Exception.Message)" -ForegroundColor Red
    throw
}
Write-Host "Press ENTER to close..."
Read-Host
'@

    Set-Content -Path $spotxScript -Value $scriptContent -Force

    Write-Log "Launching SpotX as user: $currentUser..." "Info"

    $action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$spotxScript`""

    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(2)
    $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited

    try {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName

        Write-Log "SpotX launched successfully." "Success"
    }
    catch {
        Write-Log "Failed to launch SpotX: $($_.Exception.Message)" "Error"
        throw
    }
}

$Modules["Winhance"] = {
    $taskName = "FreeMixKit_Winhance_User"
    $currentUser = $env:USERNAME
    $winhanceScript = Join-Path $DataTemp "Run-Winhance.ps1"

    $scriptContent = @'
Write-Host "=== Winhance Launcher ===" -ForegroundColor Cyan
try {
    irm "https://get.winhance.net" | iex
}
catch {
    Write-Host "Winhance failed: $($_.Exception.Message)" -ForegroundColor Red
    throw
}
Write-Host "Press ENTER to close..."
Read-Host
'@

    Set-Content -Path $winhanceScript -Value $scriptContent -Force

    Write-Log "Launching Winhance as user: $currentUser..." "Info"

    $action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$winhanceScript`""

    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(2)
    $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited

    try {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName

        Write-Log "Winhance launched successfully." "Success"
    }
    catch {
        Write-Log "Failed to launch Winhance: $($_.Exception.Message)" "Error"
        throw
    }
}

$Modules["Legcord"] = {
    try {
        $u = ((Invoke-RestMethod "https://api.github.com/repos/Legcord/Legcord/releases/latest").assets | Where-Object name -match ".exe" | Select-Object -First 1).browser_download_url
        Invoke-WebRequest $u -OutFile "$DataDownloads\legcord.exe"; Start-Process "$DataDownloads\legcord.exe" -Wait
    }
    catch {
        Write-Log "Legcord install failed: $($_.Exception.Message)" "Error"
        throw
    }
}
