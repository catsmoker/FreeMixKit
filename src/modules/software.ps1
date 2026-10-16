# ==============================================================================
# MODULES: Software catalog (browsers, media, gaming, developer apps)
# ==============================================================================
# One module per application. WinGet is the default source: every ID below was
# verified with `winget search --id <Id> --exact` (and `winget show` for the
# publisher on the less obvious ones) before it was written here. Do not invent
# IDs -- re-run the search if one stops resolving.
#
# When WinGet fails, the module opens the vendor's official download page and
# fails loudly rather than falling back to an unrelated package.

function Install-FmkCatalogApp {
    <#
        Installs one catalog application via WinGet, opening the official
        download page when WinGet cannot complete the install.

        Kept as one helper rather than repeated per module so the failure
        behaviour cannot drift: a WinGet failure is logged with the package ID,
        the official page is opened for a manual install, and the module still
        reports Failed with an actionable message instead of a false Success.
    #>
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][string]$WingetId,
        [Parameter(Mandatory)][string]$FallbackUrl
    )

    Ensure-WingetInstalled

    try {
        Install-WingetPackage -Id $WingetId
    }
    catch {
        Write-Log "WinGet install failed for $Label ($WingetId): $($_.Exception.Message)" "Warn"
        Write-Log "Opening official download page instead: $FallbackUrl" "Info"
        Start-Process $FallbackUrl
        throw "WinGet install failed for $Label; opened the official download page ($FallbackUrl)."
    }

    Write-Log "$Label installed via WinGet ($WingetId)." "Success"
}

# ------------------------------------------------------------------------------
# Brave Browser
# ------------------------------------------------------------------------------
$Modules["App_Brave"] = {
    Write-Log "Installing Brave Browser..." "Info"
    Install-FmkCatalogApp -Label "Brave Browser" -WingetId "Brave.Brave" -FallbackUrl "https://brave.com/download/"
    Write-Log "Brave Browser done." "Success"
}

# ------------------------------------------------------------------------------
# Telegram Desktop
# ------------------------------------------------------------------------------
$Modules["App_Telegram"] = {
    Write-Log "Installing Telegram Desktop..." "Info"
    Install-FmkCatalogApp -Label "Telegram Desktop" -WingetId "Telegram.TelegramDesktop" -FallbackUrl "https://desktop.telegram.org/"
    Write-Log "Telegram Desktop done." "Success"
}

# ------------------------------------------------------------------------------
# VLC Media Player
# ------------------------------------------------------------------------------
$Modules["App_VLC"] = {
    Write-Log "Installing VLC Media Player..." "Info"
    Install-FmkCatalogApp -Label "VLC Media Player" -WingetId "VideoLAN.VLC" -FallbackUrl "https://www.videolan.org/vlc/"
    Write-Log "VLC Media Player done." "Success"
}

# ------------------------------------------------------------------------------
# LocalSend
# ------------------------------------------------------------------------------
$Modules["App_LocalSend"] = {
    Write-Log "Installing LocalSend..." "Info"
    Install-FmkCatalogApp -Label "LocalSend" -WingetId "LocalSend.LocalSend" -FallbackUrl "https://localsend.org/"
    Write-Log "LocalSend done." "Success"
}

# ------------------------------------------------------------------------------
# qBittorrent
# ------------------------------------------------------------------------------
$Modules["App_QBittorrent"] = {
    Write-Log "Installing qBittorrent..." "Info"
    Install-FmkCatalogApp -Label "qBittorrent" -WingetId "qBittorrent.qBittorrent" -FallbackUrl "https://www.qbittorrent.org/download"
    Write-Log "qBittorrent done." "Success"
}

# ------------------------------------------------------------------------------
# Visual Studio (Community, default edition)
# ------------------------------------------------------------------------------
$Modules["App_VisualStudio"] = {
    # The full IDE, not Visual Studio Code. Community is the default edition;
    # other editions are separate WinGet IDs and stay out of this module.
    Write-Log "Installing Visual Studio Community..." "Info"
    Install-FmkCatalogApp -Label "Visual Studio Community" -WingetId "Microsoft.VisualStudio.2022.Community" -FallbackUrl "https://visualstudio.microsoft.com/downloads/"
    Write-Log "Visual Studio Community done." "Success"
}

# ------------------------------------------------------------------------------
# Stremio
# ------------------------------------------------------------------------------
$Modules["App_Stremio"] = {
    Write-Log "Installing Stremio..." "Info"
    Install-FmkCatalogApp -Label "Stremio" -WingetId "Stremio.Stremio" -FallbackUrl "https://www.stremio.com/downloads"
    Write-Log "Stremio done." "Success"
}

# ------------------------------------------------------------------------------
# Bulk Crap Uninstaller
# ------------------------------------------------------------------------------
$Modules["App_BCUninstaller"] = {
    Write-Log "Installing Bulk Crap Uninstaller..." "Info"
    Install-FmkCatalogApp -Label "Bulk Crap Uninstaller" -WingetId "Klocman.BulkCrapUninstaller" -FallbackUrl "https://github.com/Klocman/Bulk-Crap-Uninstaller/releases"
    Write-Log "Bulk Crap Uninstaller done." "Success"
}

# ------------------------------------------------------------------------------
# Steam
# ------------------------------------------------------------------------------
$Modules["App_Steam"] = {
    Write-Log "Installing Steam..." "Info"
    Install-FmkCatalogApp -Label "Steam" -WingetId "Valve.Steam" -FallbackUrl "https://store.steampowered.com/about/"
    Write-Log "Steam done." "Success"
}

# ------------------------------------------------------------------------------
# HandBrake
# ------------------------------------------------------------------------------
$Modules["App_HandBrake"] = {
    Write-Log "Installing HandBrake..." "Info"
    Install-FmkCatalogApp -Label "HandBrake" -WingetId "HandBrake.HandBrake" -FallbackUrl "https://handbrake.fr/downloads.php"
    Write-Log "HandBrake done." "Success"
}

# ------------------------------------------------------------------------------
# Android Studio
# ------------------------------------------------------------------------------
$Modules["App_AndroidStudio"] = {
    Write-Log "Installing Android Studio..." "Info"
    Install-FmkCatalogApp -Label "Android Studio" -WingetId "Google.AndroidStudio" -FallbackUrl "https://developer.android.com/studio"
    Write-Log "Android Studio done." "Success"
}

# ------------------------------------------------------------------------------
# Godot Engine
# ------------------------------------------------------------------------------
$Modules["App_Godot"] = {
    Write-Log "Installing Godot Engine..." "Info"
    Install-FmkCatalogApp -Label "Godot Engine" -WingetId "GodotEngine.GodotEngine" -FallbackUrl "https://godotengine.org/download"
    Write-Log "Godot Engine done." "Success"
}

# ------------------------------------------------------------------------------
# Unity Hub
# ------------------------------------------------------------------------------
$Modules["App_UnityHub"] = {
    # Unity Hub only. The Unity Editor installs through the Hub's own workflow,
    # so this module deliberately does not attempt a standalone Editor install.
    Write-Log "Installing Unity Hub..." "Info"
    Install-FmkCatalogApp -Label "Unity Hub" -WingetId "Unity.UnityHub" -FallbackUrl "https://unity.com/download"
    Write-Log "Unity Hub done." "Success"
}

# ------------------------------------------------------------------------------
# ZapFast (unofficial WhatsApp desktop client)
# ------------------------------------------------------------------------------
$Modules["App_ZapFast"] = {
    # No verified WinGet package exists (checked: `winget search ZapFast`
    # returns nothing), so this installs from the official GitHub releases at
    # https://github.com/crmne/zapfast. Never compiled from source while
    # official binary installers are published.
    Write-Log "Installing ZapFast (unofficial WhatsApp desktop client)..." "Info"
    Write-Log "Note: ZapFast is a third-party client, not endorsed by WhatsApp or Meta. Unofficial clients can carry compatibility, privacy, and account-restriction risks." "Warn"

    $releasesPage = "https://github.com/crmne/zapfast/releases/latest"

    try {
        $arch = $env:PROCESSOR_ARCHITECTURE
        $want = if ($arch -eq "ARM64") { "aarch64-pc-windows-msvc-setup.exe" } else { "x86_64-pc-windows-msvc-setup.exe" }

        Write-Log "Checking official ZapFast releases..." "Info"
        $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/crmne/zapfast/releases/latest" -ErrorAction Stop
        $asset = @($rel.assets | Where-Object { $_.name -like "*$want" }) | Select-Object -First 1
        if (-not $asset) {
            throw "Latest release ($($rel.tag_name)) has no Windows installer matching '$want'."
        }

        $file = Join-Path $DataDownloads $asset.name
        Write-Log "Downloading $($asset.name) ($([Math]::Round($asset.size / 1MB, 1)) MB) from official release $($rel.tag_name)..." "Info"
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $file -UseBasicParsing -ErrorAction Stop

        # Verify SHA-256. The release API already carries each asset's digest,
        # so that is the primary source; checksums.txt is the fallback. (Its
        # download is decoded by hand because on Windows PowerShell 5.1
        # Invoke-WebRequest hands back Byte[] rather than String here.)
        $expected = $null
        if ($asset.digest -match '^sha256:([0-9a-fA-F]{64})$') {
            $expected = $Matches[1].ToLowerInvariant()
        }
        else {
            try {
                $sumsUrl = "https://github.com/crmne/zapfast/releases/download/$($rel.tag_name)/checksums.txt"
                $raw = (Invoke-WebRequest -Uri $sumsUrl -UseBasicParsing -ErrorAction Stop).Content
                $text = if ($raw -is [string]) { $raw } else { [Text.Encoding]::UTF8.GetString($raw) }
                $line = @($text -split "`r?`n" | Where-Object { $_ -match [regex]::Escape($asset.name) }) | Select-Object -First 1
                if ($line -and ($line -match "([0-9a-fA-F]{64})")) {
                    $expected = $Matches[1].ToLowerInvariant()
                }
            }
            catch {
                Write-Log "Could not fetch checksums: $($_.Exception.Message)" "Warn"
            }
        }

        if ($expected) {
            $actual = (Get-FileHash -LiteralPath $file -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
            if ($actual -ne $expected) {
                throw "Checksum mismatch for $($asset.name): expected $expected, got $actual."
            }
            Write-Log "Checksum verified (SHA-256)." "Success"
        }
        else {
            Write-Log "No checksum available for $($asset.name); skipping verification." "Warn"
        }

        Write-Log "Running ZapFast installer..." "Info"
        Start-Process -FilePath $file -Wait
    }
    catch {
        Write-Log "ZapFast install failed: $($_.Exception.Message)" "Error"
        Write-Log "Opening official releases page instead: $releasesPage" "Info"
        Start-Process $releasesPage
        throw "ZapFast install failed; opened the official releases page ($releasesPage)."
    }

    Write-Log "ZapFast done." "Success"
}
