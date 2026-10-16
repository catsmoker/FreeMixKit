# ==============================================================================
# MODULE REGISTRATION
# ==============================================================================
# Single source of truth for labels, descriptions, risk, and UI placement.
# Category = sidebar entry, Group = section heading within that page.
#
# The split between Software and Utilities, applied throughout:
#   Software   - standalone applications you install and then use.
#   Utilities  - tools for managing, optimising, configuring, or maintaining
#                the system itself.
# Activation and About keep their own categories: Activation is where the
# licensing entries live, About is product info and support, not software.

# Risk ratings: None = no warning warranted (no badge). Low/Medium/High carry
# RiskReason plus dimensions; only Medium/High draw a badge. See
# src/functions/08-risk.ps1 for level criteria. Reviewed 2026-10-16 against the
# module bodies in src/modules/ and the linked publishers. No audit is claimed.

# --- System -------------------------------------------------------------------
Register-Module "CleanSystem" "Clean System Junk" "Removes temp files, prefetch, and clears DNS cache." $Modules["CleanSystem"] "None" @{
    Category    = "System"
    Group       = "Cleanup & Repair"
    ActionLabel = "Run"
}

Register-Module "SystemRepair" "System Repair" "Runs SFC Scannow and DISM RestoreHealth." $Modules["SystemRepair"] "Low" @{
    Category       = "System"
    Group          = "Cleanup & Repair"
    ActionLabel    = "Run"
    RiskReason     = "Official Microsoft repair tools; slow, admin, touches system files."
    RiskTechnical  = "Low: SFC/DISM are inbox repair tools with predictable behavior."
    RiskSource     = "High confidence: inbox Windows tools, no download."
    RiskPrivacy    = "None identified."
    RiskLicense    = "None."
    RiskPrivileges = "Requires admin; modifies protected system files."
    RiskReviewed   = "2026-10-16"
}

Register-Module "MalwareScan" "Malware Scan" "Runs MRT, then downloads and runs Kaspersky Virus Removal Tool (KVRT)." $Modules["MalwareScan"] "Low" @{
    RequiresNetwork = $true
    Category        = "System"
    Group           = "Cleanup & Repair"
    ActionLabel     = "Run"
    RiskReason      = "Downloads and runs KVRT scanner with admin rights."
    RiskTechnical   = "Low: MRT is inbox; KVRT is a vendor scanner, well-understood behavior."
    RiskSource      = "Medium-high: MRT inbox; KVRT from Kaspersky dev-builds host."
    RiskPrivacy     = "Low: on-demand scanner; scan results stay local unless shared."
    RiskLicense     = "None."
    RiskPrivileges  = "Requires admin; full-system scan."
    RiskEvidence    = "https://devbuilds.s.kaspersky-labs.com/kvrt/latest/full/kvrt.exe"
    RiskReviewed    = "2026-10-16"
}

Register-Module "SystemReport" "System Report" "Collects the full hardware and Windows report and saves it on your desktop." $Modules["SystemReport"] "None" @{
    Category    = "System"
    Group       = "Cleanup & Repair"
    ActionLabel = "Run"
}

# --- Activation ---------------------------------------------------------------
# MAS/IAS/GenP are High because of licensing impact plus remote-code and
# provenance concerns, each documented separately below. High is never assigned
# for licensing alone; the technical and source dimensions carry the rating.
Register-Module "MAS" "Activate Windows & Office" "Runs MAS (Microsoft Activation Scripts) to activate Windows and Office." $Modules["MAS"] "High" @{
    Category       = "Activation"
    Group          = "Microsoft"
    ActionLabel    = "Run"
    ConfirmMessage = "MAS activates Windows/Office outside Microsoft licensing.`nThis may violate license terms or local law, modifies licensing state, requires admin, and runs remote code from get.activated.win.`nContinue only where legal and at your own risk."
    RiskReason     = "Bypasses Microsoft licensing; runs remote activation code as admin."
    RiskTechnical  = "High: modifies licensing/activation state; remote script execution."
    RiskSource     = "Medium: well-known open project, but fetched live over HTTPS at run time."
    RiskPrivacy    = "Low: no exfiltration identified; activation traffic to Microsoft/remote host."
    RiskLicense    = "High: use outside Microsoft terms may violate license or law."
    RiskPrivileges = "Requires admin; system-wide licensing changes."
    RiskEvidence   = "https://github.com/massgravel/Microsoft-Activation-Scripts"
    RiskReviewed   = "2026-10-16"
}

Register-Module "IAS" "Activate IDM" "Activates Internet Download Manager (IDM)." $Modules["IAS"] "High" @{
    Category       = "Activation"
    Group          = "Other"
    ActionLabel    = "Run"
    ConfirmMessage = "IAS activates IDM outside its license terms.`nThis may violate the IDM license or local law and runs remote code. Continue only where legal and at your own risk."
    RiskReason     = "Bypasses IDM licensing; runs remote activation code."
    RiskTechnical  = "High: modifies licensed app state; remote script execution."
    RiskSource     = "Medium: fetched live from coporton.com/ias at run time."
    RiskPrivacy    = "Low: no exfiltration identified."
    RiskLicense    = "High: use outside IDM terms may violate license or law."
    RiskPrivileges = "Requires admin."
    RiskEvidence   = "https://coporton.com/ias"
    RiskReviewed   = "2026-10-16"
}

# Moved from Software: the Adobe entries are licensing/activation work, not
# standalone applications. AdobeUninstall came along because splitting the two
# Adobe entries across two tabs would leave a one-member "Adobe" group behind in
# Software; see the session report for why that call was made.
Register-Module "AdobeGenP" "Adobe CC & GenP" "Downloads Creative Cloud and GenP activator." $Modules["AdobeGenP"] "High" @{
    RequiresNetwork = $true
    Category        = "Activation"
    Group           = "Adobe"
    ActionLabel     = "Run"
    ConfirmMessage  = "GenP patches Adobe apps outside Adobe licensing.`nThis may violate Adobe terms or local law. Builds come from community mirrors (unverified executables) and patching requires admin. Continue only where legal and at your own risk."
    RiskReason      = "Unverified GenP executable from mirrors; patches Adobe licensing."
    RiskTechnical   = "High: runs an unverified third-party patcher that modifies installed Adobe software."
    RiskSource      = "Low-medium: community mirrors, executable provenance varies per release."
    RiskPrivacy     = "Unknown: closed patcher; no independent audit claimed."
    RiskLicense     = "High: use outside Adobe terms may violate license or law."
    RiskPrivileges  = "Requires admin; modifies installed applications."
    RiskEvidence    = "https://wiki.dbzer0.com/genp-guides/guide/"
    RiskReviewed    = "2026-10-16"
}

Register-Module "AdobeUninstall" "Adobe Uninstall" "Downloads and launches Adobe's official Creative Cloud Cleaner Tool." $Modules["AdobeUninstall"] "Low" @{
    RequiresNetwork = $true
    Category        = "Activation"
    Group           = "Adobe"
    ActionLabel     = "Download"
    Logo            = "App_AdobeUninstall"
    RiskReason      = "Official Adobe removal tool; removes Creative Cloud components."
    RiskTechnical   = "Low: vendor cleaner with predictable removal behavior."
    RiskSource      = "High confidence: Adobe official cleaner host."
    RiskPrivacy     = "None identified."
    RiskLicense     = "None."
    RiskPrivileges  = "Requires admin; removes installed software."
    RiskEvidence    = "https://swupmf.adobe.com/webfeed/CleanerTool/win/AdobeCreativeCloudCleanerTool.exe"
    RiskReviewed    = "2026-10-16"
}

# --- Software: developer tooling ----------------------------------------------
# These were one "Dev Choice (Full)" bundle. Each entry below installs exactly
# one component, using the bundle's own command for that component.

Register-Module "Dev_VCRedist" "Visual C++ Redistributables" "Visual C++ runtimes, 2005-2022, x86 and x64." $Modules["Dev_VCRedist"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "Dev_VCRedist"
}

Register-Module "Dev_DotNet" ".NET" ".NET 10 SDK plus the 8 and 9 runtimes." $Modules["Dev_DotNet"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "Dev_DotNet"
}

Register-Module "Dev_NodeJS" "Node.js" "Node.js LTS runtime." $Modules["Dev_NodeJS"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "Dev_NodeJS"
}

Register-Module "Dev_Python" "Python" "Latest Python 3 release." $Modules["Dev_Python"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "Dev_Python"
}

Register-Module "Dev_Java" "Java" "Eclipse Temurin JDK." $Modules["Dev_Java"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "Dev_Java"
}

Register-Module "Dev_PowerShell" "PowerShell" "PowerShell 7, alongside Windows PowerShell 5.1." $Modules["Dev_PowerShell"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "Dev_PowerShell"
}

Register-Module "Dev_Git" "Git" "Git for Windows." $Modules["Dev_Git"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "Dev_Git"
}

Register-Module "Dev_FFmpeg" "FFmpeg" "Command-line audio and video toolkit." $Modules["Dev_FFmpeg"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "Dev_FFmpeg"
}

Register-Module "Dev_BibataCursor" "Bibata Cursor" "Bibata cursor theme." $Modules["Dev_BibataCursor"] "Low" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "Dev_BibataCursor"
    RiskReason      = "Third-party cursor theme installed via INF; reversible."
    RiskTechnical   = "Low: INF cursor install, well-understood, no services."
    RiskSource      = "Medium-high: official ful1e5/Bibata_Cursor GitHub release asset."
    RiskPrivacy     = "None identified."
    RiskLicense     = "None; open theme release."
    RiskPrivileges  = "Requires admin for INF install."
    RiskEvidence    = "https://github.com/ful1e5/Bibata_Cursor/releases"
    RiskReviewed    = "2026-10-16"
}

Register-Module "App_VisualStudio" "Visual Studio" "Visual Studio Community IDE." $Modules["App_VisualStudio"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "App_VisualStudio"
}

Register-Module "App_AndroidStudio" "Android Studio" "Official Android IDE." $Modules["App_AndroidStudio"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "App_AndroidStudio"
}

Register-Module "App_Godot" "Godot Engine" "Open-source game engine." $Modules["App_Godot"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "App_Godot"
}

Register-Module "App_UnityHub" "Unity Hub" "Manage Unity installs and projects." $Modules["App_UnityHub"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Development"
    ActionLabel     = "Install"
    Logo            = "App_UnityHub"
}

Register-Module "App_NanaZip" "NanaZip" "Open-source file archiver for Windows." $Modules["App_NanaZip"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Productivity"
    ActionLabel     = "Install"
    Logo            = "App_NanaZip"
    Verify          = { Test-WingetPackageInstalled -Id "M2Team.NanaZip" }
}

Register-Module "Dev_NotepadPlusPlus" "Notepad++" "Notepad++ text editor." $Modules["Dev_NotepadPlusPlus"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Productivity"
    ActionLabel     = "Install"
    Logo            = "Dev_NotepadPlusPlus"
}

Register-Module "Dev_FileConverter" "File Converter" "Right-click file format conversion." $Modules["Dev_FileConverter"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Video & Image Tools"
    ActionLabel     = "Install"
    Logo            = "Dev_FileConverter"
}

Register-Module "App_HandBrake" "HandBrake" "Video transcoding tool." $Modules["App_HandBrake"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Video & Image Tools"
    ActionLabel     = "Install"
    Logo            = "App_HandBrake"
}

# DirectX and XNA were part of the old bundle but belong to gaming rather than
# development, so they get their own entry instead of being folded into one of
# the component modules above.
Register-Module "Dev_GameRuntimes" "Game Runtimes" "DirectX and XNA runtimes for older games." $Modules["Dev_GameRuntimes"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Gaming"
    ActionLabel     = "Install"
    Logo            = "Dev_GameRuntimes"
}

Register-Module "App_Steam" "Steam" "PC gaming platform." $Modules["App_Steam"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Gaming"
    ActionLabel     = "Install"
    Logo            = "App_Steam"
}

# --- Software: memory cleaners ------------------------------------------------
# Moved from Utilities: both are standalone tray/window applications with their
# own interface, not system tooling.
Register-Module "MemReduct" "Mem Reduct" "Lightweight memory cleaner." $Modules["MemReduct"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "System & Maintenance"
    ActionLabel     = "Install"
    Logo            = "App_MemReduct"
    Verify          = { Test-WingetPackageInstalled -Id "Henry++.MemReduct" }
}

Register-Module "WinMemoryCleaner" "WinMemoryCleaner" "One-click RAM optimization." $Modules["WinMemoryCleaner"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "System & Maintenance"
    ActionLabel     = "Install"
    Logo            = "App_WinMemoryCleaner"
    Verify          = { Test-WingetPackageInstalled -Id "IgorMundstein.WinMemoryCleaner" }
}

Register-Module "App_BCUninstaller" "BCUninstaller" "Bulk program uninstaller." $Modules["App_BCUninstaller"] "Low" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "System & Maintenance"
    ActionLabel     = "Install"
    Logo            = "App_BCUninstaller"
    RiskReason      = "Install is safe; bulk use can remove much at once, review before use."
    RiskTechnical   = "Low for install; operational care needed when bulk-uninstalling."
    RiskSource      = "High confidence: Klocman Bulk-Crap-Uninstaller open project, Winget package."
    RiskPrivacy     = "None identified."
    RiskLicense     = "None."
    RiskPrivileges  = "Requires admin for install; elevated for removals."
    RiskEvidence    = "https://github.com/Klocman/Bulk-Crap-Uninstaller"
    RiskReviewed    = "2026-10-16"
}

# --- Software: Spotify / Discord ----------------------------------------------
Register-Module "Spicetify" "Spicetify" "Spotify customization and ad-blocking." $Modules["Spicetify"] "Medium" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Media"
    ActionLabel     = "Install"
    Logo            = "App_Spicetify"
    RiskReason      = "Patches Spotify client; scheduled-task install; ToS risk separate."
    RiskTechnical   = "Medium: modifies Spotify install and schedules a user task."
    RiskSource      = "Medium-high: spicetify marketplace install script, fetched at run time."
    RiskPrivacy     = "Low: local theming; no exfiltration identified."
    RiskLicense     = "Medium: ad-blocking/customization may conflict with Spotify terms; not a safety verdict."
    RiskPrivileges  = "Admin to schedule; patch runs as user."
    RiskEvidence    = "https://github.com/spicetify/marketplace"
    RiskReviewed    = "2026-10-16"
}

Register-Module "SpotX" "SpotX" "Spotify ad-block patcher." $Modules["SpotX"] "Medium" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Media"
    # "Run" rather than "Install": this patches an existing Spotify install.
    ActionLabel     = "Patch"
    Logo            = "App_SpotX"
    RiskReason      = "Patches the Spotify client; ad-blocking may conflict with Spotify terms."
    RiskTechnical   = "Medium: modifies the installed Spotify client; reversible via reinstall."
    RiskSource      = "Medium-high: SpotX-Official GitHub project, fetched at run time."
    RiskPrivacy     = "Low: local patch; no exfiltration identified."
    RiskLicense     = "Medium: ad-blocking may conflict with Spotify terms."
    RiskPrivileges  = "Runs as user via scheduled task; admin only to register it."
    RiskEvidence    = "https://github.com/SpotX-Official/SpotX"
    RiskReviewed    = "2026-10-16"
}

Register-Module "App_VLC" "VLC Media Player" "Plays almost any media format." $Modules["App_VLC"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Media"
    ActionLabel     = "Install"
    Logo            = "App_VLC"
}

Register-Module "App_Stremio" "Stremio" "Media discovery and streaming application." $Modules["App_Stremio"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Media"
    ActionLabel     = "Install"
    Logo            = "App_Stremio"
}

Register-Module "Legcord" "Legcord" "Lightweight Discord client." $Modules["Legcord"] "Medium" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Messaging"
    ActionLabel     = "Install"
    Logo            = "App_Legcord"
    RiskReason      = "Unofficial Discord client: handles account tokens; ToS caveats."
    RiskTechnical   = "Medium: third-party client replacing the official app surface."
    RiskSource      = "Medium-high: Legcord GitHub releases; third party, not Discord."
    RiskPrivacy     = "Medium: chat client handles tokens/messages; use official privacy controls."
    RiskLicense     = "Medium: third-party clients may conflict with Discord terms."
    RiskPrivileges  = "Standard install."
    RiskEvidence    = "https://github.com/Legcord/Legcord/releases"
    RiskReviewed    = "2026-10-16"
}

Register-Module "App_Telegram" "Telegram Desktop" "Desktop messaging app." $Modules["App_Telegram"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Messaging"
    ActionLabel     = "Install"
    Logo            = "App_Telegram"
}

Register-Module "App_ZapFast" "ZapFast" "Unofficial WhatsApp client." $Modules["App_ZapFast"] "Medium" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Messaging"
    ActionLabel     = "Install"
    Logo            = "App_ZapFast"
    RiskReason      = "Unofficial client: account-restriction, privacy, and compatibility caveats."
    RiskTechnical   = "Medium: third-party client install; SHA-256 verified against release metadata."
    RiskSource      = "Medium-high: official crmne/zapfast GitHub releases; not WhatsApp/Meta."
    RiskPrivacy     = "Medium: messaging client handles account and contacts; review repo privacy notes."
    RiskLicense     = "Medium: unofficial clients may conflict with WhatsApp terms."
    RiskPrivileges  = "Standard install."
    RiskEvidence    = "https://github.com/crmne/zapfast/releases"
    RiskReviewed    = "2026-10-16"
}

# --- Software: general catalog ------------------------------------------------
# WinGet first (every ID verified with `winget search --id <Id> --exact`
# before it was written here), official download page as fallback. ZapFast has
# no verified WinGet package and installs from its official GitHub releases.
Register-Module "App_Brave" "Brave Browser" "Privacy-focused web browser." $Modules["App_Brave"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "Browsers"
    ActionLabel     = "Install"
    Logo            = "App_Brave"
}

Register-Module "App_LocalSend" "LocalSend" "Local-network file transfer." $Modules["App_LocalSend"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "File Sharing & Downloads"
    ActionLabel     = "Install"
    Logo            = "App_LocalSend"
}

Register-Module "App_QBittorrent" "qBittorrent" "Open-source torrent client." $Modules["App_QBittorrent"] "None" @{
    RequiresNetwork = $true
    Category        = "Software"
    Group           = "File Sharing & Downloads"
    ActionLabel     = "Install"
    Logo            = "App_QBittorrent"
}

# --- Utilities ----------------------------------------------------------------
Register-Module "FixResolution" "Fix Resolution" "Uses CRU to restart graphics driver and fix resolution." $Modules["FixResolution"] "Medium" @{
    RequiresNetwork = $true
    Category        = "Utilities"
    Group           = "Fixes"
    ActionLabel     = "Run"
    RiskReason      = "Restarts graphics driver and applies custom resolution data."
    RiskTechnical   = "Medium: driver restart plus EDID override; display may go blank until revert."
    RiskSource      = "Medium-high: ToastyX CRU via monitortests host; long-standing utility."
    RiskPrivacy     = "None identified."
    RiskLicense     = "None."
    RiskPrivileges  = "Requires admin; display-driver level change."
    RiskReviewed    = "2026-10-16"
}

Register-Module "MSGameBarFix" "ms-gamebar-fix" "Downloads and runs the Game Bar popup fixer from ajw0/ms-gamebar-fix." $Modules["MSGameBarFix"] "Medium" @{
    RequiresNetwork = $true
    Category        = "Utilities"
    Group           = "Fixes"
    ActionLabel     = "Run"
    RiskReason      = "Small-repo script that changes Game Bar behavior."
    RiskTechnical   = "Medium: registry/service-level Game Bar fix; limited verification."
    RiskSource      = "Medium: small ajw0/ms-gamebar-fix repo fetched at run time; review before use."
    RiskPrivacy     = "None identified."
    RiskLicense     = "None."
    RiskPrivileges  = "Requires admin."
    RiskEvidence    = "https://github.com/ajw0/ms-gamebar-fix"
    RiskReviewed    = "2026-10-16"
}

# Moved from Software, together with WinUtil. Winget Upgrade is in the same
# "Windows Tweaks" group and is system maintenance rather than an application, so
# leaving it behind would have split the group and contradicted the split above.
Register-Module "WingetUpgrade" "Winget Upgrade" "Upgrades all installed software via Winget." $Modules["WingetUpgrade"] "Low" @{
    RequiresNetwork = $true
    Category        = "Utilities"
    Group           = "Windows Tweaks"
    ActionLabel     = "Run"
    Verify          = { [bool](Get-Command winget -ErrorAction SilentlyContinue) }
    RiskReason      = "Upgrades everything at once; one bad update can break a workflow."
    RiskTechnical   = "Low-medium: official Winget upgrade path; broad blast radius."
    RiskSource      = "High confidence: Winget itself."
    RiskPrivacy     = "None identified."
    RiskLicense     = "None."
    RiskPrivileges  = "Requires admin."
    RiskReviewed    = "2026-10-16"
}

Register-Module "WinUtil" "WinUtil" "Launches Chris Titus Tech's Windows Utility." $Modules["WinUtil"] "Low" @{
    RequiresNetwork = $true
    Category        = "Utilities"
    Group           = "Windows Tweaks"
    ActionLabel     = "Launch"
    RiskReason      = "Established open utility; individual tweaks vary, review inside WinUtil."
    RiskTechnical   = "Low overall for launching; operations inside can make substantial Windows changes."
    RiskSource      = "High confidence: ChrisTitusTech/winutil public project; script fetched at run time."
    RiskPrivacy     = "Low: local tweaks; review each option before applying."
    RiskLicense     = "None for the launcher itself."
    RiskPrivileges  = "Requires admin; inner operations may change services, registry, features."
    RiskEvidence    = "https://github.com/ChrisTitusTech/winutil"
    RiskReviewed    = "2026-10-16"
}

Register-Module "Winhance" "Winhance" "Launches the Winhance bootstrap script from get.winhance.net." $Modules["Winhance"] "Low" @{
    RequiresNetwork = $true
    Category        = "Utilities"
    Group           = "Windows Tweaks"
    ActionLabel     = "Launch"
    RiskReason      = "Established Windows utility; review options before applying."
    RiskTechnical   = "Low: launcher applies user-selected local tweaks."
    RiskSource      = "Medium-high: get.winhance.net bootstrap, fetched at run time."
    RiskPrivacy     = "Low: local tweaks; review each option."
    RiskLicense     = "None for the launcher itself."
    RiskPrivileges  = "Requires admin for most tweaks."
    RiskEvidence    = "https://get.winhance.net"
    RiskReviewed    = "2026-10-16"
}

# --- About --------------------------------------------------------------------
# Formerly the Links tab: repository, website, and shortcut cards are
# preserved unchanged, joined by the About lead-in panel (description,
# features, Report an Issue, Share Logs) built in 10-support.ps1.
Register-Module "OpenGitHub" "GitHub Repository" "Opens the FreeMixKit GitHub page in your browser." $Modules["OpenGitHub"] "None" @{
    Category    = "About"
    Group       = "Project"
    ActionLabel = "Open"
}

Register-Module "OpenWebsite" "Website" "Opens the author's website in your browser." $Modules["OpenWebsite"] "None" @{
    Category    = "About"
    Group       = "Project"
    ActionLabel = "Open"
}

Register-Module "AddShortcut" "Add Shortcut" "Creates a shortcut for this script on the Desktop." $Modules["AddShortcut"] "None" @{
    Category    = "About"
    Group       = "Desktop"
    ActionLabel = "Create"
}
