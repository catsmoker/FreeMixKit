# ==============================================================================
# MODULES: Developer tooling
# ==============================================================================
# One module per component. These used to be a single "Dev Choice (Full)" bundle
# that installed all of them at once; each module below is that bundle's own
# command for one component, unchanged, so a user can install just the pieces
# they want. Package IDs, download URLs, and the registry replacement are the
# originals.

function Install-FmkWingetPackages {
    <#
        Winget installs with the per-package error handling and logging the old
        bundle used. Kept as one helper rather than repeated in a dozen modules so
        the failure behaviour cannot drift between them: a package that fails is
        logged and the remaining ones still run.

        Note this deliberately does not swallow the exit code. Invoke-ModuleAction
        turns a non-zero $LASTEXITCODE into a Failed result, which is the existing
        behaviour and what the results CSV has always recorded.
    #>
    param([Parameter(Mandatory)][string[]]$Ids)

    foreach ($id in $Ids) {
        try {
            Install-WingetPackage -Id $id
        }
        catch {
            Write-Log "Failed package '$id': $($_.Exception.Message)" "Error"
        }
    }
}

# ------------------------------------------------------------------------------
# Visual C++ Redistributables
# ------------------------------------------------------------------------------
$Modules["Dev_VCRedist"] = {
    Write-Log "Installing Visual C++ Redistributables (x86/x64)..." "Info"

    Ensure-WingetInstalled

    # VC++ runtime coverage (x86 + x64):
    # Latest v14 (VS 2017-2026), VS 2015 (VC++ 14.0), 2013 (12.0), 2012 (11.0), 2010 (10.0), 2008 (9.0), 2005 (8.0)
    $vcRedists = @(
        @{ Version = "Latest v14 (VS 2017-2026)"; Arch = "x86"; Id = "Microsoft.VCRedist.2015+.x86" },
        @{ Version = "Latest v14 (VS 2017-2026)"; Arch = "x64"; Id = "Microsoft.VCRedist.2015+.x64" },
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

    # Sort -Unique because several entries above deliberately share an ID: the
    # "Latest v14" row and the "Visual Studio 2015" row resolve to the same
    # package, which winget only needs to be asked for once.
    Install-FmkWingetPackages -Ids @($vcRedists | Where-Object { $_ -and $_.Id } |
            ForEach-Object { $_.Id } | Sort-Object -Unique)

    Write-Log "Visual C++ Redistributables done." "Success"
}

# ------------------------------------------------------------------------------
# .NET
# ------------------------------------------------------------------------------
$Modules["Dev_DotNet"] = {
    Write-Log "Installing .NET..." "Info"

    Ensure-WingetInstalled
    Install-FmkWingetPackages -Ids @("Microsoft.DotNet.SDK.10", "Microsoft.DotNet.Runtime.9", "Microsoft.DotNet.Runtime.8")

    Write-Log "Enabling .NET Framework 3.5 (Windows Feature)..." "Info"
    try {
        Start-Process "dism.exe" -ArgumentList "/Online /Enable-Feature /FeatureName:NetFx3 /All /NoRestart" -Wait -NoNewWindow
    }
    catch { Write-Log ".NET 3.5 enable failed: $($_.Exception.Message)" "Warn" }

    Write-Log ".NET done." "Success"
}

# ------------------------------------------------------------------------------
# Node.js
# ------------------------------------------------------------------------------
$Modules["Dev_NodeJS"] = {
    Write-Log "Installing Node.js..." "Info"
    Ensure-WingetInstalled
    Install-FmkWingetPackages -Ids @("OpenJS.NodeJS.LTS")
    Write-Log "Node.js done." "Success"
}

# ------------------------------------------------------------------------------
# Python
# ------------------------------------------------------------------------------
$Modules["Dev_Python"] = {
    Write-Log "Installing Python..." "Info"
    Ensure-WingetInstalled
    Install-FmkWingetPackages -Ids @("Python.Python.3.14")
    Write-Log "Python done." "Success"
}

# ------------------------------------------------------------------------------
# Java
# ------------------------------------------------------------------------------
$Modules["Dev_Java"] = {
    Write-Log "Installing Java (Eclipse Temurin JDK)..." "Info"
    Ensure-WingetInstalled
    Install-FmkWingetPackages -Ids @("EclipseAdoptium.Temurin.25.JDK")
    Write-Log "Java done." "Success"
}

# ------------------------------------------------------------------------------
# PowerShell
# ------------------------------------------------------------------------------
$Modules["Dev_PowerShell"] = {
    Write-Log "Installing PowerShell..." "Info"
    Ensure-WingetInstalled
    # Microsoft.PowerShell is PowerShell 7 side by side with Windows PowerShell 5.1.
    Install-FmkWingetPackages -Ids @("Microsoft.PowerShell")
    Write-Log "PowerShell done." "Success"
}

# ------------------------------------------------------------------------------
# Git
# ------------------------------------------------------------------------------
$Modules["Dev_Git"] = {
    Write-Log "Installing Git..." "Info"
    Ensure-WingetInstalled
    Install-FmkWingetPackages -Ids @("Git.Git")
    Write-Log "Git done." "Success"
}

# ------------------------------------------------------------------------------
# FFmpeg
# ------------------------------------------------------------------------------
$Modules["Dev_FFmpeg"] = {
    Write-Log "Installing FFmpeg..." "Info"
    Ensure-WingetInstalled
    Install-FmkWingetPackages -Ids @("Gyan.FFmpeg")
    Write-Log "FFmpeg done." "Success"
}

# ------------------------------------------------------------------------------
# NanaZip
# ------------------------------------------------------------------------------
$Modules["App_NanaZip"] = {
    # Open-source file archiver for the modern Windows experience, forked from
    # 7-Zip (M2Team/NanaZip, MIT). Catalog replacement for the old 7-Zip entry:
    # installs NanaZip only, never touches an existing 7-Zip installation.
    Write-Log "Installing NanaZip..." "Info"
    Ensure-WingetInstalled
    Install-FmkWingetPackages -Ids @("M2Team.NanaZip")
    Write-Log "NanaZip done." "Success"
}

# ------------------------------------------------------------------------------
# Notepad++
# ------------------------------------------------------------------------------
$Modules["Dev_NotepadPlusPlus"] = {
    Write-Log "Installing Notepad++..." "Info"
    Ensure-WingetInstalled
    Install-FmkWingetPackages -Ids @("Notepad++.Notepad++")

    # Notepad Fix -- file associations and the "Open with Notepad++" shell entry.
    Write-Log "Replacing Windows Notepad with Notepad++..." "Info"
    $nppPath = "C:\Program Files\Notepad++\notepad++.exe"
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

    Write-Log "Notepad++ done." "Success"
}

# ------------------------------------------------------------------------------
# File Converter
# ------------------------------------------------------------------------------
$Modules["Dev_FileConverter"] = {
    Write-Log "Installing File Converter..." "Info"
    Ensure-WingetInstalled
    Install-FmkWingetPackages -Ids @("AdrienAllard.FileConverter")
    Write-Log "File Converter done." "Success"
}

# ------------------------------------------------------------------------------
# Bibata Cursor
# ------------------------------------------------------------------------------
$Modules["Dev_BibataCursor"] = {
    Write-Log "Installing Bibata Cursor..." "Info"
    try {
        $zip = "$DataDownloads\Bibata.zip"; $dest = "$DataDownloads\Bibata"
        Invoke-WebRequest "https://github.com/ful1e5/Bibata_Cursor/releases/download/v2.0.7/Bibata-Modern-Classic-Windows.zip" -OutFile $zip
        Expand-Archive $zip -Dest $dest -Force
        $regularInf = Get-ChildItem "$dest" -Recurse -Filter "*.inf" | Where-Object { $_.FullName -match "Regular" } | Select-Object -First 1
        if (-not $regularInf) { $regularInf = Get-ChildItem "$dest" -Recurse -Filter "*.inf" | Select-Object -First 1 }
        if ($regularInf) { Start-Process "RUNDLL32.EXE" -Arg "SETUPAPI.DLL,InstallHinfSection DefaultInstall 128 $($regularInf.FullName)" -Wait }
    }
    catch { Write-Log "Cursor Failed" "Error" }

    Write-Log "Bibata Cursor done." "Success"
}

# ------------------------------------------------------------------------------
# Game runtimes
# ------------------------------------------------------------------------------
$Modules["Dev_GameRuntimes"] = {
    # DirectX and XNA were part of the old bundle but are not development tools,
    # so they get their own entry instead of being folded into one of the
    # component modules.
    Write-Log "Installing game runtimes (DirectX, XNA)..." "Info"
    Ensure-WingetInstalled

    foreach ($gid in @("Microsoft.DirectX", "Microsoft.XNARedist")) {
        try {
            Write-Log "Installing game runtime: $gid" "Info"
            Install-WingetPackage -Id $gid
        }
        catch {
            Write-Log "Failed game runtime '$gid': $($_.Exception.Message)" "Error"
        }
    }

    Write-Log "Game runtimes done." "Success"
}
