# ==============================================================================
# CORE HELPERS
# ==============================================================================

# Write-Log  lives in 03-ui-log.ps1, Confirm-Action in 05-dialogs.ps1.
# Both are UI-aware and are defined later in the build so they win over any
# earlier definition. Everything below is UI-agnostic.

function Ensure-WingetInstalled {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Log "Winget not found. Installing Winget..." "Warn"
        $bundlePath = Join-Path $DataTemp "winget.msixbundle"
        Invoke-WebRequest -Uri "https://aka.ms/getwinget" -OutFile $bundlePath
        Add-AppxPackage -Path $bundlePath
    }

    # Keep winget and its sources fresh. Their real output is streamed line by
    # line instead of discarded: these are quick, already in memory (no
    # process plumbing needed), and silent runs hide verification failures.
    try {
        Write-Log "Refreshing winget sources..." "Info"
        $refreshLines = @(winget source update --disable-interactivity 2>&1 | ForEach-Object { "$_" })
        foreach ($line in $refreshLines) {
            if ($line -and $line.Trim()) { Write-Log $line.TrimEnd() "Info" }
        }
    }
    catch {
        Write-Log "Winget source refresh failed: $($_.Exception.Message)" "Warn"
    }

    try {
        Write-Log "Checking winget client updates..." "Info"
        $upgradeLines = @(winget upgrade --id Microsoft.DesktopAppInstaller --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity 2>&1 | ForEach-Object { "$_" })
        foreach ($line in $upgradeLines) {
            if ($line -and $line.Trim()) { Write-Log $line.TrimEnd() "Info" }
        }
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

function Get-InstalledExecutablePath {
    param(
        [Parameter(Mandatory = $true)][string[]]$ExecutableNames,
        [Parameter()][string[]]$CandidatePaths = @()
    )

    foreach ($path in $CandidatePaths | Where-Object { $_ }) {
        if (Test-Path $path) { return $path }
    }

    foreach ($name in $ExecutableNames | Where-Object { $_ }) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($command -and $command.Source -and (Test-Path $command.Source)) {
            return $command.Source
        }
    }

    $searchRoots = @()
    if ($env:LOCALAPPDATA) {
        $searchRoots += Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages"
        $searchRoots += Join-Path $env:LOCALAPPDATA "Programs"
    }
    if ($env:ProgramFiles) {
        $searchRoots += Join-Path $env:ProgramFiles "WinGet\Packages"
        $searchRoots += $env:ProgramFiles
    }
    if (${env:ProgramFiles(x86)}) {
        $searchRoots += ${env:ProgramFiles(x86)}
    }
    $searchRoots = $searchRoots | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

    foreach ($root in $searchRoots) {
        foreach ($name in $ExecutableNames | Where-Object { $_ }) {
            $match = Get-ChildItem -Path $root -Filter $name -File -Recurse -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if ($match) { return $match.FullName }
        }
    }

    return $null
}

function Install-And-LaunchWingetPackage {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string[]]$ExecutableNames,
        [Parameter()][string[]]$CandidatePaths = @()
    )

    Install-WingetPackage -Id $Id

    $exePath = Get-InstalledExecutablePath -ExecutableNames $ExecutableNames -CandidatePaths $CandidatePaths
    if (-not $exePath) {
        throw "Installed package '$Id', but could not locate executable: $($ExecutableNames -join ', ')"
    }

    Write-Log "Launching $exePath" "Info"
    Start-Process -FilePath $exePath
}

function Test-NetworkConnectivity {
    try {
        return [bool](Test-Connection -ComputerName "1.1.1.1" -Count 1 -Quiet -ErrorAction Stop)
    }
    catch {
        return $false
    }
}

function Get-WindowsBuildNumber {
    try {
        return [int](Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "CurrentBuild")
    }
    catch {
        return 0
    }
}

function Test-IsWindows11 {
    if ($SysInfo.OS -match "Windows 11") {
        return $true
    }

    return (Get-WindowsBuildNumber) -ge 22000
}

function Get-SpotifyExecutablePath {
    $candidates = @(
        "$env:APPDATA\Spotify\Spotify.exe",
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\Spotify.exe"
    )

    $msixRoot = Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps"
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

# Lines streamed per poll tick at most: a pathological burst is paced across
# ticks instead of freezing the UI behind thousands of synchronous log writes.
$FmkStreamPollLines = 500

function Invoke-ExternalCommand {
    <#
        Runs an external tool and STREAMS its real output into the log while it
        runs -- stdout as Info, stderr as Warn -- instead of replaying it after
        exit. The console host is hidden for the life of the GUI, so without
        this the output of winget/sfc/DISM would go to a window nobody can see.

        Mechanism, and why not .NET output events: event callbacks would run on
        threadpool threads inside PowerShell scriptblocks bound to the very
        runspace that is blocked waiting -- a deadlock. Instead this thread
        polls both pipes itself with BeginRead plus a 200ms wait handle (pure
        .NET, no scriptblock callbacks, no shared locks), and each new segment
        is logged immediately, so the Output panel updates during long
        installs. Per-stream order is exact; across the two independent
        streams the order is arrival order, which the OS does not guarantee --
        and is not claimed to be reconstructable.

        Decoding is UTF-8 unless the stream proves otherwise: a BOM or a
        NUL-interleaved prefix (how UTF-16LE console output such as sfc.exe
        arrives) switches that stream to a UTF-16 decoder for the rest of the
        run. Anything undecodable degrades to replacement characters rather
        than throwing mid-stream. Carriage-return progress rewrites are split
        into segments so they cannot arrive as one enormous line. Leading
        indentation is preserved; only trailing whitespace is trimmed and empty
        segments are skipped. Success or failure comes from the exit code
        alone -- stderr text is never itself a failure.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter()][string[]]$Arguments = @()
    )

    Write-Log "> $FilePath $($Arguments -join ' ')" "Info"

    # Quote exactly like an argv array would: Start-Process used to take the
    # array directly, and a naive join would corrupt paths containing spaces.
    $escapedArgs = foreach ($a in @($Arguments)) {
        if ($null -eq $a -or "$a" -eq '') { continue }
        if ("$a" -match '[\s"]') { '"' + ("$a" -replace '"', '\"') + '"' }
        else { "$a" }
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = ($escapedArgs -join ' ')
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    $proc = $null
    try {
        $proc = [System.Diagnostics.Process]::Start($psi)
    }
    catch {
        throw "Could not start '$FilePath': $($_.Exception.Message)"
    }

    # One state block per pipe. BaseStream is read directly -- never the
    # StreamReader wrappers, which would fight over the same pipe. The text
    # decoder is chosen per stream by Read-FmkStreamBytes (UTF-8 unless the
    # bytes prove UTF-16) and then kept for the whole run, so a multibyte or
    # surrogate-split character across two reads still decodes correctly.
    $streams = @(
        @{ Pipe = $proc.StandardOutput.BaseStream; Tag = 'OUT'; LogType = 'Info';
           Enc = $null; Decoder = $null; Lead = (New-Object System.Collections.Generic.List[byte]);
           Buffer = (New-Object byte[] 4096); PendingAsync = $null;
           Fragment = ''; Eof = $false },
        @{ Pipe = $proc.StandardError.BaseStream; Tag = 'ERR'; LogType = 'Warn';
           Enc = $null; Decoder = $null; Lead = (New-Object System.Collections.Generic.List[byte]);
           Buffer = (New-Object byte[] 4096); PendingAsync = $null;
           Fragment = ''; Eof = $false }
    )

    try {
        $openStreams = @($streams | Where-Object { -not $_.Eof })
        while ((-not $proc.HasExited) -or ($openStreams.Count -gt 0)) {
            foreach ($st in $streams) {
                if ($st.Eof) { continue }
                try {
                    if ($null -eq $st.PendingAsync) {
                        $st.PendingAsync = $st.Pipe.BeginRead($st.Buffer, 0, $st.Buffer.Length, $null, $null)
                    }
                    if ($st.PendingAsync.AsyncWaitHandle.WaitOne(200)) {
                        $got = $st.Pipe.EndRead($st.PendingAsync)
                        $st.PendingAsync = $null
                        if ($got -eq 0) { $st.Eof = $true; continue }
                        $text = Read-FmkStreamBytes -State $st -Buffer $st.Buffer -Count $got
                        if ($text -ne '') {
                            Write-FmkStreamText -State $st -Text $text -Final:$false
                        }
                    }
                }
                catch {
                    # A broken pipe ends that stream, not the task. The exit
                    # code below still decides success or failure.
                    $st.Eof = $true
                    $st.PendingAsync = $null
                }
            }
            $openStreams = @($streams | Where-Object { -not $_.Eof })
        }
        foreach ($st in $streams) {
            $tail = Read-FmkStreamBytes -State $st -Buffer (New-Object byte[] 0) -Count 0 -Final
            Write-FmkStreamText -State $st -Text $tail -Final:$true
        }
    }
    catch {
        # Capture must never stick the task: wait out our own child (waiting
        # is correct -- never kill) and let the exit code decide.
        Write-Log "Stream capture degraded, waiting for exit: $($_.Exception.Message)" "Warn"
        try { $proc.WaitForExit() } catch { }
    }

    $exitCode = 0
    try { $exitCode = $proc.ExitCode } catch { $exitCode = 0 }
    try { $proc.Dispose() } catch { }

    if ($exitCode -ne 0) {
        throw "Command failed (exit $exitCode): $FilePath $($Arguments -join ' ')"
    }

    return [pscustomobject]@{
        FilePath  = $FilePath
        Arguments = ($Arguments -join " ")
        ExitCode  = $exitCode
    }
}

function Test-FmkStreamEncoding {
    <#
        Decides a pipe's text encoding from its leading bytes. A BOM decides
        immediately (UTF-16LE, or UTF-8 with its BOM skipped). Otherwise the
        decision waits for at least 4 bytes: NUL-interleaved ASCII-range text
        means UTF-16LE, anything else means UTF-8. Short prefixes stay
        undecided ($null) until more arrive or the stream ends (-Final falls
        back to UTF-8). Pure function, unit tested.
    #>
    param([byte[]]$Lead = @(), [switch]$Final)
    $Lead = @($Lead)
    if ($Lead.Count -ge 2 -and $Lead[0] -eq 0xFF -and $Lead[1] -eq 0xFE) {
        return @{ Enc = 'utf16'; Skip = 2 }
    }
    if ($Lead.Count -ge 3 -and $Lead[0] -eq 0xEF -and $Lead[1] -eq 0xBB -and $Lead[2] -eq 0xBF) {
        return @{ Enc = 'utf8'; Skip = 3 }
    }
    if ($Lead.Count -ge 4) {
        $pairs = [Math]::Floor($Lead.Count / 2)
        $zeroOdd = 0
        $nonZeroEven = 0
        for ($i = 0; $i -lt $pairs; $i++) {
            if ($Lead[$i * 2 + 1] -eq 0) { $zeroOdd++ }
            if ($Lead[$i * 2] -ne 0) { $nonZeroEven++ }
        }
        if ($pairs -ge 2 -and $zeroOdd -eq $pairs -and $nonZeroEven -gt 0) {
            return @{ Enc = 'utf16'; Skip = 0 }
        }
        return @{ Enc = 'utf8'; Skip = 0 }
    }
    if ($Final) { return @{ Enc = 'utf8'; Skip = 0 } }
    return $null
}

function Convert-FmkBytesToText {
    <#
        Decodes bytes with a stateful decoder; never throws. Empty input is a
        legal value (end of stream), not a binding error: $Bytes is therefore
        deliberately NOT a mandatory [byte[]], because PowerShell refuses to
        bind an empty array to a mandatory typed parameter.
    #>
    param($Decoder, $Bytes, [switch]$Flush)
    try {
        $arr = @()
        if ($null -ne $Bytes) { $arr = @($Bytes) }
        $src = New-Object byte[] 0
        if ($arr.Count -gt 0) { $src = [byte[]]$arr }
        if ($src.Length -eq 0 -and -not $Flush) { return '' }
        $chars = New-Object char[] ($src.Length + 1)
        $used = $Decoder.GetChars($src, 0, $src.Length, $chars, 0, [bool]$Flush)
        if ($used -gt 0) { return (-join $chars[0..($used - 1)]) }
        return ''
    }
    catch { return '' }
}

function Read-FmkStreamBytes {
    <#
        Decodes one raw pipe read into text using the stream's sticky decoder,
        choosing it on first evidence via Test-FmkStreamEncoding. Split
        multibyte characters survive because decoder state persists across
        reads; -Final flushes a dangling tail. Returns '' when there is
        nothing decodable yet.
    #>
    param(
        [Parameter(Mandatory)]$State,
        [byte[]]$Buffer,
        [int]$Count,
        [switch]$Final
    )
    if (-not $State.Enc) {
        # Accumulate until decidable. Test-FmkStreamEncoding always decides at
        # 4 bytes, so this buffer stays tiny; nothing is dropped meanwhile.
        for ($i = 0; $i -lt $Count; $i++) { [void]$State.Lead.Add($Buffer[$i]) }
        $leadBytes = @($State.Lead.ToArray())
        $dec = Test-FmkStreamEncoding -Lead $leadBytes -Final:$Final
        if (-not $dec) { return '' }
        $State.Enc = [string]$dec.Enc
        if ($State.Enc -eq 'utf16') {
            $State.Decoder = [System.Text.Encoding]::Unicode.GetDecoder()
        }
        else {
            $State.Decoder = [System.Text.Encoding]::UTF8.GetDecoder()
        }
        $skip = [int]$dec.Skip
        $payload = @()
        if ($skip -lt $leadBytes.Length) { $payload = @($leadBytes[$skip..($leadBytes.Length - 1)]) }
        $State.Lead = (New-Object System.Collections.Generic.List[byte])
        return (Convert-FmkBytesToText -Decoder $State.Decoder -Bytes $payload -Flush:$Final)
    }
    if ($Count -le 0) {
        if ($Final) { return (Convert-FmkBytesToText -Decoder $State.Decoder -Flush:$true) }
        return ''
    }
    return (Convert-FmkBytesToText -Decoder $State.Decoder -Bytes @($Buffer[0..($Count - 1)]) -Flush:$Final)
}

function Get-FmkObservedPercent {
    <#
        Extracts a genuine percentage from real installer output text, e.g.
        "Downloading 42%" or "Progress: 42 %". Returns the LAST percentage on
        the line (later rewrites supersede earlier ones), clamped to 0..100,
        or $null when the line carries no percentage. Pure function: no
        estimation, no timers, no line counting.
    #>
    param([string]$Text)

    if ([string]::IsNullOrEmpty($Text)) { return $null }
    $hits = @([regex]::Matches($Text, '(\d{1,3})\s*%') | ForEach-Object { [int]$_.Groups[1].Value })
    if ($hits.Count -eq 0) { return $null }
    $pct = $hits[$hits.Count - 1]
    if ($pct -lt 0) { $pct = 0 }
    if ($pct -gt 100) { $pct = 100 }
    return $pct
}

function Write-FmkStreamText {
    <#
        Logs newly arrived text from one pipe. Complete segments (split on CR
        and LF, so progress rewrites cannot fuse into one giant line) stream
        immediately; a trailing fragment without a terminator is held back
        until more arrives or the process exits (-Final flushes it).
    #>
    param(
        [Parameter(Mandatory)]$State,
        [string]$Text = '',
        [switch]$Final
    )

    $text = [string]$State.Fragment + $Text
    $State.Fragment = ''
    if ($text -eq '') { return }

    $terminated = $text.EndsWith("`r") -or $text.EndsWith("`n")
    $parts = $text -split "`r`n|`n|`r"
    $complete = $parts
    if (-not $terminated -and -not $Final) {
        $State.Fragment = $parts[$parts.Count - 1]
        if ($parts.Count -gt 1) { $complete = $parts[0..($parts.Count - 2)] }
        else { $complete = @() }
    }

    $shown = 0
    for ($i = 0; $i -lt $complete.Count; $i++) {
        $segment = $complete[$i]
        if ($shown -ge $FmkStreamPollLines) {
            # Burst protection: hold THIS segment and everything after it for
            # the next tick, in order, instead of flooding the UI thread
            # behind one poll. Nothing is dropped.
            $rest = @($complete[$i..($complete.Count - 1)])
            $State.Fragment = (($rest -join "`n") + "`n" + $State.Fragment)
            break
        }
        if ($segment -eq '' -or $segment.Trim() -eq '') { continue }
        Write-Log $segment.TrimEnd() $State.LogType
        $shown++
        # A percentage observed in genuine output feeds the shared bar via the
        # queue/total math (or directly for lone tasks). Lines without one
        # change nothing -- the bar simply stays where measured evidence put it.
        $observed = Get-FmkObservedPercent -Text $segment
        if ($null -ne $observed) { Update-FmkObservedProgress -ItemPercent $observed }
    }
}
