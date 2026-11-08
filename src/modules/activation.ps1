# ==============================================================================
# MODULE: Activation
# ==============================================================================
# NOTE: These modules download and execute third-party activation scripts that
# bypass software licensing. Kept functionally identical to the original script
# as an explicit project decision -- the UI rework does not alter their behaviour.

$Modules["MAS"] = {
    # Three download paths, tried in order. Each one is caught so a failure
    # moves on to the next instead of aborting the whole action.
    $err1 = $null; $err2 = $null; $err3 = $null

    # Primary: standard HTTPS
    Write-Log "MAS: Trying primary (https://get.activated.win)..." "Info"
    try {
        Invoke-RestMethod https://get.activated.win -UseBasicParsing -ErrorAction Stop | Invoke-Expression
        Write-Log "MAS: Primary succeeded." "Success"
        return
    }
    catch { $err1 = $_; Write-Log "MAS: Primary failed - $($err1.Exception.Message)" "Warn" }

    # Fallback 1: curl.exe with DoH (Cloudflare 1.1.1.1) to bypass DNS issues
    Write-Log "MAS: Trying fallback 1 (curl.exe with Cloudflare DoH)..." "Info"
    try {
        $response = curl.exe -s --doh-url https://1.1.1.1/dns-query https://get.activated.win
        if ($LASTEXITCODE -eq 0 -and $response) {
            $response | Out-String | Invoke-Expression
            Write-Log "MAS: Fallback 1 succeeded." "Success"
            return
        }
        throw "curl.exe exited with code $LASTEXITCODE and returned no content."
    }
    catch { $err2 = $_; Write-Log "MAS: Fallback 1 failed - $($err2.Exception.Message)" "Warn" }

    # Fallback 2: Azure DevOps API - download MAS_AIO.cmd and run it
    Write-Log "MAS: Trying fallback 2 (Azure DevOps - MAS_AIO.cmd)..." "Info"
    try {
        $url = 'https://dev.azure.com/massgrave/Microsoft-Activation-Scripts/_apis/git/repositories/Microsoft-Activation-Scripts/items?path=/MAS/All-In-One-Version-KL/MAS_AIO.cmd&download=true'
        $tempCmd = "$env:TEMP\MAS_AIO_$([Guid]::NewGuid()).cmd"
        Invoke-RestMethod $url -OutFile $tempCmd -UseBasicParsing -ErrorAction Stop
        if (-not (Test-Path -LiteralPath $tempCmd)) { throw "MAS_AIO.cmd was not downloaded." }
        Start-Process -FilePath $tempCmd -ArgumentList '-el' -Wait -Verb RunAs
        Remove-Item -Path $tempCmd -Force -ErrorAction SilentlyContinue
        Write-Log "MAS: Fallback 2 succeeded." "Success"
        return
    }
    catch { $err3 = $_; Write-Log "MAS: Fallback 2 failed - $($err3.Exception.Message)" "Warn" }

    $detail = (@(
            "primary: $(if ($err1) { $err1.Exception.Message } else { 'failed' })"
            "fallback 1: $(if ($err2) { $err2.Exception.Message } else { 'failed' })"
            "fallback 2: $(if ($err3) { $err3.Exception.Message } else { 'failed' })"
        ) -join '; ')
    throw "MAS activation failed. All download methods failed ($detail)."
}
$Modules["IAS"] = { Invoke-RestMethod https://raw.githubusercontent.com/hanmaoye/IDM-Activation-Script/main/IAS.ps1 | Invoke-Expression }
