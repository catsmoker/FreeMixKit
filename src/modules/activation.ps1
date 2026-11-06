# ==============================================================================
# MODULE: Activation
# ==============================================================================
# NOTE: These modules download and execute third-party activation scripts that
# bypass software licensing. Kept functionally identical to the original script
# as an explicit project decision -- the UI rework does not alter their behaviour.

$Modules["MAS"] = {
    # Primary: standard HTTPS
    try {
        Invoke-RestMethod https://get.activated.win -UseBasicParsing -ErrorAction Stop | Invoke-Expression
        return
    }
    catch { $err1 = $_ }

    # Fallback 1: curl.exe with DoH (Cloudflare 1.1.1.1) to bypass DNS issues
    try {
        $response = curl.exe -s --doh-url https://1.1.1.1/dns-query https://get.activated.win
        if ($LASTEXITCODE -eq 0 -and $response) {
            $response | Out-String | Invoke-Expression
            return
        }
    }
    catch { $err2 = $_ }

    # Fallback 2: Azure DevOps API - download MAS_AIO.cmd and run it
    try {
        $url = 'https://dev.azure.com/massgrave/Microsoft-Activation-Scripts/_apis/git/repositories/Microsoft-Activation-Scripts/items?path=/MAS/All-In-One-Version-KL/MAS_AIO.cmd&download=true'
        $tempCmd = "$env:TEMP\MAS_AIO_$([Guid]::NewGuid()).cmd"
        Invoke-RestMethod $url -OutFile $tempCmd -UseBasicParsing -ErrorAction Stop
        if (Test-Path $tempCmd) {
            Start-Process -FilePath $tempCmd -ArgumentList '-el' -Wait -Verb RunAs
            Remove-Item -Path $tempCmd -Force -ErrorAction SilentlyContinue
            return
        }
    }
    catch { $err3 = $_ }

    Write-Error "MAS activation failed. Primary: $($err1.Exception.Message); Fallback1: $($err2.Exception.Message); Fallback2: $($err3.Exception.Message)"
    throw "All MAS download methods failed"
}
$Modules["IAS"] = { Invoke-RestMethod https://raw.githubusercontent.com/hanmaoye/IDM-Activation-Script/main/IAS.ps1 | Invoke-Expression }
