# ==============================================================================
# MODULE: Cleanup & repair
# ==============================================================================

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
        Invoke-WebRequest "https://go.microsoft.com/fwlink/?LinkID=212732" -OutFile "$DataDownloads\MRT.exe"
        $mrt = "$DataDownloads\MRT.exe"
    }
    Start-Process $mrt -Wait

    $kvrt = Join-Path $DataDownloads "kvrt.exe"
    Invoke-WebRequest "https://devbuilds.s.kaspersky-labs.com/kvrt/latest/full/kvrt.exe" -OutFile $kvrt
    Start-Process $kvrt -Wait
}

$Modules["SystemReport"] = {
    # Full hardware and Windows report via the shared collection service.
    # Sensitive identifiers stay withheld. Desktop path is the long-standing
    # behavior: Run generates the file and opens it, nothing renders on screen.
    Write-Log "Collecting full system report..." "Info"
    $report = Get-FmkSystemReportData
    $text = Format-FmkSysTextReport -Report $report -IncludeSensitive:$false -ForExport:$true
    $f = "$env:USERPROFILE\Desktop\SysReport.txt"
    Write-FmkSysReportFile -Path $f -Text $text
    Write-Log "System report saved (sensitive identifiers withheld): $f" "Success"
    Invoke-Item $f
}
