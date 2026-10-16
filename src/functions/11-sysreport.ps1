# ==============================================================================
# SYSTEM REPORT -- data collection service
# ==============================================================================
# One shared service behind the System Report module action, which writes the
# full report to a text file. Built-in Windows interfaces only (CIM/WMI,
# registry, inbox cmdlets and commands): no installs, no drivers, no services,
# no configuration changes, and no separate PowerShell process per field --
# one collection pass gathers everything, section by section. The application
# UI stays a button workflow: Run generates the file, nothing renders the
# report on screen.
#
# Robustness contract: every data source is optional. A failed query, a
# missing property, a permission denial, or unsupported hardware produces an
# 'Unavailable' / 'Not supported' / 'Access denied' value plus a section
# warning -- never an exception, and never a hole that stops the rest of the
# report. Missing information is never rendered as zero.
#
# Privacy contract: product keys are never collected. Hardware serial numbers,
# MAC addresses, and IP addresses are collected with Sensitivity='Private'
# and hidden unless an export explicitly includes them. Computer and user
# names are Sensitivity='Identifying': shown on screen (the label says so)
# but withheld from exports unless included.
#
# ASCII only in this file: the build ships UTF-8 without a BOM, and non-ASCII
# glyphs depend on host-codepage parsing and font coverage agreeing.

$FmkSysMissing = 'Unavailable'

function Read-FmkCim {
    <#
        One guarded CIM query. Always returns an array (possibly empty) or
        $null when the class, namespace, provider, or permissions are
        unavailable. The -NoEnumerate is load-bearing: a bare return would
        unroll a single-instance result into a bare object, and every
        `$x.Count` downstream would then throw under strict mode.
    #>
    param(
        [Parameter(Mandatory)][string]$Class,
        [string]$Namespace = 'root/cimv2',
        [string]$Filter = ''
    )
    try {
        if ($Filter -ne '') {
            $found = @(Get-CimInstance -ClassName $Class -Namespace $Namespace -Filter $Filter -ErrorAction Stop)
        }
        else {
            $found = @(Get-CimInstance -ClassName $Class -Namespace $Namespace -ErrorAction Stop)
        }
        Write-Output $found -NoEnumerate
        return
    }
    catch { return $null }
}

function Get-FmkSysField {
    <#
        First present, non-blank property value, else the Missing fallback.
        Numeric zero and $false are real values and pass through untouched.
        Missing is deliberately untyped so callers can pass $null or @() as
        the "nothing there" sentinel for their own $null checks.
    #>
    param($Object, [Parameter(Mandatory)][string[]]$Names, $Missing = 'Unavailable')
    try {
        if ($null -eq $Object) { return $Missing }
        foreach ($n in $Names) {
            $v = $Object.$n
            if ($null -ne $v -and "$v" -ne '') { return $v }
        }
    }
    catch { }
    return $Missing
}

function New-FmkSysSection {
    param([Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)][string]$Title)
    return @{
        Key      = $Key
        Title    = $Title
        Rows     = (New-Object System.Collections.ArrayList)
        Warnings = (New-Object System.Collections.ArrayList)
    }
}

function Add-FmkSysRow {
    <#
        Appends a Label/Value row. $null and blank become 'Unavailable';
        numeric zero and $false are kept as real values.
    #>
    param($Section, [Parameter(Mandatory)][string]$Label, $Value, [string]$Sensitivity = 'None')
    $text = 'Unavailable'
    try {
        if ($null -ne $Value -and "$Value" -ne '') { $text = [string]$Value }
    }
    catch { }
    [void]$Section.Rows.Add(@{ Label = $Label; Value = $text; Sensitivity = $Sensitivity })
}

function Add-FmkSysWarning {
    param($Section, [Parameter(Mandatory)][string]$Text)
    [void]$Section.Warnings.Add($Text)
}

function Find-FmkSysRow {
    <# Row Value by label from a finished section, or '' when absent. #>
    param($Section, [Parameter(Mandatory)][string]$Label)
    try {
        foreach ($row in @($Section.Rows)) {
            if ($row.Label -eq $Label) { return [string]$row.Value }
        }
    }
    catch { }
    return ''
}

function Format-FmkBytes {
    param($Bytes)
    try {
        $b = [double]$Bytes
        if ($b -lt 0) { return 'Unavailable' }
        if ($b -ge 1GB) { return ('{0:N1} GB' -f ($b / 1GB)) }
        if ($b -ge 1MB) { return ('{0:N1} MB' -f ($b / 1MB)) }
        if ($b -ge 1KB) { return ('{0:N1} KB' -f ($b / 1KB)) }
        return ('{0:N0} B' -f $b)
    }
    catch { return 'Unavailable' }
}

function Format-FmkGigahertz {
    <# Advertised/measured megahertz to a GHz string; 0 means not reported. #>
    param($Megahertz)
    try {
        $m = [double]$Megahertz
        if ($m -le 0) { return 'Unavailable' }
        return ('{0:N2} GHz' -f ($m / 1000))
    }
    catch { return 'Unavailable' }
}

function ConvertFrom-FmkWmiDate {
    <#
        WMI datetime to a DateTime, or $null when unparseable. Get-CimInstance
        already converts datetime properties to DateTime objects, so both the
        raw DMTF string and a ready DateTime are accepted.
    #>
    param($Dmtf)
    try {
        if ($Dmtf -is [datetime]) { return $Dmtf }
        if ([string]::IsNullOrEmpty("$Dmtf")) { return $null }
        return [System.Management.ManagementDateTimeConverter]::ToDateTime([string]$Dmtf)
    }
    catch { return $null }
}

function Format-FmkWmiDate {
    param($Dmtf)
    try {
        $dt = ConvertFrom-FmkWmiDate $Dmtf
        if ($null -eq $dt) { return 'Unavailable' }
        return $dt.ToString('yyyy-MM-dd HH:mm')
    }
    catch { return 'Unavailable' }
}

function Format-FmkUptime {
    param($BootTime)
    try {
        $ts = (Get-Date) - [datetime]$BootTime
        if ($ts.TotalSeconds -lt 0) { return 'Unavailable' }
        $d = [int][Math]::Floor($ts.TotalDays)
        $out = ''
        if ($d -eq 1) { $out = '1 day, ' }
        elseif ($d -gt 1) { $out = "$d days, " }
        return ('{0}{1} h {2} min' -f $out, $ts.Hours, $ts.Minutes)
    }
    catch { return 'Unavailable' }
}

function Get-FmkMemoryTypeName {
    <# SMBIOS memory-device type codes; unknown codes report their code. #>
    param($Code)
    try { $c = [int]$Code } catch { return 'Unknown' }
    switch ($c) {
        0x01 { return 'Other' }
        0x02 { return 'Unknown' }
        0x03 { return 'DRAM' }
        0x0F { return 'SDRAM' }
        0x11 { return 'RDRAM' }
        0x12 { return 'DDR' }
        0x13 { return 'DDR2' }
        0x18 { return 'DDR3' }
        0x1A { return 'DDR4' }
        0x1B { return 'LPDDR' }
        0x1C { return 'LPDDR2' }
        0x1D { return 'LPDDR3' }
        0x1E { return 'LPDDR4' }
        0x22 { return 'DDR5' }
        0x23 { return 'LPDDR5' }
        default { return ('Unknown (SMBIOS code 0x{0:X2})' -f $c) }
    }
}

function Get-FmkCpuArchName {
    param($Code)
    try { $c = [int]$Code } catch { return 'Unknown' }
    switch ($c) {
        0 { return 'x86' }
        5 { return 'ARM' }
        6 { return 'Itanium' }
        9 { return 'x64' }
        12 { return 'ARM64' }
        default { return 'Unknown' }
    }
}

function Get-FmkOsInstallTypeName {
    param($Code)
    try { $c = [int]$Code } catch { return 'Unknown' }
    switch ($c) {
        1 { return 'Client (workstation)' }
        2 { return 'Domain controller' }
        3 { return 'Server' }
        default { return 'Unknown' }
    }
}

function Get-FmkEccName {
    param($Code)
    try { $c = [int]$Code } catch { return 'Unknown' }
    switch ($c) {
        1 { return 'Other' }
        2 { return 'Unknown' }
        3 { return 'None' }
        4 { return 'Parity' }
        5 { return 'Single-bit ECC' }
        6 { return 'Multi-bit ECC' }
        7 { return 'CRC' }
        default { return 'Unknown' }
    }
}

function Format-FmkLinkSpeed {
    param($BitsPerSecond)
    try {
        $b = [double]$BitsPerSecond
        if ($b -le 0) { return 'Unavailable' }
        if ($b -ge 1GB) { return ('{0:N1} Gbps' -f ($b / 1GB)) }
        if ($b -ge 1MB) { return ('{0:N0} Mbps' -f ($b / 1MB)) }
        return ('{0:N0} bps' -f $b)
    }
    catch { return 'Unavailable' }
}

function ConvertFrom-FmkWslLine {
    <# wsl.exe emits UTF-16, which arrives NUL-interleaved; strip and trim. #>
    param($Line)
    try { return ([string]$Line).Replace([string][char]0, '').Trim() }
    catch { return '' }
}

function ConvertFrom-FmkMonitorString {
    <# WMI monitor uint16 arrays are NUL-padded ASCII; decode and strip. #>
    param($Words)
    try {
        $chars = foreach ($w in @($Words)) { [char][int]$w }
        return (-join $chars).Trim([char]0).Trim()
    }
    catch { return '' }
}

function Get-FmkFirmwareModeAndSecureBoot {
    <#
        Shared UEFI/Secure Boot probe for the Windows and Firmware sections.
        Confirm-SecureBootUEFI answers on UEFI and throws on legacy boot, so a
        boolean result itself proves UEFI. A privilege-flavored failure is
        reported as an access limit, not as legacy hardware.
    #>
    try {
        $sb = Confirm-SecureBootUEFI -ErrorAction Stop
        return @{ Mode = 'UEFI'; SecureBoot = $sb }
    }
    catch {
        if ("$($_.Exception.Message)" -match '(?i)privilege|access|denied|require') {
            return @{ Mode = 'Unknown (administrator rights may be required)'; SecureBoot = $null }
        }
        return @{ Mode = 'Legacy BIOS or unknown'; SecureBoot = $null }
    }
}

function Add-FmkSecureBootRow {
    <# Secure Boot row matching the firmware probe outcome. #>
    param($Section, $Firmware)
    if ($null -eq $Firmware.SecureBoot) {
        if ([string]$Firmware.Mode -like 'Unknown*') {
            Add-FmkSysRow $Section 'Secure Boot' 'Unavailable (see firmware boot mode)'
        }
        else {
            Add-FmkSysRow $Section 'Secure Boot' 'Unavailable (not a UEFI boot)'
        }
    }
    elseif ([bool]$Firmware.SecureBoot) { Add-FmkSysRow $Section 'Secure Boot' 'Enabled' }
    else { Add-FmkSysRow $Section 'Secure Boot' 'Disabled' }
}

function Get-FmkSysWindowsSection {
    param([bool]$IsAdmin)
    $s = New-FmkSysSection -Key 'Windows' -Title 'Windows and Operating System'
    $os = Read-FmkCim 'Win32_OperatingSystem'
    if ($os -and $os.Count -gt 0) { $os = $os[0] } else { $os = $null }
    if (-not $os) { Add-FmkSysWarning $s 'Win32_OperatingSystem returned nothing.' }

    $reg = @{}
    try {
        $rk = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        foreach ($n in @('ProductName', 'EditionID', 'DisplayVersion', 'ReleaseId', 'CurrentBuild', 'UBR', 'InstallDate')) {
            $reg[$n] = $rk.$n
        }
    }
    catch { Add-FmkSysWarning $s 'Version registry key unreadable.' }

    $osCaption = Get-FmkSysField $os @('Caption')
    $product = Get-FmkSysField @{ V = $reg['ProductName'] } @('V') -Missing ''
    if ($product -eq '') { $product = $osCaption }
    Add-FmkSysRow $s 'Product name' $product
    Add-FmkSysRow $s 'Edition' (Get-FmkSysField @{ V = $reg['EditionID'] } @('V'))
    $release = Get-FmkSysField @{ V = $reg['DisplayVersion'] } @('V') -Missing ''
    if ($release -eq '') {
        $rid = Get-FmkSysField @{ V = $reg['ReleaseId'] } @('V') -Missing ''
        if ($rid -ne '') { $release = "Release $rid" } else { $release = 'Unavailable' }
    }
    Add-FmkSysRow $s 'Release' $release
    $build = Get-FmkSysField @{ V = $reg['CurrentBuild'] } @('V') -Missing ''
    $ubr = Get-FmkSysField @{ V = $reg['UBR'] } @('V') -Missing ''
    if ($build -ne '' -and $ubr -ne '') { $buildText = "$build.$ubr" }
    elseif ($build -ne '') { $buildText = $build }
    else { $buildText = Get-FmkSysField $os @('BuildNumber') }
    Add-FmkSysRow $s 'Build' $buildText
    Add-FmkSysRow $s 'OS version string' (Get-FmkSysField $os @('Version'))
    $arch = Get-FmkSysField $os @('OSArchitecture') -Missing ''
    if ($arch -eq '') {
        if ([System.Environment]::Is64BitOperatingSystem) { $arch = '64-bit' }
        else { $arch = 'Unavailable' }
    }
    Add-FmkSysRow $s 'Architecture' $arch
    Add-FmkSysRow $s 'Installation type' (Get-FmkOsInstallTypeName (Get-FmkSysField $os @('ProductType') -Missing -1))

    $installText = 'Unavailable'
    $wmiInstall = ConvertFrom-FmkWmiDate (Get-FmkSysField $os @('InstallDate') -Missing '')
    if ($wmiInstall) { $installText = $wmiInstall.ToString('yyyy-MM-dd HH:mm') }
    else {
        try {
            $epoch = $reg['InstallDate']
            if ($null -ne $epoch -and [int64]$epoch -gt 0) {
                $installText = (Get-Date '1970-01-01Z').ToLocalTime().AddSeconds([int64]$epoch).ToString('yyyy-MM-dd HH:mm')
            }
        }
        catch { }
    }
    Add-FmkSysRow $s 'Installed on' $installText

    $bootDt = ConvertFrom-FmkWmiDate (Get-FmkSysField $os @('LastBootUpTime') -Missing '')
    if ($bootDt) {
        Add-FmkSysRow $s 'Last boot' ($bootDt.ToString('yyyy-MM-dd HH:mm'))
        Add-FmkSysRow $s 'Uptime' (Format-FmkUptime $bootDt)
    }
    else {
        Add-FmkSysRow $s 'Last boot' 'Unavailable'
        Add-FmkSysRow $s 'Uptime' 'Unavailable'
        Add-FmkSysWarning $s 'Boot time unreadable, so uptime cannot be computed.'
    }

    try {
        $tz = Get-TimeZone -ErrorAction Stop
        Add-FmkSysRow $s 'Time zone' ("$($tz.Id) ($($tz.DisplayName))")
    }
    catch {
        $wtz = Read-FmkCim 'Win32_TimeZone'
        if ($wtz -and $wtz.Count -gt 0) { Add-FmkSysRow $s 'Time zone' (Get-FmkSysField $wtz[0] @('Caption', 'Description')) }
        else {
            Add-FmkSysRow $s 'Time zone' 'Unavailable'
            Add-FmkSysWarning $s 'Time zone unreadable.'
        }
    }
    try {
        $sysLoc = (Get-WinSystemLocale -ErrorAction Stop).Name
        Add-FmkSysRow $s 'System locale' $sysLoc
    }
    catch {
        Add-FmkSysRow $s 'System locale' 'Unavailable'
        Add-FmkSysWarning $s 'System locale unreadable.'
    }
    try {
        $cult = Get-Culture -ErrorAction Stop
        Add-FmkSysRow $s 'Display language' ("$($cult.DisplayName) ($($cult.Name))")
    }
    catch {
        Add-FmkSysRow $s 'Display language' 'Unavailable'
        Add-FmkSysWarning $s 'Display language unreadable.'
    }

    Add-FmkSysRow $s 'Computer name (potentially sensitive)' $env:COMPUTERNAME -Sensitivity 'Identifying'
    $who = "$env:USERDOMAIN\$env:USERNAME"
    Add-FmkSysRow $s 'Current user (potentially sensitive)' $who -Sensitivity 'Identifying'

    $fw = Get-FmkFirmwareModeAndSecureBoot
    Add-FmkSysRow $s 'Firmware boot mode' $fw.Mode
    Add-FmkSecureBootRow $s $fw

    $dg = Read-FmkCim 'Win32_DeviceGuard'
    if ($dg -and $dg.Count -gt 0) {
        $running = @()
        try { $running = @($dg[0].SecurityServicesRunning) } catch { $running = @() }
        if ($running.Count -gt 0) { Add-FmkSysRow $s 'Virtualization-based security' ('On (running: ' + ($running -join ', ') + ')') }
        else {
            $vbs = -1
            try { $vbs = [int]$dg[0].VirtualizationBasedSecurityStatus } catch { }
            if ($vbs -eq 1) { Add-FmkSysRow $s 'Virtualization-based security' 'On' }
            elseif ($vbs -eq 0) { Add-FmkSysRow $s 'Virtualization-based security' 'Off' }
            else { Add-FmkSysRow $s 'Virtualization-based security' 'Unavailable' }
        }
    }
    else {
        Add-FmkSysRow $s 'Virtualization-based security' 'Unavailable'
        Add-FmkSysWarning $s 'Device Guard status not exposed on this system.'
    }

    if ($IsAdmin) { Add-FmkSysRow $s 'Running as administrator' 'Yes' }
    else { Add-FmkSysRow $s 'Running as administrator' 'No' }
    return $s
}

function Get-FmkSysCpuSection {
    param([bool]$IsAdmin)
    $s = New-FmkSysSection -Key 'CPU' -Title 'Processor (CPU)'
    $cpus = Read-FmkCim 'Win32_Processor'
    if (-not $cpus -or $cpus.Count -eq 0) {
        Add-FmkSysWarning $s 'Win32_Processor returned nothing.'
        Add-FmkSysRow $s 'Processors' 'Unavailable'
        return $s
    }
    $multi = $cpus.Count -gt 1
    $idx = 0
    foreach ($c in $cpus) {
        $idx++
        $tag = ''
        if ($multi) { $tag = "Socket $idx " }
        Add-FmkSysRow $s ($tag + 'Model') (Get-FmkSysField $c @('Name'))
        Add-FmkSysRow $s ($tag + 'Manufacturer') (Get-FmkSysField $c @('Manufacturer'))
        $fam = Get-FmkSysField $c @('Family') -Missing ''
        $mod = Get-FmkSysField $c @('Model') -Missing ''
        $stp = Get-FmkSysField $c @('Stepping') -Missing ''
        $fmsParts = @()
        if ($fam -ne '') { $fmsParts += "Family $fam" }
        if ($mod -ne '') { $fmsParts += "model $mod" }
        if ($stp -ne '') { $fmsParts += "stepping $stp" }
        if ($fmsParts.Count -eq 0) {
            Add-FmkSysRow $s ($tag + 'Family / model / stepping') 'Unavailable'
        }
        else {
            Add-FmkSysRow $s ($tag + 'Family / model / stepping') ($fmsParts -join ', ')
        }
        Add-FmkSysRow $s ($tag + 'Architecture') (Get-FmkCpuArchName (Get-FmkSysField $c @('Architecture') -Missing -1))
        Add-FmkSysRow $s ($tag + 'Processor ID') (Get-FmkSysField $c @('ProcessorId'))
        Add-FmkSysRow $s ($tag + 'Physical cores') (Get-FmkSysField $c @('NumberOfCores'))
        Add-FmkSysRow $s ($tag + 'Logical processors') (Get-FmkSysField $c @('NumberOfLogicalProcessors'))
        Add-FmkSysRow $s ($tag + 'Base frequency (advertised)') (Format-FmkGigahertz (Get-FmkSysField $c @('MaxClockSpeed') -Missing 0))
        Add-FmkSysRow $s ($tag + 'Current frequency (measured at collection)') (Format-FmkGigahertz (Get-FmkSysField $c @('CurrentClockSpeed') -Missing 0))
        $l2 = Get-FmkSysField $c @('L2CacheSize') -Missing $null
        if ($null -eq $l2) { Add-FmkSysRow $s ($tag + 'L2 cache') 'Unavailable' }
        else { Add-FmkSysRow $s ($tag + 'L2 cache') (Format-FmkBytes ([double]$l2 * 1KB)) }
        $l3 = Get-FmkSysField $c @('L3CacheSize') -Missing $null
        if ($null -eq $l3) { Add-FmkSysRow $s ($tag + 'L3 cache') 'Unavailable' }
        else { Add-FmkSysRow $s ($tag + 'L3 cache') (Format-FmkBytes ([double]$l3 * 1KB)) }
        $virt = Get-FmkSysField $c @('VirtualizationFirmwareEnabled') -Missing $null
        if ($null -eq $virt) { Add-FmkSysRow $s ($tag + 'Hardware virtualization') 'Unavailable' }
        elseif ([bool]$virt) { Add-FmkSysRow $s ($tag + 'Hardware virtualization') 'Enabled in firmware' }
        else { Add-FmkSysRow $s ($tag + 'Hardware virtualization') 'Disabled or unsupported' }
        $slat = Get-FmkSysField $c @('SecondLevelAddressTranslationExtensions') -Missing $null
        if ($null -eq $slat) { Add-FmkSysRow $s ($tag + 'Second-level address translation') 'Unavailable' }
        elseif ([bool]$slat) { Add-FmkSysRow $s ($tag + 'Second-level address translation') 'Supported' }
        else { Add-FmkSysRow $s ($tag + 'Second-level address translation') 'Not supported' }
    }
    Add-FmkSysRow $s 'Instruction sets' 'Unavailable from inbox interfaces (no CPUID source); x64 implies an x86-64 baseline.'
    $numa = Read-FmkCim 'Win32_NumaNode'
    if ($numa -and $numa.Count -gt 0) {
        Add-FmkSysRow $s 'NUMA nodes' $numa.Count
    }
    else {
        Add-FmkSysRow $s 'NUMA nodes' 'Unavailable'
        Add-FmkSysWarning $s 'NUMA topology is not exposed on this system.'
    }
    return $s
}

function Get-FmkSysMemorySection {
    param([bool]$IsAdmin)
    $s = New-FmkSysSection -Key 'Memory' -Title 'Memory (RAM)'
    $cs = Read-FmkCim 'Win32_ComputerSystem'
    if ($cs -and $cs.Count -gt 0) { $cs = $cs[0] } else { $cs = $null }
    $os = Read-FmkCim 'Win32_OperatingSystem'
    if ($os -and $os.Count -gt 0) { $os = $os[0] } else { $os = $null }
    if (-not $cs -and -not $os) {
        Add-FmkSysWarning $s 'Memory counters unreadable.'
        Add-FmkSysRow $s 'Memory' 'Unavailable'
        return $s
    }

    $installed = Get-FmkSysField $cs @('TotalPhysicalMemory') -Missing $null
    if ($null -eq $installed) { Add-FmkSysRow $s 'Installed physical memory' 'Unavailable' }
    else { Add-FmkSysRow $s 'Installed physical memory' (Format-FmkBytes $installed) }

    $totalVis = Get-FmkSysField $os @('TotalVisibleMemorySize') -Missing $null
    $freePhys = Get-FmkSysField $os @('FreePhysicalMemory') -Missing $null
    if ($null -eq $totalVis -or $null -eq $freePhys) {
        Add-FmkSysRow $s 'Total visible to OS' 'Unavailable'
        Add-FmkSysRow $s 'Available' 'Unavailable'
        Add-FmkSysRow $s 'In use' 'Unavailable'
        Add-FmkSysWarning $s 'Live usage counters unreadable.'
    }
    else {
        $totalB = [double]$totalVis * 1KB
        $freeB = [double]$freePhys * 1KB
        $usedB = $totalB - $freeB
        if ($usedB -lt 0) { $usedB = 0 }
        $pct = 0
        if ($totalB -gt 0) { $pct = [int][Math]::Round($usedB / $totalB * 100) }
        Add-FmkSysRow $s 'Total visible to OS' (Format-FmkBytes $totalB)
        Add-FmkSysRow $s 'Available' (Format-FmkBytes $freeB)
        Add-FmkSysRow $s 'In use' ((Format-FmkBytes $usedB) + " ($pct%)")
    }

    $totalVirt = Get-FmkSysField $os @('TotalVirtualMemorySize') -Missing $null
    $freeVirt = Get-FmkSysField $os @('FreeVirtualMemory') -Missing $null
    if ($null -ne $totalVirt -and $null -ne $freeVirt -and $null -ne $totalVis -and $null -ne $freePhys) {
        $pfTotal = ([double]$totalVirt - [double]$totalVis) * 1KB
        $pfFree = ([double]$freeVirt - [double]$freePhys) * 1KB
        if ($pfTotal -lt 0) { $pfTotal = 0 }
        if ($pfFree -lt 0) { $pfFree = 0 }
        Add-FmkSysRow $s 'Page file total' (Format-FmkBytes $pfTotal)
        Add-FmkSysRow $s 'Page file free' (Format-FmkBytes $pfFree)
    }
    else {
        Add-FmkSysRow $s 'Page file total' 'Unavailable'
        Add-FmkSysWarning $s 'Page-file counters unreadable.'
    }

    $mods = Read-FmkCim 'Win32_PhysicalMemory'
    $arr = Read-FmkCim 'Win32_PhysicalMemoryArray'
    $arrInfo = $null
    if ($arr -and $arr.Count -gt 0) { $arrInfo = $arr[0] }
    $populated = @($mods).Count
    $slots = -1
    if ($arrInfo) {
        try { $slots = [int]$arrInfo.MemoryDevices } catch { $slots = -1 }
    }
    if ($populated -eq 0) {
        Add-FmkSysRow $s 'Memory modules' 'Unavailable'
        Add-FmkSysWarning $s 'No memory modules reported. Some laptops and systems with soldered memory expose totals but no per-module details through Windows; the totals above remain accurate.'
    }
    else {
        $slotText = "$populated module(s)"
        if ($slots -gt $populated) { $slotText += " in $slots slots ($($slots - $populated) empty)" }
        elseif ($slots -eq $populated -and $slots -gt 0) { $slotText += " in $slots slots (none empty)" }
        Add-FmkSysRow $s 'Memory modules' $slotText
        $mi = 0
        foreach ($m in @($mods)) {
            $mi++
            # DeviceLocator first: BankLabel repeats on some firmware (two
            # sticks both answering "BANK 0"), while the locator names the
            # actual slot.
            $loc = Get-FmkSysField $m @('DeviceLocator', 'BankLabel') -Missing "#$mi"
            $cap = Get-FmkSysField $m @('Capacity') -Missing $null
            $capText = 'Unknown size'
            if ($null -ne $cap) { $capText = Format-FmkBytes $cap }
            $typeName = Get-FmkMemoryTypeName (Get-FmkSysField $m @('SMBIOSMemoryType') -Missing -1)
            $spd = Get-FmkSysField $m @('ConfiguredClockSpeed', 'Speed') -Missing 0
            $speedText = ''
            if ($spd -and [double]$spd -gt 0) { $speedText = "-$([int]$spd)" }
            $mfr = Get-FmkSysField $m @('Manufacturer') -Missing ''
            $part = Get-FmkSysField $m @('PartNumber') -Missing ''
            if ($mfr -ne '') { $mfr = $mfr.Trim() }
            if ($part -ne '') { $part = $part.Trim() }
            $bits = @($capText)
            $typeSpeed = ($typeName + $speedText).Trim()
            if ($typeSpeed -ne '' -and $typeSpeed -ne 'Unknown') { $bits += $typeSpeed }
            if ($mfr -ne '') { $bits += $mfr }
            if ($part -ne '') { $bits += $part }
            Add-FmkSysRow $s "Module $loc" ($bits -join ' / ')
        }
    }
    if ($arrInfo) {
        Add-FmkSysRow $s 'Error correction' (Get-FmkEccName (Get-FmkSysField $arrInfo @('MemoryErrorCorrection') -Missing -1))
        $depth = Get-FmkSysField $arrInfo @('InterleaveDataDepth') -Missing $null
        if ($null -eq $depth) { Add-FmkSysRow $s 'Interleave depth' 'Unavailable' }
        else { Add-FmkSysRow $s 'Interleave depth' "$depth" }
    }
    else {
        Add-FmkSysWarning $s 'Memory array details (ECC, slots) not exposed on this system.'
    }
    return $s
}

function Get-FmkSysGpuSection {
    param([bool]$IsAdmin)
    $s = New-FmkSysSection -Key 'GPU' -Title 'Graphics (GPU)'
    $gpus = Read-FmkCim 'Win32_VideoController'
    if (-not $gpus -or $gpus.Count -eq 0) {
        Add-FmkSysWarning $s 'No video controllers reported.'
        Add-FmkSysRow $s 'Graphics adapters' 'Unavailable'
        return $s
    }
    $multi = $gpus.Count -gt 1
    Add-FmkSysRow $s 'Adapter count' $gpus.Count
    $idx = 0
    foreach ($g in $gpus) {
        $idx++
        $tag = ''
        if ($multi) { $tag = "Adapter $idx " }
        Add-FmkSysRow $s ($tag + 'Name') (Get-FmkSysField $g @('Name'))
        Add-FmkSysRow $s ($tag + 'Vendor') (Get-FmkSysField $g @('AdapterCompatibility'))
        Add-FmkSysRow $s ($tag + 'Status') (Get-FmkSysField $g @('Status'))
        $pnp = Get-FmkSysField $g @('PNPDeviceID') -Missing ''
        if ($pnp -ne '' -and $pnp -match 'VEN_([0-9A-Fa-f]{4}).*DEV_([0-9A-Fa-f]{4})') {
            Add-FmkSysRow $s ($tag + 'PCI hardware ID') ("VEN $($Matches[1].ToUpper()) DEV $($Matches[2].ToUpper())")
        }
        else {
            Add-FmkSysRow $s ($tag + 'PCI hardware ID') 'Unavailable'
        }
        $vram = Get-FmkSysField $g @('AdapterRAM') -Missing $null
        if ($null -eq $vram) { Add-FmkSysRow $s ($tag + 'Dedicated video memory') 'Unavailable' }
        else { Add-FmkSysRow $s ($tag + 'Dedicated video memory') (Format-FmkBytes $vram) }
        Add-FmkSysRow $s ($tag + 'Driver version') (Get-FmkSysField $g @('DriverVersion'))
        Add-FmkSysRow $s ($tag + 'Driver date') (Format-FmkWmiDate (Get-FmkSysField $g @('DriverDate') -Missing ''))
        $mode = Get-FmkSysField $g @('VideoModeDescription') -Missing ''
        # A 32-bit color depth is the driver's "unknown" sentinel, not eight
        # billion colors; drop it rather than printing nonsense.
        if ($mode -match '(\d+)\s+colors?$') {
            try {
                if ([double]$Matches[1] -gt 64) {
                    $mode = $mode.Substring(0, $mode.LastIndexOf($Matches[0])).TrimEnd()
                }
            }
            catch { }
        }
        if ($mode -eq '') {
            $w = Get-FmkSysField $g @('CurrentHorizontalResolution') -Missing $null
            $h = Get-FmkSysField $g @('CurrentVerticalResolution') -Missing $null
            $r = Get-FmkSysField $g @('CurrentRefreshRate') -Missing $null
            if ($null -ne $w -and $null -ne $h) {
                $mode = "${w} x ${h}"
                if ($null -ne $r -and [int]$r -gt 1) { $mode += " @ ${r} Hz" }
            }
            else { $mode = 'Unavailable' }
        }
        Add-FmkSysRow $s ($tag + 'Current display mode') $mode
    }
    Add-FmkSysWarning $s 'AdapterRAM is a 32-bit field and may cap at 4 GB on some systems.'
    Add-FmkSysRow $s 'Shared GPU memory' 'Unavailable (requires a DXGI query)'
    $dxv = ''
    try {
        $dxv = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\DirectX' -Name 'Version' -ErrorAction Stop).Version
    }
    catch { $dxv = '' }
    # A 4.09 stamp is a leftover DirectX 9 version key, not the installed
    # DirectX; reporting it as the version would mislead.
    if ($dxv -eq '') { Add-FmkSysRow $s 'DirectX version' 'Unavailable' }
    elseif ("$dxv" -like '4.09*') {
        Add-FmkSysRow $s 'DirectX version' 'Unavailable (registry holds only a legacy DirectX 9 version stamp)'
    }
    else { Add-FmkSysRow $s 'DirectX version' $dxv }
    Add-FmkSysRow $s 'DirectX feature level' 'Unavailable (no inbox source)'
    Add-FmkSysRow $s 'WDDM driver model' 'Unavailable (no inbox source)'
    Add-FmkSysRow $s 'HDR / variable refresh' 'Unavailable (requires display-core query)'
    return $s
}

function Get-FmkBusTypeName {
    param($Code)
    try { $c = [int]$Code } catch { return 'Unknown' }
    switch ($c) {
        0 { return 'Unknown' }
        7 { return 'USB' }
        8 { return 'RAID' }
        11 { return 'SATA' }
        14 { return 'Virtual' }
        15 { return 'File-backed virtual' }
        16 { return 'Storage Spaces' }
        17 { return 'NVMe' }
        default { return 'Unknown' }
    }
}

function Get-FmkDiskMediaName {
    param($Code)
    try { $c = [int]$Code } catch { return 'Unknown' }
    switch ($c) {
        3 { return 'HDD' }
        4 { return 'SSD' }
        5 { return 'SCM' }
        default { return 'Unknown' }
    }
}

function Get-FmkOpStatusName {
    <# DMTF operational-status codes shared by the MSFT storage classes. #>
    param($Code)
    try { $c = [int]$Code } catch { return 'Unknown' }
    switch ($c) {
        0 { return 'Unknown' }
        1 { return 'Other' }
        2 { return 'OK' }
        3 { return 'Degraded' }
        4 { return 'Stressed' }
        5 { return 'Predictive failure' }
        6 { return 'Error' }
        default { return 'Unknown' }
    }
}

function Get-FmkDiskHealthName {
    param($Code)
    try { $c = [int]$Code } catch { return 'Unknown' }
    switch ($c) {
        0 { return 'Healthy' }
        1 { return 'Warning' }
        2 { return 'Unhealthy' }
        default { return 'Unknown' }
    }
}

function Get-FmkDriveTypeName {
    param($Code)
    try { $c = [int]$Code } catch { return 'Unknown' }
    switch ($c) {
        0 { return 'Unknown' }
        1 { return 'No root directory' }
        2 { return 'Removable' }
        3 { return 'Local fixed disk' }
        4 { return 'Network drive' }
        5 { return 'Compact disc' }
        6 { return 'RAM disk' }
        default { return 'Unknown' }
    }
}

function Get-FmkSysStorageSection {
    param([bool]$IsAdmin)
    $s = New-FmkSysSection -Key 'Storage' -Title 'Storage'
    Add-FmkSysRow $s 'Scope note' 'Physical disks and logical volumes are listed separately below.'
    $disks = Read-FmkCim 'Win32_DiskDrive'
    if (-not $disks -or $disks.Count -eq 0) {
        Add-FmkSysWarning $s 'No physical disks reported.'
        Add-FmkSysRow $s 'Physical disks' 'Unavailable'
        return $s
    }

    $osDrive = ''
    try {
        $osq = Read-FmkCim 'Win32_OperatingSystem'
        if ($osq -and $osq.Count -gt 0) { $osDrive = [string]$osq[0].SystemDrive }
    }
    catch { }

    # Disk index -> drive letters, best effort through the association classes.
    $letterMap = @{}
    $bootParts = @{}
    try {
        $partToDisk = @{}
        foreach ($a in @(Read-FmkCim 'Win32_DiskDriveToDiskPartition')) {
            $di = -1
            $pi = -1
            if ([string]$a.Antecedent -match 'PHYSICALDRIVE(\d+)') { $di = [int]$Matches[1] }
            if ([string]$a.Dependent -match 'Disk #(\d+),\s*Partition #(\d+)') {
                if ([int]$Matches[1] -eq $di) { $pi = [int]$Matches[2] }
            }
            if ($di -ge 0 -and $pi -ge 0) { $partToDisk["$di/$pi"] = $di }
        }
        foreach ($a in @(Read-FmkCim 'Win32_LogicalDiskToPartition')) {
            $letter = ''
            if ([string]$a.Dependent -match 'DeviceID="([A-Za-z]:)"') { $letter = $Matches[1] }
            $key = ''
            if ([string]$a.Antecedent -match 'Disk #(\d+),\s*Partition #(\d+)') { $key = "$($Matches[1])/$($Matches[2])" }
            if ($letter -ne '' -and $key -ne '' -and $partToDisk.ContainsKey($key)) {
                $di = $partToDisk[$key]
                if (-not $letterMap.ContainsKey($di)) { $letterMap[$di] = @() }
                $letterMap[$di] += $letter
            }
        }
        foreach ($p in @(Read-FmkCim 'Win32_DiskPartition')) {
            try {
                if ([bool]$p.BootPartition -and ("$($p.DeviceID)" -match 'Disk #(\d+)')) {
                    $bootParts[[int]$Matches[1]] = $true
                }
            }
            catch { }
        }
    }
    catch { Add-FmkSysWarning $s 'Disk-to-volume mapping unavailable; volumes listed separately.' }

    $di = -1
    foreach ($d in $disks) {
        $di++
        $model = Get-FmkSysField $d @('Model')
        $size = Get-FmkSysField $d @('Size') -Missing $null
        $sizeText = 'Unknown size'
        if ($null -ne $size) { $sizeText = Format-FmkBytes $size }
        $iface = Get-FmkSysField $d @('InterfaceType') -Missing ''
        $media = Get-FmkSysField $d @('MediaType') -Missing ''
        $kind = $iface
        if ($media -ne '' -and $media -ne $iface) {
            if ($kind -ne '') { $kind += ' / ' }
            $kind += $media
        }
        if ($kind -eq '') { $kind = 'Unknown interface' }
        $tags = @()
        if ($bootParts.ContainsKey($di)) { $tags += 'boot disk' }
        $tagText = ''
        if ($tags.Count -gt 0) { $tagText = ' (' + ($tags -join ', ') + ')' }
        Add-FmkSysRow $s "Disk $di" ("$model -- $sizeText, $kind$tagText")
        Add-FmkSysRow $s "Disk $di manufacturer" (Get-FmkSysField $d @('Manufacturer'))
        Add-FmkSysRow $s "Disk $di partitions" (Get-FmkSysField $d @('Partitions'))
        Add-FmkSysRow $s "Disk $di status" (Get-FmkSysField $d @('Status'))
        Add-FmkSysRow $s "Disk $di serial number" (Get-FmkSysField $d @('SerialNumber')) -Sensitivity 'Private'
        if ($letterMap.ContainsKey($di)) {
            foreach ($letter in @($letterMap[$di])) {
                # Device IDs already end in a colon ("C:"); the row renderer
                # adds its own, so strip it here to avoid "Volume C::".
                Add-FmkSysRow $s ("Disk $di volume " + $letter.TrimEnd(':')) (Get-FmkVolumeText $letter $osDrive)
            }
        }
    }

    $msft = Read-FmkCim 'MSFT_PhysicalDisk' -Namespace 'root/Microsoft/Windows/Storage'
    if ($msft -and $msft.Count -gt 0) {
        $mi = -1
        foreach ($m in $msft) {
            $mi++
            $fname = Get-FmkSysField $m @('FriendlyName')
            $health = Get-FmkDiskHealthName (Get-FmkSysField $m @('HealthStatus') -Missing -1)
            $opRaw = Get-FmkSysField $m @('OperationalStatus') -Missing $null
            $opText = 'Unknown'
            if ($null -ne $opRaw) { $opText = ((@($opRaw) | ForEach-Object { Get-FmkOpStatusName $_ }) -join ', ') }
            $bus = Get-FmkBusTypeName (Get-FmkSysField $m @('BusType') -Missing -1)
            $mediaName = Get-FmkDiskMediaName (Get-FmkSysField $m @('MediaType') -Missing -1)
            $msize = Get-FmkSysField $m @('Size') -Missing $null
            $msizeText = 'Unknown size'
            if ($null -ne $msize) { $msizeText = Format-FmkBytes $msize }
            Add-FmkSysRow $s "Health entry $($mi + 1)" ("$fname -- health: $health, state: $opText, bus: $bus, media: $mediaName, size: $msizeText")
        }
    }
    else {
        Add-FmkSysWarning $s 'Storage health details unavailable (the Storage namespace may require administrator rights).'
    }

    $smart = $null
    try { $smart = Read-FmkCim 'MSStorageDriver_FailurePredictStatus' -Namespace 'root/wmi' } catch { $smart = $null }
    if ($smart -and $smart.Count -gt 0) {
        $si = 0
        foreach ($entry in $smart) {
            $si++
            $pred = $null
            try { $pred = [bool]$entry.PredictFailure } catch { $pred = $null }
            if ($null -eq $pred) { Add-FmkSysRow $s "SMART prediction (drive $si)" 'Unavailable' }
            elseif ($pred) { Add-FmkSysRow $s "SMART prediction (drive $si)" 'Failure predicted -- back up and replace the drive' }
            else { Add-FmkSysRow $s "SMART prediction (drive $si)" 'OK (no failure predicted)' }
        }
    }
    else {
        Add-FmkSysWarning $s 'SMART status unavailable (often requires administrator rights).'
    }

    $vols = Read-FmkCim 'Win32_LogicalDisk'
    if ($vols -and $vols.Count -gt 0) {
        foreach ($v in $vols) {
            $letter = Get-FmkSysField $v @('DeviceID') -Missing ''
            if ($letter -eq '') { continue }
            $mapped = $false
            foreach ($letters in @($letterMap.Values)) {
                if (@($letters) -contains $letter) { $mapped = $true }
            }
            if (-not $mapped) {
                Add-FmkSysRow $s ("Volume " + $letter.TrimEnd(':')) (Get-FmkVolumeText $letter $osDrive)
            }
        }
    }
    return $s
}

function Get-FmkVolumeText {
    <# One-line volume summary; a zero-byte volume stays zero, never 'Unavailable'. #>
    param([string]$Letter, [string]$OsDrive)
    try {
        $list = @(Get-CimInstance -ClassName 'Win32_LogicalDisk' -Filter "DeviceID='$Letter'" -ErrorAction Stop)
        if (-not $list -or $list.Count -eq 0) { return 'Unavailable' }
        $v = $list[0]
        $label = Get-FmkSysField $v @('VolumeName') -Missing ''
        $fs = Get-FmkSysField $v @('FileSystem') -Missing ''
        $size = Get-FmkSysField $v @('Size') -Missing $null
        $free = Get-FmkSysField $v @('FreeSpace') -Missing $null
        $dtype = Get-FmkDriveTypeName (Get-FmkSysField $v @('DriveType') -Missing -1)
        $bits = @()
        if ($label -ne '') { $bits += "($label)" }
        if ($fs -ne '') { $bits += $fs }
        $bits += $dtype
        if ($null -ne $size -and $null -ne $free) {
            $sz = [double]$size
            $fr = [double]$free
            if ($fr -lt 0) { $fr = 0 }
            if ($fr -gt $sz) { $fr = $sz }
            $pct = 0
            if ($sz -gt 0) { $pct = [int][Math]::Round($fr / $sz * 100) }
            $bits += ((Format-FmkBytes $fr) + ' free of ' + (Format-FmkBytes $sz) + " ($pct% free)")
        }
        else {
            $bits += 'size unavailable'
        }
        if ($OsDrive -ne '' -and $letter -ieq $OsDrive) { $bits += '(system drive)' }
        return ($bits -join ' ')
    }
    catch { return 'Unavailable' }
}

function Get-FmkSysFirmwareSection {
    param([bool]$IsAdmin)
    $s = New-FmkSysSection -Key 'Firmware' -Title 'Motherboard and Firmware'
    try {
        $csq = Read-FmkCim 'Win32_ComputerSystem'
        if ($csq -and $csq.Count -gt 0) {
            Add-FmkSysRow $s 'System manufacturer' (Get-FmkSysField $csq[0] @('Manufacturer'))
            Add-FmkSysRow $s 'System model' (Get-FmkSysField $csq[0] @('Model'))
            $fam = Get-FmkSysField $csq[0] @('SystemFamily') -Missing ''
            if ($fam -ne '') { Add-FmkSysRow $s 'System family' $fam }
        }
        else { Add-FmkSysWarning $s 'System identity unreadable.' }
    }
    catch { Add-FmkSysWarning $s 'System identity unreadable.' }
    $bb = Read-FmkCim 'Win32_BaseBoard'
    if ($bb -and $bb.Count -gt 0) {
        Add-FmkSysRow $s 'Motherboard manufacturer' (Get-FmkSysField $bb[0] @('Manufacturer'))
        Add-FmkSysRow $s 'Motherboard model' (Get-FmkSysField $bb[0] @('Product'))
        Add-FmkSysRow $s 'Motherboard version' (Get-FmkSysField $bb[0] @('Version'))
    }
    else {
        Add-FmkSysWarning $s 'Baseboard details not exposed on this system.'
        Add-FmkSysRow $s 'Motherboard' 'Unavailable'
    }
    $bios = Read-FmkCim 'Win32_BIOS'
    if ($bios -and $bios.Count -gt 0) {
        Add-FmkSysRow $s 'BIOS/UEFI vendor' (Get-FmkSysField $bios[0] @('Manufacturer'))
        $bver = Get-FmkSysField $bios[0] @('SMBIOSBIOSVersion') -Missing ''
        if ($bver -eq '') { $bver = Get-FmkSysField $bios[0] @('Version') }
        Add-FmkSysRow $s 'BIOS/UEFI version' $bver
        Add-FmkSysRow $s 'BIOS/UEFI date' (Format-FmkWmiDate (Get-FmkSysField $bios[0] @('ReleaseDate') -Missing ''))
    }
    else {
        Add-FmkSysWarning $s 'BIOS details not exposed on this system.'
        Add-FmkSysRow $s 'BIOS/UEFI' 'Unavailable'
    }
    $fw = Get-FmkFirmwareModeAndSecureBoot
    Add-FmkSysRow $s 'Firmware boot mode' $fw.Mode
    Add-FmkSecureBootRow $s $fw
    $tpmPresent = $null
    $tpmReady = $null
    $tpmVer = ''
    try {
        if (Get-Command Get-Tpm -ErrorAction SilentlyContinue) {
            $tpm = Get-Tpm -ErrorAction Stop
            $tpmPresent = [bool]$tpm.TpmPresent
            $tpmReady = [bool]$tpm.TpmReady
            try { $tpmVer = [string]$tpm.ManufacturerVersion } catch { $tpmVer = '' }
        }
    }
    catch { }
    if ($null -eq $tpmPresent) {
        Add-FmkSysRow $s 'TPM' 'Unavailable (may require administrator rights)'
        Add-FmkSysWarning $s 'TPM status unreadable with current permissions.'
    }
    else {
        if ($tpmPresent) {
            $tpmText = 'Present'
            if ($tpmReady) { $tpmText += ', ready' } else { $tpmText += ', not ready' }
            if ($tpmVer -ne '') { $tpmText += " (spec version $tpmVer)" }
            Add-FmkSysRow $s 'TPM' $tpmText
        }
        else { Add-FmkSysRow $s 'TPM' 'Not present' }
    }
    return $s
}

function Get-FmkSysNetworkSection {
    param([bool]$IsAdmin)
    $s = New-FmkSysSection -Key 'Network' -Title 'Network Adapters'
    $adapters = $null
    try {
        if (Get-Command Get-NetAdapter -ErrorAction SilentlyContinue) {
            $adapters = @(Get-NetAdapter -ErrorAction Stop)
        }
    }
    catch { $adapters = $null }
    if (-not $adapters -or $adapters.Count -eq 0) {
        Add-FmkSysWarning $s 'No network adapters reported.'
        Add-FmkSysRow $s 'Network adapters' 'Unavailable'
        return $s
    }
    $shown = 0
    $cap = 12
    foreach ($a in $adapters) {
        if ($shown -ge $cap) { break }
        $shown++
        $aname = Get-FmkSysField $a @('Name')
        Add-FmkSysRow $s "Adapter: $aname" (Get-FmkSysField $a @('InterfaceDescription'))
        Add-FmkSysRow $s "$aname status" (Get-FmkSysField $a @('Status'))
        Add-FmkSysRow $s "$aname link speed" (Get-FmkSysField $a @('LinkSpeed'))
        Add-FmkSysRow $s "$aname MAC address" (Get-FmkSysField $a @('MacAddress')) -Sensitivity 'Private'
        $drv = Read-FmkCim 'Win32_NetworkAdapter' -Filter ("NetConnectionID='" + ($aname -replace "'", "''") + "'")
        if ($drv -and $drv.Count -gt 0) {
            Add-FmkSysRow $s "$aname driver" (Get-FmkSysField $drv[0] @('DriverVersion'))
            $ddate = Format-FmkWmiDate (Get-FmkSysField $drv[0] @('DriverDate') -Missing '')
            Add-FmkSysRow $s "$aname driver date" $ddate
        }
        else {
            Add-FmkSysRow $s "$aname driver" 'Unavailable'
        }
    }
    if ($adapters.Count -gt $cap) {
        Add-FmkSysRow $s 'Adapter list' ("Showing $cap of $($adapters.Count) adapters.")
    }
    try {
        if (Get-Command Get-NetIPAddress -ErrorAction SilentlyContinue) {
            $ips = @(Get-NetIPAddress -ErrorAction Stop | Where-Object {
                    ($_.AddressFamily -eq 'IPv4' -or $_.AddressFamily -eq 'IPv6') -and
                    $_.IPAddress -ne '127.0.0.1' -and $_.IPAddress -ne '::1'
                })
            $shownIp = 0
            foreach ($ip in $ips) {
                if ($shownIp -ge 12) { break }
                $shownIp++
                $fam = 'IPv4'
                if ([string]$ip.AddressFamily -eq 'IPv6') { $fam = 'IPv6' }
                Add-FmkSysRow $s ("$($ip.InterfaceAlias) $fam") ("$($ip.IPAddress)/$($ip.PrefixLength)") -Sensitivity 'Private'
            }
            if ($ips.Count -eq 0) { Add-FmkSysWarning $s 'No non-loopback IP addresses found.' }
        }
    }
    catch { Add-FmkSysWarning $s 'IP addresses unreadable.' }
    try {
        if (Get-Command Get-NetIPInterface -ErrorAction SilentlyContinue) {
            $ifs = @(Get-NetIPInterface -ErrorAction Stop)
            $shownDhcp = 0
            foreach ($ni in $ifs) {
                if ($shownDhcp -ge 12) { break }
                $dh = Get-FmkSysField $ni @('Dhcp') -Missing $null
                if ($null -eq $dh) { continue }
                $shownDhcp++
                if ([string]$dh -eq 'Enabled') { Add-FmkSysRow $s ("$($ni.InterfaceAlias) DHCP") 'Enabled' }
                else { Add-FmkSysRow $s ("$($ni.InterfaceAlias) DHCP") 'Disabled' }
            }
        }
    }
    catch { Add-FmkSysWarning $s 'DHCP status unreadable.' }
    try {
        if (Get-Command Get-DnsClientServerAddress -ErrorAction SilentlyContinue) {
            $dns = @(Get-DnsClientServerAddress -ErrorAction Stop)
            $shownDns = 0
            foreach ($d in $dns) {
                $addrs = @()
                try { $addrs = @($d.ServerAddresses) } catch { $addrs = @() }
                if ($addrs.Count -eq 0) { continue }
                if ($shownDns -ge 8) { break }
                $shownDns++
                Add-FmkSysRow $s ("$($d.InterfaceAlias) DNS") ($addrs -join '; ') -Sensitivity 'Private'
            }
        }
    }
    catch { Add-FmkSysWarning $s 'DNS servers unreadable.' }
    try {
        if (Get-Command Get-NetRoute -ErrorAction SilentlyContinue) {
            $gw = @(Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop |
                Where-Object { $_.NextHop -ne '0.0.0.0' -and $_.NextHop -ne '' } |
                Select-Object -First 4)
            foreach ($r in $gw) {
                Add-FmkSysRow $s 'Default gateway' ([string]$r.NextHop) -Sensitivity 'Private'
            }
            if ($gw.Count -eq 0) { Add-FmkSysWarning $s 'No default IPv4 gateway found.' }
        }
    }
    catch { Add-FmkSysWarning $s 'Default gateway unreadable.' }
    return $s
}

function Get-FmkSysBatterySection {
    param([bool]$IsAdmin)
    $s = New-FmkSysSection -Key 'Battery' -Title 'Battery and Power'
    $batt = Read-FmkCim 'Win32_Battery'
    if (-not $batt -or $batt.Count -eq 0) {
        Add-FmkSysRow $s 'Battery' 'Not applicable (no battery detected)'
        return $s
    }
    $b = $batt[0]
    Add-FmkSysRow $s 'Battery' (Get-FmkSysField $b @('Name'))
    if ($batt.Count -gt 1) { Add-FmkSysRow $s 'Battery count' $batt.Count }
    $chg = Get-FmkSysField $b @('EstimatedChargeRemaining') -Missing $null
    if ($null -eq $chg) { Add-FmkSysRow $s 'Charge' 'Unavailable' }
    else { Add-FmkSysRow $s 'Charge' "$chg%" }
    $wmiStat = $null
    try { $wmiStat = Read-FmkCim 'BatteryStatus' -Namespace 'root/wmi' } catch { $wmiStat = $null }
    if ($wmiStat -and $wmiStat.Count -gt 0) {
        $charging = $false
        $discharging = $false
        $onAc = $false
        try { $charging = [bool]$wmiStat[0].Charging } catch { }
        try { $discharging = [bool]$wmiStat[0].Discharging } catch { }
        try { $onAc = [bool]$wmiStat[0].PowerOnline } catch { }
        if ($charging) { Add-FmkSysRow $s 'Status' 'Charging' }
        elseif ($discharging) { Add-FmkSysRow $s 'Status' 'Discharging' }
        elseif ($onAc) { Add-FmkSysRow $s 'Status' 'On AC power' }
        else { Add-FmkSysRow $s 'Status' 'Unknown' }
        if ($onAc) { Add-FmkSysRow $s 'AC power' 'Connected' }
        else { Add-FmkSysRow $s 'AC power' 'Not connected' }
    }
    else {
        Add-FmkSysRow $s 'Status' 'Unavailable'
        Add-FmkSysWarning $s 'Live battery status not exposed on this system.'
    }
    $designed = $null
    $full = $null
    try {
        $static = Read-FmkCim 'BatteryStaticData' -Namespace 'root/wmi'
        if ($static -and $static.Count -gt 0) { $designed = $static[0].DesignedCapacity }
    }
    catch { }
    try {
        $fcc = Read-FmkCim 'BatteryFullChargedCapacity' -Namespace 'root/wmi'
        if ($fcc -and $fcc.Count -gt 0) { $full = $fcc[0].FullChargedCapacity }
    }
    catch { }
    # WMI capacities arrive in milliwatt-hours: show watt-hours, not bytes.
    if ($null -ne $designed -and [double]$designed -gt 0) {
        Add-FmkSysRow $s 'Design capacity' (('{0:N1} Wh' -f ([double]$designed / 1000)))
    }
    else { Add-FmkSysRow $s 'Design capacity' 'Unavailable' }
    if ($null -ne $full -and [double]$full -gt 0) {
        Add-FmkSysRow $s 'Full-charge capacity' (('{0:N1} Wh' -f ([double]$full / 1000)))
    }
    else { Add-FmkSysRow $s 'Full-charge capacity' 'Unavailable' }
    if ($null -ne $designed -and $null -ne $full -and [double]$designed -gt 0 -and [double]$full -gt 0) {
        $health = [int][Math]::Round([double]$full / [double]$designed * 100)
        if ($health -gt 100) { $health = 100 }
        Add-FmkSysRow $s 'Battery health estimate' "$health% of design capacity"
    }
    else {
        Add-FmkSysRow $s 'Battery health estimate' 'Unavailable (needs design and full-charge capacities)'
    }
    try {
        $cyc = Read-FmkCim 'BatteryCycleCount' -Namespace 'root/wmi'
        if ($cyc -and $cyc.Count -gt 0 -and $null -ne $cyc[0].CycleCount) {
            Add-FmkSysRow $s 'Cycle count' ([int]$cyc[0].CycleCount)
        }
        else { Add-FmkSysRow $s 'Cycle count' 'Unavailable' }
    }
    catch { Add-FmkSysRow $s 'Cycle count' 'Unavailable' }
    try {
        $plan = powercfg /getactivescheme 2>$null
        $planText = (@($plan) | ForEach-Object { "$_" }) -join ' '
        if ($planText -match '\(([^)]+)\)\s*$') { Add-FmkSysRow $s 'Power plan' $Matches[1].Trim() }
        elseif ($planText.Trim() -ne '') { Add-FmkSysRow $s 'Power plan' $planText.Trim() }
        else { Add-FmkSysRow $s 'Power plan' 'Unavailable' }
    }
    catch {
        Add-FmkSysRow $s 'Power plan' 'Unavailable'
        Add-FmkSysWarning $s 'Active power plan unreadable.'
    }
    return $s
}

function Get-FmkSysAudioDisplaySection {
    param([bool]$IsAdmin)
    $s = New-FmkSysSection -Key 'AudioDisplay' -Title 'Audio, Displays and Peripherals'
    $snd = Read-FmkCim 'Win32_SoundDevice'
    if ($snd -and $snd.Count -gt 0) {
        $shownSnd = 0
        foreach ($dev in $snd) {
            if ($shownSnd -ge 6) { break }
            $shownSnd++
            $nm = Get-FmkSysField $dev @('Name')
            $mf = Get-FmkSysField $dev @('Manufacturer') -Missing ''
            if ($mf -ne '') { Add-FmkSysRow $s "Audio device $shownSnd" "$nm ($mf)" }
            else { Add-FmkSysRow $s "Audio device $shownSnd" $nm }
        }
        if ($snd.Count -gt 6) { Add-FmkSysRow $s 'Audio devices' ("Showing 6 of $($snd.Count).") }
    }
    else {
        Add-FmkSysWarning $s 'No audio devices reported.'
        Add-FmkSysRow $s 'Audio devices' 'Unavailable'
    }
    $mons = $null
    try { $mons = Read-FmkCim 'WmiMonitorID' -Namespace 'root/wmi' } catch { $mons = $null }
    if ($mons -and $mons.Count -gt 0) {
        $shownMon = 0
        foreach ($mon in $mons) {
            if ($shownMon -ge 6) { break }
            $shownMon++
            $mfr = ConvertFrom-FmkMonitorString (Get-FmkSysField $mon @('ManufacturerName') -Missing @())
            $friendly = ConvertFrom-FmkMonitorString (Get-FmkSysField $mon @('UserFriendlyName') -Missing @())
            if ($friendly -eq '') { $friendly = 'Generic monitor' }
            $label = $friendly
            if ($mfr -ne '') { $label = "$mfr $friendly" }
            Add-FmkSysRow $s "Monitor $shownMon" $label
            $serial = ConvertFrom-FmkMonitorString (Get-FmkSysField $mon @('SerialNumberID') -Missing @())
            if ($serial -ne '' -and $serial -ne '0') {
                Add-FmkSysRow $s "Monitor $shownMon serial number" $serial -Sensitivity 'Private'
            }
        }
        if ($mons.Count -gt 6) { Add-FmkSysRow $s 'Monitors' ("Showing 6 of $($mons.Count).") }
    }
    else {
        $desk = Read-FmkCim 'Win32_DesktopMonitor'
        if ($desk -and $desk.Count -gt 0) {
            $shownDesk = 0
            foreach ($d in $desk) {
                if ($shownDesk -ge 6) { break }
                $shownDesk++
                Add-FmkSysRow $s "Monitor $shownDesk" (Get-FmkSysField $d @('Name'))
            }
            Add-FmkSysWarning $s 'Detailed monitor names unavailable; showing generic entries.'
        }
        else {
            Add-FmkSysWarning $s 'No monitors reported.'
            Add-FmkSysRow $s 'Monitors' 'Unavailable'
        }
    }
    $kbd = Read-FmkCim 'Win32_Keyboard'
    if ($kbd -and $kbd.Count -gt 0) {
        $shownK = 0
        foreach ($k in $kbd) {
            if ($shownK -ge 2) { break }
            $shownK++
            Add-FmkSysRow $s "Keyboard $shownK" (Get-FmkSysField $k @('Name'))
        }
    }
    $ptr = Read-FmkCim 'Win32_PointingDevice'
    if ($ptr -and $ptr.Count -gt 0) {
        $shownP = 0
        foreach ($p in $ptr) {
            if ($shownP -ge 2) { break }
            $shownP++
            Add-FmkSysRow $s "Pointing device $shownP" (Get-FmkSysField $p @('Name'))
        }
    }
    $cams = Read-FmkCim 'Win32_PnPEntity' -Filter "PNPClass='Image'"
    if ($cams -and $cams.Count -gt 0) {
        $shownC = 0
        foreach ($c in $cams) {
            if ($shownC -ge 4) { break }
            $shownC++
            Add-FmkSysRow $s "Camera $shownC" (Get-FmkSysField $c @('Name'))
        }
    }
    else {
        Add-FmkSysRow $s 'Cameras' 'None detected'
    }
    $usb = Read-FmkCim 'Win32_PnPEntity' -Filter "PNPClass='USB'"
    if ($usb -and $usb.Count -gt 0) {
        $shownU = 0
        foreach ($u in $usb) {
            if ($shownU -ge 8) { break }
            $shownU++
            Add-FmkSysRow $s "USB device $shownU" (Get-FmkSysField $u @('Name'))
        }
        if ($usb.Count -gt 8) { Add-FmkSysRow $s 'USB devices' ("Showing 8 of $($usb.Count). Generic HID devices are summarized, not listed individually.") }
    }
    else {
        Add-FmkSysRow $s 'USB devices' 'Unavailable'
        Add-FmkSysWarning $s 'USB device list unreadable.'
    }
    return $s
}

function Get-FmkSysRuntimeSection {
    param([bool]$IsAdmin)
    $s = New-FmkSysSection -Key 'Runtime' -Title 'Runtime and Virtualization'
    try {
        Add-FmkSysRow $s 'PowerShell version' ([string]$PSVersionTable.PSVersion)
        Add-FmkSysRow $s 'PowerShell edition' ([string]$PSVersionTable.PSEdition)
    }
    catch {
        Add-FmkSysRow $s 'PowerShell version' 'Unavailable'
        Add-FmkSysWarning $s 'PowerShell version table unreadable.'
    }
    try {
        if ([System.Environment]::Is64BitOperatingSystem) { Add-FmkSysRow $s 'OS architecture (process view)' '64-bit OS' }
        else { Add-FmkSysRow $s 'OS architecture (process view)' '32-bit OS' }
        if ([System.Environment]::Is64BitProcess) { Add-FmkSysRow $s 'FreeMixKit process' '64-bit' }
        else { Add-FmkSysRow $s 'FreeMixKit process' '32-bit' }
    }
    catch { Add-FmkSysWarning $s 'Process architecture unreadable.' }
    $dotnetOk = $false
    try {
        $dn = Get-Command dotnet -ErrorAction SilentlyContinue
        if ($dn -and ("$($dn.Source)" -notmatch 'WindowsApps')) { $dotnetOk = $true }
    }
    catch { $dotnetOk = $false }
    if ($dotnetOk) {
        try {
            $sdks = @(dotnet --list-sdks 2>$null)
            if ($sdks.Count -gt 0) {
                $show = @($sdks | Select-Object -First 8)
                foreach ($sdk in $show) { Add-FmkSysRow $s '.NET SDK' ([string]$sdk).Trim() }
                if ($sdks.Count -gt 8) { Add-FmkSysRow $s '.NET SDKs' ("Showing 8 of $($sdks.Count).") }
            }
            else { Add-FmkSysRow $s '.NET SDKs' 'None reported' }
        }
        catch { Add-FmkSysWarning $s '.NET SDK list unreadable.' }
        try {
            $rts = @(dotnet --list-runtimes 2>$null)
            if ($rts.Count -gt 0) {
                $showRt = @($rts | Select-Object -First 8)
                foreach ($rt in $showRt) { Add-FmkSysRow $s '.NET runtime' ([string]$rt).Trim() }
                if ($rts.Count -gt 8) { Add-FmkSysRow $s '.NET runtimes' ("Showing 8 of $($rts.Count).") }
            }
            else { Add-FmkSysRow $s '.NET runtimes' 'None reported' }
        }
        catch { Add-FmkSysWarning $s '.NET runtime list unreadable.' }
    }
    else {
        Add-FmkSysRow $s '.NET SDKs / runtimes' 'Unavailable (dotnet not installed; nothing was installed to check)'
    }
    try {
        $csq = Read-FmkCim 'Win32_ComputerSystem'
        $hyp = $null
        if ($csq -and $csq.Count -gt 0) {
            try { $hyp = [bool]$csq[0].HypervisorPresent } catch { $hyp = $null }
            $mm = (Get-FmkSysField $csq[0] @('Manufacturer') -Missing '') + ' ' + (Get-FmkSysField $csq[0] @('Model') -Missing '')
            if ($mm.Trim() -ne '') {
                if ($mm -match '(?i)virtual machine|vmware|virtualbox|qemu|xen|kvm|hyper-v|parallels|bhyve|Pacer') {
                    Add-FmkSysRow $s 'Virtual machine indicators' ("Likely virtualized (model string: $($mm.Trim()))")
                }
                else {
                    Add-FmkSysRow $s 'Virtual machine indicators' 'No virtual-machine indicators in system model'
                }
            }
            else { Add-FmkSysRow $s 'Virtual machine indicators' 'Unavailable' }
        }
        if ($null -eq $hyp) { Add-FmkSysRow $s 'Hypervisor present' 'Unavailable' }
        elseif ($hyp) { Add-FmkSysRow $s 'Hypervisor present' 'Yes' }
        else { Add-FmkSysRow $s 'Hypervisor present' 'No' }
    }
    catch { Add-FmkSysWarning $s 'Virtualization status unreadable.' }
    $wslOk = $false
    try {
        if (Get-Command wsl.exe -ErrorAction SilentlyContinue) { $wslOk = $true }
    }
    catch { $wslOk = $false }
    if ($wslOk) {
        try {
            $st = @(wsl.exe --status 2>$null)
            if ($st.Count -gt 0) {
                foreach ($line in @($st | Select-Object -First 6)) {
                    $clean = ConvertFrom-FmkWslLine $line
                    if ($clean -ne '') { Add-FmkSysRow $s 'WSL' $clean }
                }
            }
            else { Add-FmkSysRow $s 'WSL' 'No status reported' }
        }
        catch {
            Add-FmkSysRow $s 'WSL' 'Unavailable'
            Add-FmkSysWarning $s 'WSL status unreadable.'
        }
    }
    else {
        Add-FmkSysRow $s 'WSL' 'Not installed or unavailable'
    }
    return $s
}

function Get-FmkSystemReportData {
    <#
        One collection pass over every section. Each section is independent:
        a failing section degrades to a warning row instead of stopping the
        report. Runs in a background runspace from the UI; runs inline
        anywhere else. Nothing here needs administrator rights to attempt,
        and sections needing more access say so when denied.
    #>
    $isAdmin = $false
    try {
        $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch { }

    $sections = New-Object System.Collections.ArrayList
    $collectors = @(
        @{ Key = 'Windows'; Title = 'Windows and Operating System'; Fn = ${function:Get-FmkSysWindowsSection} },
        @{ Key = 'CPU'; Title = 'Processor (CPU)'; Fn = ${function:Get-FmkSysCpuSection} },
        @{ Key = 'Memory'; Title = 'Memory (RAM)'; Fn = ${function:Get-FmkSysMemorySection} },
        @{ Key = 'GPU'; Title = 'Graphics (GPU)'; Fn = ${function:Get-FmkSysGpuSection} },
        @{ Key = 'Storage'; Title = 'Storage'; Fn = ${function:Get-FmkSysStorageSection} },
        @{ Key = 'Firmware'; Title = 'Motherboard and Firmware'; Fn = ${function:Get-FmkSysFirmwareSection} },
        @{ Key = 'Network'; Title = 'Network Adapters'; Fn = ${function:Get-FmkSysNetworkSection} },
        @{ Key = 'Battery'; Title = 'Battery and Power'; Fn = ${function:Get-FmkSysBatterySection} },
        @{ Key = 'AudioDisplay'; Title = 'Audio, Displays and Peripherals'; Fn = ${function:Get-FmkSysAudioDisplaySection} },
        @{ Key = 'Runtime'; Title = 'Runtime and Virtualization'; Fn = ${function:Get-FmkSysRuntimeSection} }
    )
    foreach ($c in $collectors) {
        try {
            [void]$sections.Add((& $c.Fn -IsAdmin $isAdmin))
        }
        catch {
            $fallback = New-FmkSysSection -Key $c.Key -Title $c.Title
            Add-FmkSysRow $fallback 'Section' 'Unavailable'
            Add-FmkSysWarning $fallback ("Collector failed: $($_.Exception.Message)")
            [void]$sections.Add($fallback)
        }
    }

    $byKey = @{}
    foreach ($sec in $sections) { $byKey[$sec.Key] = $sec }
    $overview = New-FmkSysSection -Key 'Overview' -Title 'System overview'
    Add-FmkSysRow $overview 'Operating system' (Find-FmkSysRow $byKey['Windows'] 'Product name')
    $rel = Find-FmkSysRow $byKey['Windows'] 'Release'
    $bld = Find-FmkSysRow $byKey['Windows'] 'Build'
    if ($rel -ne '' -and $bld -ne '') { Add-FmkSysRow $overview 'Release / build' "$rel (build $bld)" }
    Add-FmkSysRow $overview 'Processor' (Find-FmkSysRow $byKey['CPU'] 'Model')
    $cores = Find-FmkSysRow $byKey['CPU'] 'Physical cores'
    $threads = Find-FmkSysRow $byKey['CPU'] 'Logical processors'
    if ($cores -ne '' -and $threads -ne '') { Add-FmkSysRow $overview 'Cores / threads' "$cores cores, $threads threads" }
    $installed = Find-FmkSysRow $byKey['Memory'] 'Installed physical memory'
    $inuse = Find-FmkSysRow $byKey['Memory'] 'In use'
    if ($installed -ne '' -and $inuse -ne '') { Add-FmkSysRow $overview 'Memory' "$installed installed, $inuse in use" }
    elseif ($installed -ne '') { Add-FmkSysRow $overview 'Memory' $installed }
    Add-FmkSysRow $overview 'Graphics' (Find-FmkSysRow $byKey['GPU'] 'Name')
    $diskCount = 0
    try {
        foreach ($row in @($byKey['Storage'].Rows)) {
            if ([string]$row.Label -match '^Disk \d+$') { $diskCount++ }
        }
    }
    catch { }
    if ($diskCount -gt 0) { Add-FmkSysRow $overview 'Storage' "$diskCount physical disk(s) (see Storage section)" }

    $all = New-Object System.Collections.ArrayList
    [void]$all.Add($overview)
    foreach ($sec in $sections) { [void]$all.Add($sec) }

    return @{
        GeneratedAt = (Get-Date)
        IsAdmin     = $isAdmin
        Sections    = @($all)
    }
}

# ==============================================================================
# SYSTEM REPORT -- text and JSON output
# ==============================================================================

function Get-FmkSysDisplayValue {
    <#
        What a row shows. Untagged values always show. Identifying values
        (computer/user names) show on screen but are withheld from exports
        unless included. Private values (serials, MACs, IPs) hide everywhere
        unless an export explicitly includes them.
    #>
    param($Row, [bool]$IncludeSensitive = $false, [bool]$ForExport = $false)
    $sens = ''
    try { $sens = [string]$Row.Sensitivity } catch { $sens = '' }
    if ($sens -eq '' -or $sens -eq 'None') { return [string]$Row.Value }
    if ($IncludeSensitive) { return [string]$Row.Value }
    if ($ForExport) { return '[withheld]' }
    if ($sens -eq 'Identifying') { return [string]$Row.Value }
    return '[hidden]'
}

function Format-FmkSysTextReport {
    # ForExport selects the export wording for withheld values ('[withheld]');
    # the on-screen and clipboard form says '[hidden]' instead.
    param($Report, [bool]$IncludeSensitive = $false, [bool]$ForExport = $false)
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add('FreeMixKit System Report')
    try { [void]$lines.Add("Report format version: 1 (FreeMixKit v$AppVersion)") } catch { }
    try {
        $when = ([datetime]$Report.GeneratedAt).ToString('yyyy-MM-dd HH:mm')
        [void]$lines.Add("Generated: $when")
    }
    catch { [void]$lines.Add('Generated: unknown') }
    $adminText = 'No'
    try { if ([bool]$Report.IsAdmin) { $adminText = 'Yes' } } catch { }
    [void]$lines.Add("Running as administrator: $adminText")
    if ($IncludeSensitive) { [void]$lines.Add('Sensitive identifiers: included') }
    else { [void]$lines.Add('Sensitive identifiers: withheld') }
    try {
        foreach ($sec in @($Report.Sections)) {
            [void]$lines.Add('')
            [void]$lines.Add("==== $($sec.Title) ====")
            foreach ($row in @($sec.Rows)) {
                [void]$lines.Add(("{0}: {1}" -f $row.Label, (Get-FmkSysDisplayValue $row -IncludeSensitive:$IncludeSensitive -ForExport:$ForExport)))
            }
            foreach ($warn in @($sec.Warnings)) {
                [void]$lines.Add("Note: $warn")
            }
        }
    }
    catch { [void]$lines.Add('(report body unreadable)') }
    return ($lines -join "`r`n")
}

function Write-FmkSysReportFile {
    <#
        Writes report text to a .txt file as UTF-8 WITH BOM. The BOM is the
        entire point: without it, BOM-less bytes are bare bytes, and any
        reader that assumes or sniffs UTF-16 decodes plain ASCII as CJK
        ("Free Mix Kit ..." becomes U+7246 U+6565 ..., i.e. Chinese-looking
        mojibake). The BOM makes the file self-describing for Notepad and
        every other Windows reader, while old and new PowerShell/.NET readers
        detect and strip it either way.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text
    )
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Text, $utf8Bom)
}

function ConvertTo-FmkSysJsonReport {
    <# Redacted plain-data copy of the report, ready for ConvertTo-Json. #>
    param($Report, [bool]$IncludeSensitive = $false)
    $secs = New-Object System.Collections.ArrayList
    try {
        foreach ($sec in @($Report.Sections)) {
            $rows = New-Object System.Collections.ArrayList
            foreach ($row in @($sec.Rows)) {
                [void]$rows.Add([ordered]@{
                        Label       = [string]$row.Label
                        Value       = (Get-FmkSysDisplayValue $row -IncludeSensitive:$IncludeSensitive -ForExport:$true)
                        Sensitivity = [string]$row.Sensitivity
                    })
            }
            [void]$secs.Add([ordered]@{
                    Key      = [string]$sec.Key
                    Title    = [string]$sec.Title
                    Rows     = @($rows)
                    Warnings = @($sec.Warnings | ForEach-Object { [string]$_ })
                })
        }
    }
    catch { }
    $stamp = ''
    try { $stamp = ([datetime]$Report.GeneratedAt).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') } catch { }
    $admin = $false
    try { $admin = [bool]$Report.IsAdmin } catch { }
    return [ordered]@{
        GeneratedAt       = $stamp
        IsAdmin           = $admin
        SensitiveIncluded = [bool]$IncludeSensitive
        Sections          = @($secs)
    }
}
