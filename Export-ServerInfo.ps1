#Requires -Version 5.1
<#
.SYNOPSIS
    Collects server hardware and print-management software information
    for Solutions Consultant system records.
.DESCRIPTION
    READ-ONLY. Queries WMI, registry Uninstall keys, Windows services,
    and known install directories. No system changes made.
.OUTPUTS
    ServerInfo_HOSTNAME_DATE.txt  - drop into ConsultantApp PMS folder
    ExportLog_HOSTNAME_DATE.txt   - full audit trail
#>

$ErrorActionPreference = 'SilentlyContinue'

# --- Output paths ---
$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$stamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
$hostname   = $env:COMPUTERNAME
$resultsDir = Join-Path $scriptDir 'results'
$logsDir    = Join-Path $scriptDir 'logs'
if (-not (Test-Path $resultsDir)) { New-Item -ItemType Directory -Path $resultsDir | Out-Null }
if (-not (Test-Path $logsDir))    { New-Item -ItemType Directory -Path $logsDir    | Out-Null }
$outFile    = Join-Path $resultsDir "ServerInfo_${hostname}_${stamp}.txt"
$logFile    = Join-Path $resultsDir "ExportLog_${hostname}_${stamp}.txt"
$traceFile  = Join-Path $logsDir    "TraceLog_${hostname}_${stamp}.txt"

# --- Helpers ---
function Write-Log([string]$msg) {
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "[$ts] $msg" | Out-File $logFile -Append -Encoding UTF8
}

function Write-Trace([string]$msg) {
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "[$ts] $msg" | Out-File $traceFile -Append -Encoding UTF8
    Write-Host "  >> $msg"
}

function Write-Out([string]$line = '') {
    $line | Out-File $outFile -Append -Encoding UTF8
}

function Write-Field([string]$label, [string]$value) {
    "${label}: ${value}" | Out-File $outFile -Append -Encoding UTF8
}

function Write-Section([string]$title) {
    Write-Out ''
    Write-Out "[$title]"
    Write-Log "--- Section: $title ---"
    Write-Trace "Starting section: $title"
}

# --- Audit log header ---
$auditHeader = @"
============================================================
 EXECUTION AUDIT LOG
 Script  : $($MyInvocation.MyCommand.Path)
 Host    : $hostname
 User    : $($env:USERNAME)
 Domain  : $($env:USERDOMAIN)
 Started : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
============================================================

 PURPOSE: Records every action for security and compliance review.

 WHAT THIS SCRIPT DOES:
   - Reads OS, hardware, network info via WMI (read-only)
   - Reads installed software from registry Uninstall keys (read-only)
   - Checks Windows service list (read-only)
   - Checks known install directories for print management apps
   - Reads version/config files inside those directories (read-only)

 WHAT THIS SCRIPT DOES NOT DO:
   - Does not modify any system settings or files
   - Does not write to the registry
   - Does not start, stop, or modify any services
   - Does not send data over the network
   - Does not create files outside this script folder (outputs go to results\ subfolder)

============================================================

"@
$auditHeader | Out-File $logFile -Encoding UTF8

# --- Trace log header ---
@"
TraceLog -- Export-ServerInfo.ps1
Host    : $hostname
User    : $($env:USERNAME)
Started : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Results : $resultsDir
"@ | Out-File $traceFile -Encoding UTF8
Write-Trace 'Script started'

# --- Data file header ---
$dataHeader = @"
============================================================
 Server Information Export
 Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
 Computer  : $hostname
 User      : $($env:USERNAME)
============================================================
"@
$dataHeader | Out-File $outFile -Encoding UTF8


# ===========================================================
# SECTION 1 - System Identity (Physical vs VM, UUID)
# ===========================================================
Write-Section 'System Identity'
Write-Log 'QUERY: Get-CimInstance Win32_ComputerSystem (manufacturer, model)'
Write-Log 'QUERY: Get-CimInstance Win32_ComputerSystemProduct (UUID)'
Write-Log 'QUERY: Get-CimInstance Win32_BIOS (serial, version)'

$cs      = Get-CimInstance Win32_ComputerSystem
$csp     = Get-CimInstance Win32_ComputerSystemProduct
$bios    = Get-CimInstance Win32_BIOS

$mfr     = $cs.Manufacturer
$model   = $cs.Model
$uuid    = $csp.UUID
$biosSerial = $bios.SerialNumber
$biosVer    = $bios.SMBIOSBIOSVersion

# --- Detect VM platform ---
$platform = 'Physical Server'
$hypervisor = ''

$vmChecks = @(
    @{ Pattern = 'VMware';          Name = 'VMware';       Fields = @($mfr, $model, $biosVer) },
    @{ Pattern = 'Virtual Machine'; Name = 'Hyper-V';      Fields = @($model) },
    @{ Pattern = 'Microsoft.*Hyper';Name = 'Hyper-V';      Fields = @($mfr, $model) },
    @{ Pattern = 'VirtualBox';      Name = 'VirtualBox';   Fields = @($mfr, $model) },
    @{ Pattern = 'innotek';         Name = 'VirtualBox';   Fields = @($mfr) },
    @{ Pattern = 'QEMU';            Name = 'KVM/QEMU';     Fields = @($mfr, $model, $biosVer) },
    @{ Pattern = 'KVM';             Name = 'KVM/QEMU';     Fields = @($model) },
    @{ Pattern = 'Xen';             Name = 'Xen';          Fields = @($mfr, $model, $biosVer) },
    @{ Pattern = 'Bochs';           Name = 'Bochs/KVM';    Fields = @($mfr, $biosVer) },
    @{ Pattern = 'Parallels';       Name = 'Parallels';    Fields = @($mfr, $model) },
    @{ Pattern = 'Google';          Name = 'Google Cloud'; Fields = @($mfr, $model) },
    @{ Pattern = 'Amazon EC2';      Name = 'AWS EC2';      Fields = @($mfr, $model) }
)

foreach ($check in $vmChecks) {
    foreach ($field in $check.Fields) {
        if ($field -match $check.Pattern) {
            $platform   = 'Virtual Machine'
            $hypervisor = $check.Name
            break
        }
    }
    if ($hypervisor -ne '') { break }
}

# Secondary check: CPUID hypervisor bit via registry (Hyper-V guest services key)
if ($hypervisor -eq '') {
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters') {
        $platform   = 'Virtual Machine'
        $hypervisor = 'Hyper-V'
    }
}

Write-Field 'Computer Name'   $cs.Name
Write-Field 'Manufacturer'    $mfr
Write-Field 'Model'           $model
Write-Field 'Platform'        $platform
if ($hypervisor -ne '') {
    Write-Field 'Hypervisor'  $hypervisor
}
Write-Field 'System UUID'     $uuid
Write-Field 'BIOS Version'    $biosVer
Write-Field 'BIOS Serial'     $biosSerial
Write-Log "RESULT: Platform=$platform Hypervisor=$hypervisor UUID=$uuid Manufacturer=$mfr Model=$model"


# ===========================================================
# SECTION 2 - Operating System
# ===========================================================
Write-Section 'Operating System'
Write-Log 'QUERY: Get-CimInstance Win32_OperatingSystem'
$os = Get-CimInstance Win32_OperatingSystem
Write-Field 'OS Type'        $os.Caption
Write-Field 'OS Version'     $os.Version
Write-Field 'Server Version' $os.Caption
Write-Field 'Architecture'   $os.OSArchitecture
Write-Field 'Install Date'   ($os.InstallDate.ToString('yyyy-MM-dd'))
Write-Log "RESULT: $($os.Caption) $($os.Version)"


# ===========================================================
# SECTION 3 - Hardware
# ===========================================================
Write-Section 'Hardware'
Write-Log 'QUERY: Get-CimInstance Win32_ComputerSystem (RAM)'
Write-Log 'QUERY: Get-CimInstance Win32_Processor (CPU)'
# $cs already loaded in Section 1
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$ram = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
Write-Field 'RAM'             "$ram GB"
Write-Field 'Memory'          "$ram GB"
Write-Field 'Physical Memory' "$ram GB"
Write-Field 'CPU'             $cpu.Name
Write-Field 'Processor'       $cpu.Name
Write-Field 'CPU Cores'       ([string]$cpu.NumberOfCores)
Write-Field 'CPU Threads'     ([string]$cpu.NumberOfLogicalProcessors)
Write-Log "RESULT: RAM=$ram GB  CPU=$($cpu.Name)"


# ===========================================================
# SECTION 3 - Storage
# ===========================================================
Write-Section 'Storage'
Write-Log 'QUERY: Get-PSDrive -PSProvider FileSystem'
$drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -ne $null }
foreach ($d in $drives) {
    $total = [math]::Round(($d.Used + $d.Free) / 1GB, 1)
    $used  = [math]::Round($d.Used / 1GB, 1)
    $free  = [math]::Round($d.Free / 1GB, 1)
    Write-Field "Storage $($d.Name)" "$total GB total, $used GB used, $free GB free"
    Write-Field "HDD $($d.Name)"     "$total GB"
    Write-Log "RESULT: Drive $($d.Name) = $total GB total"
}


# ===========================================================
# SECTION 4 - Network
# ===========================================================
Write-Section 'Network'
Write-Log 'QUERY: Get-CimInstance Win32_NetworkAdapterConfiguration (all adapters, enabled)'
Write-Log 'QUERY: Get-CimInstance Win32_NetworkAdapter (for connection status and type)'

$netbiosMap = @{ 0 = 'Default (via DHCP)'; 1 = 'Enabled'; 2 = 'Disabled' }

# Get all enabled adapter configs (not just those with IPs - includes disconnected NICs)
$adapterConfigs = Get-CimInstance Win32_NetworkAdapterConfiguration |
                  Where-Object { $_.IPEnabled -eq $true }

# Also get adapter objects for connection status
$adapterObjects = Get-CimInstance Win32_NetworkAdapter |
                  Where-Object { $_.NetEnabled -ne $null }

foreach ($a in $adapterConfigs) {
    # Match to physical adapter for extra info
    $phys = $adapterObjects | Where-Object { $_.InterfaceIndex -eq $a.InterfaceIndex } | Select-Object -First 1

    Write-Out "  NIC: $($a.Description)"
    Write-Out "  MAC Address         : $($a.MACAddress)"

    if ($phys) {
        $connStatus = if ($phys.NetConnectionStatus -eq 2) { 'Connected' } elseif ($phys.NetConnectionStatus -eq 7) { 'Media Disconnected' } else { "Status code $($phys.NetConnectionStatus)" }
        Write-Out "  Connection Status   : $connStatus"
        if ($phys.Speed) {
            $speedMbps = [math]::Round($phys.Speed / 1MB)
            Write-Out "  Link Speed          : $speedMbps Mbps"
        }
    }

    # IP addresses and subnet masks (can be multiple)
    if ($a.IPAddress) {
        for ($i = 0; $i -lt $a.IPAddress.Count; $i++) {
            $ip     = $a.IPAddress[$i]
            $subnet = if ($a.IPSubnet -and $i -lt $a.IPSubnet.Count) { $a.IPSubnet[$i] } else { '' }
            if ($ip -match ':') {
                Write-Out "  IPv6 Address        : $ip"
            } else {
                Write-Out "  IP Address          : $ip"
                if ($subnet) { Write-Out "  Subnet Mask         : $subnet" }
            }
        }
    } else {
        Write-Out "  IP Address          : Not assigned"
    }

    # Default gateway
    if ($a.DefaultIPGateway) {
        Write-Out "  Default Gateway     : $($a.DefaultIPGateway -join ', ')"
    } else {
        Write-Out "  Default Gateway     : None"
    }

    # DNS servers
    if ($a.DNSServerSearchOrder) {
        $dnsServers = $a.DNSServerSearchOrder
        Write-Out "  DNS Server 1        : $($dnsServers[0])"
        for ($d = 1; $d -lt $dnsServers.Count; $d++) {
            Write-Out "  DNS Server $($d + 1)        : $($dnsServers[$d])"
        }
    } else {
        Write-Out "  DNS Servers         : None configured"
    }

    # DNS domain / suffix
    if ($a.DNSDomain)          { Write-Out "  DNS Domain          : $($a.DNSDomain)" }
    if ($a.DNSDomainSuffixSearchOrder -and $a.DNSDomainSuffixSearchOrder.Count -gt 0) {
        Write-Out "  DNS Search Suffix   : $($a.DNSDomainSuffixSearchOrder -join ', ')"
    }

    # DHCP
    $dhcpStatus = if ($a.DHCPEnabled) { "Enabled (server: $($a.DHCPServer))" } else { 'Disabled (static)' }
    Write-Out "  DHCP                : $dhcpStatus"

    # WINS
    if ($a.WINSPrimaryServer)   { Write-Out "  WINS Primary        : $($a.WINSPrimaryServer)" }
    if ($a.WINSSecondaryServer) { Write-Out "  WINS Secondary      : $($a.WINSSecondaryServer)" }

    # LMHOSTS lookup
    $lmhosts = if ($a.WINSEnableLMHostsLookup) { 'Enabled' } else { 'Disabled' }
    Write-Out "  LMHOSTS Lookup      : $lmhosts"

    # NetBIOS over TCP/IP
    $netbiosVal  = $a.TcpipNetbiosOptions
    $netbiosText = if ($netbiosMap.ContainsKey([int]$netbiosVal)) { $netbiosMap[[int]$netbiosVal] } else { "Unknown ($netbiosVal)" }
    Write-Out "  NetBIOS over TCP/IP : $netbiosText"

    Write-Out '  ---'
    Write-Log "RESULT: NIC=$($a.Description) IP=$($a.IPAddress -join ',') GW=$($a.DefaultIPGateway -join ',') NetBIOS=$netbiosText LMHOSTS=$lmhosts"
}


# ===========================================================
# SECTION 5 - Domain
# ===========================================================
Write-Section 'Domain'
Write-Log 'QUERY: Get-CimInstance Win32_ComputerSystem (domain membership)'
# $cs already loaded in Section 1
Write-Field 'Domain' $cs.Domain
if ($cs.PartOfDomain) {
    Write-Field 'Domain Joined'    "Yes - $($cs.Domain)"
    Write-Field 'Active Directory' "Yes - $($cs.Domain)"
    Write-Log "RESULT: Domain joined = Yes ($($cs.Domain))"
} else {
    Write-Field 'Domain Joined'    "No (Workgroup: $($cs.Workgroup))"
    Write-Field 'Active Directory' 'No'
    Write-Log 'RESULT: Domain joined = No'
}


# ===========================================================
# SECTION 6 - Installed Print Management Software
# ===========================================================
Write-Section 'Installed Software'
Write-Log 'QUERY: Registry Uninstall keys (read-only)'
Write-Log 'QUERY: Get-Service (display name filter)'
Write-Log 'QUERY: Test-Path on known install directories'

# Load all uninstall entries and services once
$uninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
$allInstalled = Get-ItemProperty $uninstallPaths -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName }
$allServices  = Get-Service -ErrorAction SilentlyContinue


function Find-PrintApp {
    param(
        [string]   $AppLabel,
        [string[]] $NamePatterns,
        [string[]] $ServicePatterns,
        [string[]] $InstallDirs,
        [string[]] $VersionFiles
    )

    $found      = $false
    $version    = ''
    $installDir = ''
    $svcAccount = ''
    $svcStatus  = ''

    # 1 - Registry Uninstall
    foreach ($pat in $NamePatterns) {
        $reg = $allInstalled | Where-Object { $_.DisplayName -like $pat } | Select-Object -First 1
        if ($reg) {
            $found      = $true
            $version    = if ($reg.DisplayVersion) { $reg.DisplayVersion } else { 'Found (version unknown)' }
            $installDir = if ($reg.InstallLocation) { $reg.InstallLocation } else { '' }
            Write-Log "RESULT: $AppLabel found in registry: $($reg.DisplayName) v$version"
            break
        }
    }

    # 2 - Windows Services
    foreach ($pat in $ServicePatterns) {
        $svc = $allServices | Where-Object { $_.DisplayName -like $pat } | Select-Object -First 1
        if ($svc) {
            $found     = $true
            $svcStatus = [string]$svc.Status
            $wmiSvc    = Get-CimInstance Win32_Service | Where-Object { $_.Name -eq $svc.Name }
            $svcAccount = if ($wmiSvc -and $wmiSvc.StartName) { $wmiSvc.StartName } else { 'LocalSystem' }
            Write-Log "RESULT: $AppLabel service: $($svc.DisplayName) [$svcStatus] account=$svcAccount"
            break
        }
    }

    # 3 - Install directories
    foreach ($dir in $InstallDirs) {
        if (Test-Path $dir) {
            $found = $true
            if ($installDir -eq '') { $installDir = $dir }
            Write-Log "RESULT: $AppLabel directory found: $dir"
            foreach ($rel in $VersionFiles) {
                $vf = Join-Path $dir $rel
                if (Test-Path $vf) {
                    Write-Log "QUERY: Reading $vf"
                    $lines = Get-Content $vf -ErrorAction SilentlyContinue | Select-Object -First 10
                    if ($lines) {
                        $vLine = $lines | Where-Object { $_ -match 'version|build|release' } | Select-Object -First 1
                        if ($vLine) { $version = $vLine.Trim() } else { $version = $lines[0].Trim() }
                        Write-Log "RESULT: Version from file: $version"
                    }
                    break
                }
            }
            break
        }
    }

    # Output
    if ($found) {
        Write-Field 'Solution' $AppLabel
        Write-Field 'Product'  $AppLabel
        if ($version    -ne '') { Write-Field "$AppLabel Version" $version }
        if ($installDir -ne '') { Write-Field "$AppLabel Path"    $installDir }
        if ($svcStatus  -ne '') { Write-Field 'Service Status'    $svcStatus }
        if ($svcAccount -ne '') {
            Write-Field 'Service Account' $svcAccount
            Write-Field 'Account Type'    $svcAccount
        }
    } else {
        Write-Log "RESULT: $AppLabel - not detected"
    }
    return $found
}


# --- PaperCut MF / NG ---
Write-Log 'CHECK: PaperCut MF/NG'
$pcFound = Find-PrintApp `
    -AppLabel        'PaperCut MF' `
    -NamePatterns    @('*PaperCut MF*', '*PaperCut NG*') `
    -ServicePatterns @('*PaperCut*') `
    -InstallDirs     @(
        'C:\Program Files\PaperCut MF',
        'C:\Program Files (x86)\PaperCut MF',
        'C:\Program Files\PaperCut NG',
        'C:\Program Files (x86)\PaperCut NG'
    ) `
    -VersionFiles    @(
        'server\version.txt',
        'CHANGELOG.txt',
        'server\lib-ext\version.properties'
    )

if ($pcFound) {
    $pcConfigs = @(
        'C:\Program Files\PaperCut MF\server\data\conf\server.properties',
        'C:\Program Files (x86)\PaperCut MF\server\data\conf\server.properties',
        'C:\Program Files\PaperCut NG\server\data\conf\server.properties'
    )
    foreach ($cfg in $pcConfigs) {
        if (Test-Path $cfg) {
            Write-Log "QUERY: PaperCut server.properties (user.sync.source)"
            $syncLine = Get-Content $cfg | Where-Object { $_ -match '^user\.sync\.source' } | Select-Object -First 1
            if ($syncLine) {
                $syncVal = ($syncLine -split '=')[1].Trim()
                Write-Field 'AD Sync'          $syncVal
                Write-Field 'Active Directory' $syncVal
                if ($syncVal -eq 'custom') { Write-Field 'CSV Import' 'Yes' } else { Write-Field 'CSV Import' 'No' }
                Write-Log "RESULT: PaperCut sync = $syncVal"
            }
            break
        }
    }
}
Write-Out ''


# --- Equitrac / ControlSuite / Nuance / Kofax / Tungsten ---
Write-Log 'CHECK: Equitrac / ControlSuite / Nuance / Kofax / Tungsten'
$csFound = Find-PrintApp `
    -AppLabel        'Equitrac / ControlSuite' `
    -NamePatterns    @('*Equitrac*', '*ControlSuite*', '*Nuance Control*', '*Kofax Control*', '*Tungsten*') `
    -ServicePatterns @('*Equitrac*', '*ControlSuite*', '*Nuance*', '*Kofax*', '*Tungsten*') `
    -InstallDirs     @(
        'C:\Program Files\Nuance\ControlSuite',
        'C:\Program Files\Kofax\ControlSuite',
        'C:\Program Files\Equitrac',
        'C:\Program Files (x86)\Equitrac',
        'C:\Program Files\Nuance',
        'C:\Program Files\Kofax',
        'C:\Program Files\Tungsten'
    ) `
    -VersionFiles    @('version.txt', 'build.txt', 'release.txt')
Write-Out ''


# --- YSoft SafeQ ---
Write-Log 'CHECK: YSoft SafeQ'
Find-PrintApp `
    -AppLabel        'YSoft SafeQ' `
    -NamePatterns    @('*SafeQ*', '*YSoft*') `
    -ServicePatterns @('*SafeQ*', '*YSoft*') `
    -InstallDirs     @(
        'C:\SafeQ6',
        'C:\SafeQ5',
        'C:\Program Files\Y Soft',
        'C:\Program Files\YSoft'
    ) `
    -VersionFiles    @('version.txt', 'build\version.txt', 'server\version.txt') | Out-Null
Write-Out ''


# --- AWMS2 (FujiFilm Business Innovation) ---
Write-Log 'CHECK: AWMS2'
Find-PrintApp `
    -AppLabel        'AWMS2' `
    -NamePatterns    @('*AWMS*', '*ApeosWare Management*', '*Account and Workplace*') `
    -ServicePatterns @('*AWMS*', '*ApeosWare*') `
    -InstallDirs     @(
        'C:\Program Files\FujiFilm\AWMS',
        'C:\Program Files\Fuji Xerox\AWMS',
        'C:\Program Files\FUJIFILM BI\AWMS',
        'C:\AWMS2',
        'C:\Program Files\FujiFilm Business Innovation\AWMS'
    ) `
    -VersionFiles    @('version.txt', 'build.properties', 'conf\version.txt') | Out-Null
Write-Out ''


# --- Catch-all: other print management apps ---
Write-Log 'CHECK: Other print management software (broad registry scan)'
$otherKeywords = @('PrinterLogic','Printix','Pharos','UniFlow','ThinPrint','easyPRINT','MyQ','PrintManager')
foreach ($kw in $otherKeywords) {
    $match = $allInstalled | Where-Object { $_.DisplayName -like "*$kw*" } | Select-Object -First 1
    if ($match) {
        Write-Field 'Solution'                      $match.DisplayName
        Write-Field "$($match.DisplayName) Version" $match.DisplayVersion
        Write-Log "RESULT: Found $($match.DisplayName) v$($match.DisplayVersion)"
    }
}


# ===========================================================
# SECTION 6b - Print Management Solution Deep-Dive
# (Equitrac / ControlSuite / Nuance / Kofax / Tungsten Automation)
# ===========================================================
Write-Section 'Print Management Solution'

# --- Helper: dump a registry key tree to the output file ---
function Write-RegTree {
    param([string]$Path, [int]$Depth = 0, [int]$MaxDepth = 5)
    if (-not (Test-Path -LiteralPath $Path -EA SilentlyContinue)) { return }
    if ($Depth -gt $MaxDepth) { return }
    $pad = '  ' * ($Depth + 2)
    Write-Out "${pad}[$Path]"
    try {
        $props = Get-ItemProperty -LiteralPath $Path -EA SilentlyContinue
        if ($props) {
            $props.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object {
                $n = $_.Name
                $v = $_.Value
                if ($v -is [byte[]]) {
                    if ($v.Count -le 128) {
                        $hex = ($v | ForEach-Object { '{0:X2}' -f $_ }) -join ' '
                        Write-Out "${pad}  $n = [BIN] $hex"
                    } else {
                        $str = [System.Text.Encoding]::Unicode.GetString($v).TrimEnd([char]0) -replace '[\x00-\x08\x0B\x0E-\x1F]',''
                        if ($str -match '^[\x09\x0A\x0D\x20-\x7E]+$') { Write-Out "${pad}  $n = $str" }
                        else { Write-Out "${pad}  $n = [BIN $($v.Count)B] $(($v[0..31] | ForEach-Object { '{0:X2}' -f $_ }) -join ' ') ..." }
                    }
                } elseif ($v -is [string[]]) {
                    Write-Out "${pad}  $n = $($v -join ' | ')"
                } else {
                    Write-Out "${pad}  $n = $v"
                }
            }
        }
    } catch {}
    try {
        Get-ChildItem -LiteralPath $Path -EA SilentlyContinue | ForEach-Object {
            Write-RegTree -Path $_.PSPath -Depth ($Depth + 1) -MaxDepth $MaxDepth
        }
    } catch {}
}

# --- Helper: run SQL via Windows integrated auth, return DataTable ---
function Invoke-SQLWin {
    param([string]$Server, [string]$Database = 'master', [string]$Query, [int]$TimeoutSec = 30)
    try {
        $cs  = "Server=$Server;Database=$Database;Integrated Security=True;Connection Timeout=5;TrustServerCertificate=True;"
        $con = New-Object System.Data.SqlClient.SqlConnection $cs
        $con.Open()
        $cmd = $con.CreateCommand()
        $cmd.CommandText = $Query; $cmd.CommandTimeout = $TimeoutSec
        $da  = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $dt  = New-Object System.Data.DataTable
        $da.Fill($dt) | Out-Null
        $con.Close()
        return ,$dt
    } catch { return $null }
}

# --- Helper: print a DataTable aligned to Write-Out ---
function Write-SqlTable {
    param($dt, [string]$Pad = '    ', [int]$MaxRows = 500, [int]$MaxCol = 80)
    if ($dt -is [array]) { $dt = $dt[0] }
    if ($null -eq $dt -or -not ($dt -is [System.Data.DataTable]) -or $dt.Rows.Count -eq 0) {
        Write-Out "${Pad}(no rows)"
        return
    }
    $w = @{}
    foreach ($col in $dt.Columns) { $w[$col.ColumnName] = $col.ColumnName.Length }
    $rows = if ($dt.Rows.Count -gt $MaxRows) { $dt.Rows | Select-Object -First $MaxRows } else { $dt.Rows }
    foreach ($row in $rows) {
        foreach ($col in $dt.Columns) {
            $len = [Math]::Min("$($row[$col.ColumnName])".Length, $MaxCol)
            if ($len -gt $w[$col.ColumnName]) { $w[$col.ColumnName] = $len }
        }
    }
    $hdr = ($dt.Columns | ForEach-Object { $_.ColumnName.PadRight($w[$_.ColumnName]) }) -join '  '
    $sep = ($dt.Columns | ForEach-Object { '-' * $w[$_.ColumnName] }) -join '  '
    Write-Out ($Pad + $hdr)
    Write-Out ($Pad + $sep)
    foreach ($row in $rows) {
        $line = ($dt.Columns | ForEach-Object {
            $v = "$($row[$_.ColumnName])"
            if ($v.Length -gt $MaxCol) { $v = $v.Substring(0, $MaxCol - 3) + '...' }
            $v.PadRight($w[$_.ColumnName])
        }) -join '  '
        Write-Out ($Pad + $line)
    }
    if ($dt.Rows.Count -gt $MaxRows) { Write-Out ($Pad + "... ($($dt.Rows.Count - $MaxRows) more rows omitted)") }
}

# Detect which brand is present (check registry, services, install dirs)
Write-Log 'QUERY: Brand detection -- registry, services, install directories'
$brandMap = @(
    @{ Label='Tungsten Automation / ControlSuite'; Patterns=@('Tungsten'); Dirs=@('C:\Program Files\Tungsten','C:\Program Files\Tungsten Automation') },
    @{ Label='Kofax ControlSuite';  Patterns=@('Kofax');   Dirs=@('C:\Program Files\Kofax','C:\Program Files (x86)\Kofax','C:\ProgramData\Kofax') },
    @{ Label='Nuance ControlSuite'; Patterns=@('Nuance');  Dirs=@('C:\Program Files\Nuance','C:\Program Files (x86)\Nuance','C:\ProgramData\Nuance') },
    @{ Label='Equitrac';            Patterns=@('Equitrac');Dirs=@('C:\Program Files\Equitrac','C:\Program Files (x86)\Equitrac','C:\ProgramData\Equitrac') }
)
$detectedBrands = [System.Collections.Generic.List[string]]::new()
foreach ($b in $brandMap) {
    foreach ($pat in $b.Patterns) {
        if ((Test-Path "HKLM:\SOFTWARE\$pat" -EA SilentlyContinue) -or
            (Test-Path "HKLM:\SOFTWARE\WOW6432Node\$pat" -EA SilentlyContinue) -or
            ($b.Dirs | Where-Object { Test-Path $_ -EA SilentlyContinue }) -or
            ($allInstalled | Where-Object { $_.DisplayName -match "(?i)$pat" })) {
            $detectedBrands.Add($b.Label)
            Write-Log "RESULT: Brand detected: $($b.Label)"
            break
        }
    }
}
# Also catch via services
$eqSvcs = $allServices | Where-Object { $_.Name -match '(?i)^EQ[A-Z]|Equitrac|ControlSuite|eqcas|eqlog|eqpms|eqaud|eqdir|eqextend' }
if ($eqSvcs -and $detectedBrands.Count -eq 0) { $detectedBrands.Add('Equitrac / ControlSuite (service-detected)') }

if ($detectedBrands.Count -eq 0) {
    Write-Out '  No Equitrac / ControlSuite / Nuance / Kofax / Tungsten installation detected.'
    Write-Log 'RESULT: No print management solution detected'
} else {
    Write-Out "  Detected: $($detectedBrands -join ', ')"
    Write-Out ''

    # ===========================================================
    # 6b.1 - Installed Products
    # ===========================================================
    Write-Out '[Installed Products]'
    Write-Log 'QUERY: Uninstall registry -- brand products'
    $brandProducts = $allInstalled | Where-Object {
        $_.DisplayName -match '(?i)Equitrac|ControlSuite|Nuance|Kofax|Tungsten|Omtool|Autonomy'
    }
    if ($brandProducts) {
        foreach ($p in $brandProducts | Sort-Object DisplayName) {
            Write-Out "  Product   : $($p.DisplayName)"
            Write-Out "  Version   : $($p.DisplayVersion)"
            Write-Out "  Publisher : $($p.Publisher)"
            Write-Out "  InstDate  : $($p.InstallDate)"
            Write-Out "  Location  : $($p.InstallLocation)"
            Write-Out "  GUID      : $($p.PSChildName)"
            Write-Out ''
            Write-Log "RESULT: Installed: $($p.DisplayName) v$($p.DisplayVersion)"
        }
    } else {
        Write-Out '  (none in uninstall registry)'
    }

    # ===========================================================
    # 6b.2 - Services
    # ===========================================================
    Write-Out '[Services]'
    Write-Log 'QUERY: Get-Service / Win32_Service -- brand services'
    $brandSvcList = $allServices | Where-Object {
        $_.Name        -match '(?i)^EQ[A-Z]|Equitrac|ControlSuite|eqcas|eqlog|eqpms|eqaud|eqdir|eqextend|Tungsten' -or
        $_.DisplayName -match '(?i)Equitrac|ControlSuite|Nuance|Kofax|Tungsten'
    }
    if ($brandSvcList) {
        foreach ($svc in $brandSvcList | Sort-Object Name) {
            $ws = Get-CimInstance Win32_Service -Filter "Name='$($svc.Name)'" -EA SilentlyContinue
            Write-Out "  $($svc.Name.PadRight(35)) Status=$($svc.Status)  Start=$($svc.StartType)  Account=$($ws.StartName)"
            Write-Out "    Display   : $($svc.DisplayName)"
            if ($ws -and $ws.PathName) { Write-Out "    Executable: $($ws.PathName)" }
            Write-Out ''
            Write-Log "RESULT: Service $($svc.Name) status=$($svc.Status) account=$($ws.StartName)"
        }
    } else {
        Write-Out '  (no brand services found)'
    }

    # ===========================================================
    # 6b.3 - Registry Configuration
    # ===========================================================
    Write-Out '[Registry Configuration]'
    Write-Log 'QUERY: Registry brand keys recursive dump'
    foreach ($brand in @('Equitrac','Nuance','Kofax','Tungsten','ControlSuite')) {
        foreach ($root in @("HKLM:\SOFTWARE\$brand", "HKLM:\SOFTWARE\WOW6432Node\$brand")) {
            if (Test-Path $root -EA SilentlyContinue) {
                Write-Out "  [$root]"
                Write-RegTree -Path $root
                Write-Out ''
            }
        }
    }
    # EQ* service parameters -- often store DB connection strings
    Write-Log 'QUERY: HKLM\SYSTEM\CurrentControlSet\Services EQ* -- service parameters'
    $eqSvcKeys = Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services' -EA SilentlyContinue |
        Where-Object { $_.PSChildName -match '^EQ' }
    if ($eqSvcKeys) {
        Write-Out '  [EQ* Service Registry Parameters]'
        foreach ($sk in $eqSvcKeys) {
            Write-RegTree -Path $sk.PSPath -MaxDepth 2
        }
        Write-Out ''
    }
    # Additional known ControlSuite paths
    foreach ($p in @(
        'HKLM:\SOFTWARE\Equitrac\Express',
        'HKLM:\SOFTWARE\Equitrac\Office',
        'HKLM:\SOFTWARE\Nuance\ControlSuite',
        'HKLM:\SOFTWARE\Kofax\ControlSuite',
        'HKLM:\SOFTWARE\Kofax\AutoStore'
    )) {
        if (Test-Path $p -EA SilentlyContinue) {
            Write-Out "  [$p]"
            Write-RegTree -Path $p
            Write-Out ''
        }
    }

    # ===========================================================
    # 6b.4 - Install Directory Contents
    # ===========================================================
    Write-Out '[Install Directories]'
    Write-Log 'QUERY: Install directory listing and config file scan'
    $installRoots = @(
        'C:\Program Files\Nuance\ControlSuite',
        'C:\Program Files\Kofax\ControlSuite',
        'C:\Program Files\Equitrac',
        'C:\Program Files (x86)\Equitrac',
        'C:\Program Files\Nuance',
        'C:\Program Files\Kofax',
        'C:\Program Files\Tungsten',
        'C:\ProgramData\Nuance\ControlSuite',
        'C:\ProgramData\Kofax\ControlSuite',
        'C:\ProgramData\Equitrac'
    )
    foreach ($root in $installRoots) {
        if (-not (Test-Path $root -EA SilentlyContinue)) { continue }
        Write-Out "  EXISTS: $root"
        Get-ChildItem $root -EA SilentlyContinue | ForEach-Object {
            $tag = if ($_.PSIsContainer) { '[DIR] ' } else { '[FILE]' }
            $sz  = if (-not $_.PSIsContainer) { "  $([Math]::Round($_.Length/1KB,1)) KB" } else { '' }
            Write-Out "    $tag $($_.Name)$sz"
        }
        # Config files: scan up to depth 3 for connection string clues
        $cfgFiles = Get-ChildItem $root -Include @('*.config','*.xml','*.ini','*.cfg') -Recurse -Depth 3 -EA SilentlyContinue |
            Select-Object -First 20
        foreach ($cf in $cfgFiles) {
            Write-Log "QUERY: Scanning config file $($cf.FullName)"
            $lines = Get-Content $cf.FullName -EA SilentlyContinue
            $connLines = $lines | Where-Object { $_ -match '(?i)datasource|data source|server=|database=|connectionstring|dbserver|dbhost|dbname' }
            if ($connLines) {
                Write-Out "    Config: $($cf.FullName)"
                foreach ($cl in $connLines | Select-Object -First 10) {
                    Write-Out "      $($cl.Trim())"
                }
            }
        }
        Write-Out ''
    }

    # ===========================================================
    # 6b.5 - Database Connection Detection
    # ===========================================================
    Write-Out '[Database Connection]'
    Write-Log 'QUERY: Database connection detection -- registry, ODBC, active TCP connections'

    # Method 1: parse connection strings from EQ* service registry
    $parsedDbServer = $null
    $parsedDbName   = $null
    $eqSvcKeys | ForEach-Object {
        Get-ChildItem $_.PSPath -EA SilentlyContinue | ForEach-Object {
            $props = Get-ItemProperty $_.PSPath -EA SilentlyContinue
            if ($props) {
                $props.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object {
                    $v = "$($_.Value)"
                    if ($v -match '(?i)(?:server|data source)\s*[=\\]') {
                        Write-Out "  Connection String (registry $($_.Name)): $v"
                        if ($v -match '(?i)(?:server|data source)\s*=\s*([^;,\s\\]+)') {
                            $parsedDbServer = $matches[1]
                        }
                        if ($v -match '(?i)(?:database|initial catalog)\s*=\s*([^;,\s]+)') {
                            $parsedDbName = $matches[1]
                        }
                        Write-Log "RESULT: Registry connection string found: $v"
                    }
                }
            }
        }
    }

    # Method 2: brand registry DB value names
    foreach ($brand in @('Equitrac','Nuance\ControlSuite','Kofax\ControlSuite','Kofax')) {
        $key = "HKLM:\SOFTWARE\$brand"
        if (-not (Test-Path $key -EA SilentlyContinue)) { continue }
        foreach ($vn in @('DatabaseServer','DBServer','CASDatabase','SQLServer','DatabaseHost','DbServer','DataSource','CASServer','ServerName')) {
            $val = Get-RegVal $key $vn
            if ($val) {
                Write-Out "  DB Registry ($brand\$vn): $val"
                Write-Log "RESULT: DB registry key $brand\$vn = $val"
                if (-not $parsedDbServer) { $parsedDbServer = $val }
            }
        }
    }

    # Summary + external DB flag
    if ($parsedDbServer) {
        $isExternal = $parsedDbServer -notmatch '^(localhost|127\.|\.\\|::1)' -and
                      $parsedDbServer -ne '.' -and
                      $parsedDbServer -ne $env:COMPUTERNAME
        Write-Out ''
        Write-Out "  DB Server  : $parsedDbServer"
        if ($parsedDbName) { Write-Out "  DB Name    : $parsedDbName" }
        Write-Out "  External DB: $(if ($isExternal) { 'YES -- ' + $parsedDbServer } else { 'No (local)' })"
        Write-Log "RESULT: DB server=$parsedDbServer external=$isExternal"
    }

    # Method 3: active TCP connections to SQL/PG ports on non-local IPs
    Write-Log 'QUERY: Get-NetTCPConnection -- active connections to DB ports on remote hosts'
    $extDbConns = Get-NetTCPConnection -EA SilentlyContinue |
        Where-Object {
            @(1433,5432,3306,1521) -contains $_.RemotePort -and
            $_.State -eq 'Established' -and
            $_.RemoteAddress -notmatch '^(127\.|::1|0\.0\.0\.0)'
        }
    if ($extDbConns) {
        $dbPortNames = @{1433='SQL Server';5432='PostgreSQL';3306='MySQL/MariaDB';1521='Oracle'}
        Write-Out '  Active connections to remote DB hosts:'
        foreach ($c in $extDbConns) {
            $dbType = $dbPortNames[[int]$c.RemotePort]
            Write-Out "    Remote: $($c.RemoteAddress):$($c.RemotePort) ($dbType)  LocalPort=$($c.LocalPort)"
            Write-Log "RESULT: Active remote DB connection $($c.RemoteAddress):$($c.RemotePort) ($dbType)"
        }
    } else {
        Write-Out '  No active connections to remote DB hosts detected'
    }
    Write-Out ''

    # ===========================================================
    # 6b.6 - SQL Server Inspection (Windows Auth)
    # ===========================================================
    Write-Out '[SQL Database Inspection]'
    Write-Log 'QUERY: SQL Server connectivity (Windows integrated auth)'

    # Build candidate list from registry first, then fallbacks
    $sqlCandidates = [System.Collections.Generic.List[string]]::new()
    $sqlInstKey = 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL'
    if (Test-Path $sqlInstKey -EA SilentlyContinue) {
        $instProps = Get-ItemProperty $sqlInstKey -EA SilentlyContinue
        if ($instProps) {
            $instProps.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object {
                $inst = $_.Name
                $sqlCandidates.Add( $(if ($inst -eq 'MSSQLSERVER') { 'localhost' } else { "localhost\$inst" }) )
            }
        }
    }
    foreach ($fb in @('localhost\SQLEXPRESS','localhost\SQLSERVER','localhost','(local)')) {
        if ($sqlCandidates -notcontains $fb) { $sqlCandidates.Add($fb) }
    }

    $sqlSrv = $null
    foreach ($c in $sqlCandidates) {
        $r = Invoke-SQLWin -Server $c -Query 'SELECT 1 AS ok'
        if ($r) { $sqlSrv = $c; break }
    }

    if (-not $sqlSrv) {
        Write-Out '  Cannot connect to SQL Server via Windows auth -- skipping DB inspection'
        Write-Log 'RESULT: SQL connection failed -- skipping DB inspection'
    } else {
        Write-Log "RESULT: SQL connected to $sqlSrv"
        $verDt = Invoke-SQLWin -Server $sqlSrv -Query "SELECT @@SERVERNAME AS SrvName, CAST(SERVERPROPERTY('Edition') AS NVARCHAR(128)) AS Edition, CAST(SERVERPROPERTY('ProductVersion') AS NVARCHAR(128)) AS Version"
        if ($verDt -and $verDt[0].Rows.Count -gt 0) {
            $vr = $verDt[0].Rows[0]
            Write-Out "  Connected  : $sqlSrv"
            Write-Out "  Server     : $($vr.SrvName)  Edition=$($vr.Edition)  Version=$($vr.Version)"
        }
        Write-Out ''

        # List non-system databases
        $dbListQ = @"
SELECT d.name, d.state_desc,
       CAST(ROUND(SUM(f.size)*8.0/1024,1) AS VARCHAR)+'MB' AS SizeMB
FROM   sys.databases d
LEFT   JOIN sys.master_files f ON d.database_id=f.database_id
WHERE  d.name NOT IN ('master','model','msdb','tempdb')
GROUP  BY d.name, d.state_desc
ORDER  BY d.name
"@
        $dbListDt = Invoke-SQLWin -Server $sqlSrv -Query $dbListQ
        if ($dbListDt -and $dbListDt[0].Rows.Count -gt 0) {
            Write-Out '  Databases:'
            foreach ($row in $dbListDt[0].Rows) {
                Write-Out "    $($row.name.PadRight(30)) State=$($row.state_desc)  Size=$($row.SizeMB)"
                Write-Log "RESULT: SQL database: $($row.name) state=$($row.state_desc)"
            }
            Write-Out ''

            # Deep-dive brand databases
            $brandDbPatterns = @('eqcas','equitrac','controlsuite','cas','nuance','kofax')
            foreach ($row in $dbListDt[0].Rows) {
                $dbN = $row.name
                $isBrand = $brandDbPatterns | Where-Object { $dbN -match $_ }
                if (-not $isBrand) { continue }

                Write-Out ''
                Write-Out "  ===== DATABASE: $dbN ====="

                # Table list with row counts
                $tblQ = @"
SELECT   t.name AS TableName, SUM(p.rows) AS Rows,
         CAST(ROUND(SUM(a.total_pages)*8.0/1024,2) AS VARCHAR)+'MB' AS Size
FROM     sys.tables t
JOIN     sys.indexes i ON t.object_id=i.object_id AND i.index_id IN (0,1)
JOIN     sys.partitions p ON i.object_id=p.object_id AND i.index_id=p.index_id
JOIN     sys.allocation_units a ON p.partition_id=a.container_id
GROUP BY t.name
HAVING   SUM(p.rows) > 0
ORDER BY SUM(p.rows) DESC
"@
                $tblDt = Invoke-SQLWin -Server $sqlSrv -Database $dbN -Query $tblQ
                if ($tblDt -and $tblDt[0].Rows.Count -gt 0) {
                    Write-Out "  Tables (by row count):"
                    foreach ($tr in $tblDt[0].Rows | Select-Object -First 30) {
                        Write-Out "    $($tr.TableName.PadRight(45)) $("$($tr.Rows)".PadLeft(8)) rows  $($tr.Size)"
                    }
                }
                Write-Out ''

                # Sample key config tables
                $keyConfigTables = @('cas_config','cas_server','cas_site','cas_domain','cas_device',
                                     'cas_pricelist','cas_license','cas_authprovider','cas_auth_provider',
                                     'cas_ldap','cas_smtp','cas_webserver','cas_casserver')
                foreach ($tbl in $keyConfigTables) {
                    $chk = Invoke-SQLWin -Server $sqlSrv -Database $dbN -Query "SELECT COUNT(*) AS cnt FROM sys.tables WHERE name='$tbl'"
                    if (-not ($chk -and $chk[0].Rows.Count -gt 0 -and $chk[0].Rows[0]['cnt'] -gt 0)) { continue }

                    Write-Out "  --- Table: $tbl ---"
                    # Get non-blob columns
                    $colQ = "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='$tbl' ORDER BY ORDINAL_POSITION"
                    $colDt = Invoke-SQLWin -Server $sqlSrv -Database $dbN -Query $colQ
                    if ($colDt -and $colDt[0].Rows.Count -gt 0) {
                        $blobTypes = @('image','varbinary','binary','timestamp','hierarchyid','geometry','geography')
                        $safeCols  = ($colDt[0].Rows | Where-Object { $_.DATA_TYPE -notin $blobTypes } |
                                      ForEach-Object { "[$($_.COLUMN_NAME)]" }) -join ','
                        if ($safeCols) {
                            $maxRows = if ($tbl -eq 'cas_config') { 1000 } else { 200 }
                            $dataDt  = Invoke-SQLWin -Server $sqlSrv -Database $dbN -Query "SELECT TOP $maxRows $safeCols FROM [$tbl] WITH (NOLOCK) ORDER BY 1" -TimeoutSec 45
                            Write-SqlTable $dataDt -MaxRows $maxRows
                        }
                    }
                    Write-Out ''
                }
            }
        } else {
            Write-Out '  No non-system databases found'
        }
    }
    Write-Out ''

    # ===========================================================
    # 6b.7 - Scheduled Tasks
    # ===========================================================
    Write-Out '[Scheduled Tasks]'
    Write-Log 'QUERY: Scheduled tasks -- brand-related'
    try {
        $brandTasks = Get-ScheduledTask -EA SilentlyContinue |
            Where-Object { $_.TaskName -match '(?i)Equitrac|Nuance|Kofax|ControlSuite|Tungsten|^EQ' }
        if ($brandTasks) {
            foreach ($t in $brandTasks) {
                $actions = ($t.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)".Trim() }) -join '; '
                Write-Out "  $($t.TaskPath)$($t.TaskName)  State=$($t.State)"
                Write-Out "    Action: $actions"
            }
        } else { Write-Out '  (none found)' }
    } catch { Write-Out '  (error reading scheduled tasks)' }
    Write-Out ''

    # ===========================================================
    # 6b.8 - Firewall Rules
    # ===========================================================
    Write-Out '[Firewall Rules]'
    Write-Log 'QUERY: Windows Firewall rules -- brand-related'
    try {
        $brandRules = Get-NetFirewallRule -EA SilentlyContinue |
            Where-Object { $_.DisplayName -match '(?i)Equitrac|Nuance|Kofax|ControlSuite|Tungsten' }
        if ($brandRules) {
            foreach ($r in $brandRules) {
                $pf = $r | Get-NetFirewallPortFilter -EA SilentlyContinue
                $af = $r | Get-NetFirewallApplicationFilter -EA SilentlyContinue
                Write-Out "  $($r.DisplayName)"
                Write-Out "    Direction=$($r.Direction)  Action=$($r.Action)  Enabled=$($r.Enabled)"
                if ($pf) { Write-Out "    Port: Proto=$($pf.Protocol)  Local=$($pf.LocalPort)  Remote=$($pf.RemotePort)" }
                if ($af -and $af.Program -ne 'Any') { Write-Out "    Program: $($af.Program)" }
            }
        } else { Write-Out '  (none found)' }
    } catch { Write-Out '  (error reading firewall rules)' }
    Write-Out ''

    # ===========================================================
    # 6b.9 - IIS Web Configuration
    # ===========================================================
    Write-Out '[IIS Web Configuration]'
    Write-Log 'QUERY: IIS web bindings and applications (ControlSuite web front-end)'
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\InetStp' -EA SilentlyContinue) {
        try {
            Import-Module WebAdministration -EA Stop
            $sites = Get-Website -EA SilentlyContinue
            if ($sites) {
                foreach ($site in $sites) {
                    Write-Out "  Site: $($site.Name)  State=$($site.State)  Path=$($site.PhysicalPath)"
                    $site.Bindings.Collection | ForEach-Object {
                        Write-Out "    Binding: $($_.Protocol) $($_.bindingInformation)"
                    }
                }
                $apps = Get-WebApplication -EA SilentlyContinue
                foreach ($app in $apps) {
                    Write-Out "  App: $($app.Path)  Phys=$($app.PhysicalPath)"
                }
            } else { Write-Out '  IIS installed but no sites configured' }
        } catch { Write-Out '  IIS detected but WebAdministration module unavailable' }
    } else { Write-Out '  IIS not detected' }
    Write-Out ''

    # ===========================================================
    # 6b.10 - Recent Event Log Entries
    # ===========================================================
    Write-Out '[Recent Event Log (brand-related, last 20 per log)]'
    Write-Log 'QUERY: Windows Event Log Application+System -- brand provider entries'
    foreach ($logName in @('Application','System')) {
        try {
            $events = Get-WinEvent -LogName $logName -MaxEvents 2000 -EA SilentlyContinue |
                Where-Object { $_.ProviderName -match '(?i)Equitrac|Nuance|Kofax|ControlSuite|^EQ' } |
                Select-Object -First 20
            if ($events) {
                Write-Out "  $logName ($($events.Count) entries):"
                foreach ($ev in $events) {
                    $msg = ($ev.Message -replace '[\r\n]+',' ').Trim()
                    if ($msg.Length -gt 200) { $msg = $msg.Substring(0,197) + '...' }
                    Write-Out "    $($ev.TimeCreated.ToString('yyyy-MM-dd HH:mm'))  $($ev.LevelDisplayName.PadRight(8))  Id=$($ev.Id)  Provider=$($ev.ProviderName)"
                    Write-Out "      $msg"
                }
            } else {
                Write-Out "  ${logName}: no brand-matching entries in last 2000 events"
            }
        } catch { Write-Out "  ${logName}: (read error)" }
    }
    Write-Out ''
}


# ===========================================================
# SECTION 7 - Windows Roles and Features
# ===========================================================
Write-Section 'Windows Roles and Features'
Write-Log 'QUERY: Get-WindowsFeature (roles and features install state)'

$gwfAvailable = $null
try { $gwfAvailable = Get-Command Get-WindowsFeature -ErrorAction Stop } catch {}

if ($gwfAvailable) {

    # Helper: format a feature line
    function Format-Feature([string]$label, $feat) {
        if ($feat) {
            $state = if ($feat.Installed) { 'Installed' } else { "Not Installed ($($feat.InstallState))" }
            Write-Out "  $label : $state"
            Write-Log "RESULT: $label = $state"
        } else {
            Write-Out "  $label : Not available on this OS"
            Write-Log "RESULT: $label = feature not found"
        }
    }

    # Load all features once (faster than repeated individual queries)
    Write-Log 'QUERY: Get-WindowsFeature * (loading all features)'
    $allFeatures = Get-WindowsFeature -ErrorAction SilentlyContinue

    function Get-Feat([string]$name) {
        return $allFeatures | Where-Object { $_.Name -eq $name } | Select-Object -First 1
    }

    # ---- .NET Framework 3.5 (1 of 3 sub-features) ----
    Write-Out ''
    Write-Out '.NET Framework 3.5 Features:'
    Format-Feature '.NET Framework 3.5 (Core)'      (Get-Feat 'NET-Framework-Core')
    Format-Feature '.NET 3.5 HTTP Activation'        (Get-Feat 'NET-HTTP-Activation')
    Format-Feature '.NET 3.5 Non-HTTP Activation'    (Get-Feat 'NET-Non-HTTP-Activ')

    # ---- .NET Framework 4.8 (1 of 7 sub-features) ----
    Write-Out ''
    Write-Out '.NET Framework 4.8 Features:'
    Format-Feature '.NET Framework 4.8 (Core)'       (Get-Feat 'NET-Framework-45-Core')
    Format-Feature '.NET 4.8 ASP.NET'                (Get-Feat 'NET-Framework-45-ASPNET')
    Format-Feature '.NET 4.8 WCF Services'           (Get-Feat 'NET-WCF-Services45')
    Format-Feature '.NET 4.8 WCF HTTP Activation'    (Get-Feat 'NET-WCF-HTTP-Activation45')
    Format-Feature '.NET 4.8 WCF MSMQ Activation'   (Get-Feat 'NET-WCF-MSMQ-Activation45')
    Format-Feature '.NET 4.8 WCF Named Pipe'         (Get-Feat 'NET-WCF-Pipe-Activation45')
    Format-Feature '.NET 4.8 WCF TCP Activation'     (Get-Feat 'NET-WCF-TCP-Activation45')
    Format-Feature '.NET 4.8 WCF TCP Port Sharing'   (Get-Feat 'NET-WCF-TCP-PortSharing45')

    # ---- Print and Document Services ----
    Write-Out ''
    Write-Out 'Print and Document Services:'
    Format-Feature 'Print and Document Services'     (Get-Feat 'Print-Services')
    Format-Feature 'Print Server'                    (Get-Feat 'Print-Server')
    Format-Feature 'Internet Printing'               (Get-Feat 'Print-Internet')
    Format-Feature 'LPD Service'                     (Get-Feat 'Print-LPD-Service')
    Format-Feature 'LPR Port Monitor'                (Get-Feat 'LPR-Port-Monitor')

    # ---- Telnet Client ----
    Write-Out ''
    Write-Out 'Remote Access Features:'
    Format-Feature 'Telnet Client'                   (Get-Feat 'Telnet-Client')
    Format-Feature 'Telnet Server'                   (Get-Feat 'Telnet-Server')

    # ---- Common roles summary (all installed roles) ----
    Write-Out ''
    Write-Out 'All Installed Roles and Role Services:'
    $installedRoles = $allFeatures | Where-Object { $_.Installed -and $_.FeatureType -eq 'Role' }
    if ($installedRoles) {
        foreach ($r in $installedRoles | Sort-Object DisplayName) {
            Write-Out "  [Role]    $($r.DisplayName)"
            Write-Log "RESULT: Role installed: $($r.DisplayName)"
        }
    } else {
        Write-Out '  None detected'
    }

    Write-Out ''
    Write-Out 'All Installed Features:'
    $installedFeatures = $allFeatures | Where-Object { $_.Installed -and $_.FeatureType -eq 'Feature' }
    if ($installedFeatures) {
        foreach ($f in $installedFeatures | Sort-Object DisplayName) {
            Write-Out "  [Feature] $($f.DisplayName)"
            Write-Log "RESULT: Feature installed: $($f.DisplayName)"
        }
    } else {
        Write-Out '  None detected'
    }

} else {
    # Fallback for non-Server OS (Windows 10/11) using Get-WindowsOptionalFeature
    Write-Log 'QUERY: Get-WindowsOptionalFeature (non-Server OS fallback)'
    Write-Out 'Note: Get-WindowsFeature not available - using Get-WindowsOptionalFeature'
    Write-Out ''

    function Get-OptFeat([string]$name) {
        return Get-WindowsOptionalFeature -Online -FeatureName $name -ErrorAction SilentlyContinue
    }
    function Format-OptFeature([string]$label, $feat) {
        $state = if ($feat) { $feat.State } else { 'Not Available' }
        Write-Out "  $label : $state"
        Write-Log "RESULT: $label = $state"
    }

    Write-Out '.NET Framework 3.5:'
    Format-OptFeature '.NET Framework 3.5'           (Get-OptFeat 'NetFx3')
    Format-OptFeature '.NET 3.5 HTTP Activation'     (Get-OptFeat 'WCF-HTTP-Activation')
    Format-OptFeature '.NET 3.5 Non-HTTP Activation' (Get-OptFeat 'WCF-NonHTTP-Activation')

    Write-Out ''
    Write-Out '.NET Framework 4.8:'
    Format-OptFeature '.NET Framework 4.8'           (Get-OptFeat 'NetFx4-AdvSrvs')

    Write-Out ''
    Write-Out 'Print and Document Services:'
    Format-OptFeature 'Print and Document Services'  (Get-OptFeat 'Printing-PrintToPDFServices-Features')
    Format-OptFeature 'LPR Port Monitor'             (Get-OptFeat 'TFTP')
    Format-OptFeature 'Telnet Client'                (Get-OptFeat 'TelnetClient')
}


# ===========================================================
# SECTION 8 - Database Configuration
# ===========================================================
Write-Section 'Database Configuration'

# Helper: read a registry value safely
function Get-RegVal([string]$path, [string]$name) {
    try { return (Get-ItemProperty -Path $path -Name $name -ErrorAction Stop).$name }
    catch { return $null }
}

# ---- SQL Server instances ----
Write-Log 'QUERY: HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL (enumerate instances)'
$sqlInstanceKey = 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL'

if (Test-Path $sqlInstanceKey) {
    $instanceProps = Get-ItemProperty $sqlInstanceKey -ErrorAction SilentlyContinue
    $instanceNames = $instanceProps.PSObject.Properties |
                     Where-Object { $_.Name -notmatch '^PS' } |
                     Select-Object Name, Value  # Name=instance, Value=folder e.g. MSSQL15.SQLEXPRESS

    if ($instanceNames) {
        Write-Out "SQL Server Instances Found: $($instanceNames.Count)"
        Write-Out ''

        foreach ($inst in $instanceNames) {
            $instanceName  = $inst.Name    # e.g. MSSQLSERVER or SQLEXPRESS
            $folderName    = $inst.Value   # e.g. MSSQL15.SQLEXPRESS
            $displayName   = if ($instanceName -eq 'MSSQLSERVER') { 'Default Instance (MSSQLSERVER)' } else { "Named Instance: $instanceName" }
            $svcName       = if ($instanceName -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$instanceName" }

            Write-Out "[SQL Instance: $instanceName]"
            Write-Log "QUERY: SQL instance $instanceName -> folder $folderName"

            # Edition and version from registry
            $setupKey = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$folderName\Setup"
            Write-Log "QUERY: $setupKey (Edition, Version)"
            $edition   = Get-RegVal $setupKey 'Edition'
            $version   = Get-RegVal $setupKey 'Version'
            $patchLevel = Get-RegVal $setupKey 'PatchLevel'
            $sqlDir    = Get-RegVal $setupKey 'SQLDataRoot'

            Write-Out "  Display         : $displayName"
            if ($edition)    { Write-Out "  Edition         : $edition" }
            if ($version)    { Write-Out "  Version         : $version" }
            if ($patchLevel) { Write-Out "  Patch Level     : $patchLevel" }
            if ($sqlDir)     { Write-Out "  Data Root       : $sqlDir" }

            # Service status
            Write-Log "QUERY: Get-Service $svcName"
            $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
            if ($svc) {
                Write-Out "  Service Status  : $($svc.Status)"
                Write-Out "  Startup Type    : $($svc.StartType)"
                $wmiSvc = Get-CimInstance Win32_Service | Where-Object { $_.Name -eq $svcName }
                if ($wmiSvc) { Write-Out "  Service Account : $($wmiSvc.StartName)" }
            } else {
                Write-Out "  Service Status  : Not running / not found"
            }

            # SQL Server Agent
            $agentSvcName = if ($instanceName -eq 'MSSQLSERVER') { 'SQLSERVERAGENT' } else { "SQLAgent`$$instanceName" }
            $agentSvc = Get-Service -Name $agentSvcName -ErrorAction SilentlyContinue
            if ($agentSvc) { Write-Out "  SQL Agent       : $($agentSvc.Status)" }

            # Network configuration from registry
            $netKey  = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$folderName\MSSQLServer\SuperSocketNetLib"
            $tcpKey  = "$netKey\Tcp"
            $npKey   = "$netKey\Np"

            Write-Log "QUERY: $tcpKey (TCP/IP config)"
            if (Test-Path $tcpKey) {
                $tcpEnabled = Get-RegVal $tcpKey 'Enabled'
                Write-Out "  TCP/IP          : $(if ($tcpEnabled -eq 1) { 'Enabled' } else { 'Disabled' })"

                # IPAll - static and dynamic port
                $ipAllKey = "$tcpKey\IPAll"
                if (Test-Path $ipAllKey) {
                    $staticPort  = Get-RegVal $ipAllKey 'TcpPort'
                    $dynamicPort = Get-RegVal $ipAllKey 'TcpDynamicPorts'
                    if ($staticPort)  { Write-Out "  TCP Port (IPAll): $staticPort" }
                    if ($dynamicPort) { Write-Out "  Dynamic Port    : $dynamicPort" }
                }

                # IP1, IP2, IP3... individual IP entries
                $ipSubKeys = Get-ChildItem $tcpKey -ErrorAction SilentlyContinue |
                             Where-Object { $_.PSChildName -match '^IP\d+$' } |
                             Sort-Object PSChildName
                foreach ($ipKey in $ipSubKeys) {
                    $ipLabel   = $ipKey.PSChildName
                    $ipAddr    = Get-RegVal $ipKey.PSPath 'IpAddress'
                    $ipActive  = Get-RegVal $ipKey.PSPath 'Active'
                    $ipEnabled = Get-RegVal $ipKey.PSPath 'Enabled'
                    $ipPort    = Get-RegVal $ipKey.PSPath 'TcpPort'
                    $ipDyn     = Get-RegVal $ipKey.PSPath 'TcpDynamicPorts'

                    $activeText  = if ($ipActive  -eq 1) { 'Yes' } else { 'No' }
                    $enabledText = if ($ipEnabled -eq 1) { 'Yes' } else { 'No' }
                    $portInfo    = if ($ipPort) { "Port=$ipPort" } elseif ($ipDyn) { "DynamicPort=$ipDyn" } else { 'Port=inherited from IPAll' }

                    Write-Out "  ${ipLabel}              : Address=$ipAddr  Active=$activeText  Enabled=$enabledText  $portInfo"
                    Write-Log "RESULT: $ipLabel addr=$ipAddr active=$activeText enabled=$enabledText $portInfo"
                }
            } else {
                Write-Out "  TCP/IP          : Configuration key not found"
            }

            Write-Log "QUERY: $npKey (Named Pipes config)"
            if (Test-Path $npKey) {
                $npEnabled  = Get-RegVal $npKey 'Enabled'
                $pipeName   = Get-RegVal $npKey 'PipeName'
                Write-Out "  Named Pipes     : $(if ($npEnabled -eq 1) { 'Enabled' } else { 'Disabled' })"
                if ($pipeName) { Write-Out "  Pipe Name       : $pipeName" }
            } else {
                Write-Out "  Named Pipes     : Configuration key not found"
            }

            Write-Out '  ---'
            Write-Log "RESULT: SQL instance $instanceName edition=$edition version=$version"
        }
    } else {
        Write-Out 'SQL Server: No instances found in registry'
        Write-Log 'RESULT: No SQL Server instances detected'
    }
} else {
    Write-Out 'SQL Server: Not installed (registry key absent)'
    Write-Log 'RESULT: SQL Server not installed'
}
Write-Out ''

# ---- PostgreSQL ----
Write-Log 'QUERY: PostgreSQL - services, registry, install directories'
$pgSvcs = Get-Service | Where-Object { $_.DisplayName -like '*postgresql*' -or $_.Name -like 'postgresql*' }
$pgReg  = Get-ItemProperty 'HKLM:\SOFTWARE\PostgreSQL\Installations\*' -ErrorAction SilentlyContinue
$pgDirs = @('C:\Program Files\PostgreSQL') | Where-Object { Test-Path $_ }

if ($pgSvcs -or $pgReg -or $pgDirs) {
    Write-Out '[PostgreSQL]'
    if ($pgReg) {
        foreach ($pg in $pgReg) {
            Write-Out "  Version         : $($pg.Version)"
            Write-Out "  Base Directory  : $($pg.'Base Directory')"
            Write-Out "  Data Directory  : $($pg.'Data Directory')"
            Write-Out "  Port            : $($pg.Port)"
            Write-Out "  Service Account : $($pg.'Service Account')"
            Write-Log "RESULT: PostgreSQL v$($pg.Version) port=$($pg.Port)"
        }
    }
    if ($pgSvcs) {
        foreach ($svc in $pgSvcs) {
            Write-Out "  Service         : $($svc.DisplayName) [$($svc.Status)]"
        }
    }
    if (-not $pgReg -and $pgDirs) {
        # Fallback: list version folders in install dir
        $versions = Get-ChildItem 'C:\Program Files\PostgreSQL' -Directory -ErrorAction SilentlyContinue
        foreach ($v in $versions) { Write-Out "  Installed Version Dir: $($v.Name)" }
    }
} else {
    Write-Out 'PostgreSQL: Not detected'
    Write-Log 'RESULT: PostgreSQL not detected'
}
Write-Out ''

# ---- Other databases (MySQL, MariaDB, Oracle, MongoDB) ----
Write-Log 'QUERY: Other database engines (services + registry scan)'
$dbChecks = @(
    @{ Name = 'MySQL';   Patterns = @('*MySQL*') },
    @{ Name = 'MariaDB'; Patterns = @('*MariaDB*') },
    @{ Name = 'Oracle';  Patterns = @('*OracleService*', '*OracleTNS*') },
    @{ Name = 'MongoDB'; Patterns = @('*MongoDB*') },
    @{ Name = 'Redis';   Patterns = @('*Redis*') }
)
foreach ($db in $dbChecks) {
    $found = $false
    foreach ($pat in $db.Patterns) {
        $svc = $allServices | Where-Object { $_.DisplayName -like $pat -or $_.Name -like $pat } | Select-Object -First 1
        if ($svc) {
            Write-Out "$($db.Name): Detected - Service '$($svc.DisplayName)' [$($svc.Status)]"
            Write-Log "RESULT: $($db.Name) service=$($svc.DisplayName) status=$($svc.Status)"
            $found = $true; break
        }
    }
    if (-not $found) { Write-Log "RESULT: $($db.Name) not detected" }
}
Write-Out ''

# ---- ODBC Data Sources ----
Write-Log 'QUERY: HKLM ODBC System DSNs and HKCU User DSNs (read-only)'
Write-Out '[ODBC Data Sources]'

$odbcPaths = @(
    @{ Scope = 'System'; Key = 'HKLM:\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources' },
    @{ Scope = 'System (32-bit)'; Key = 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\ODBC Data Sources' },
    @{ Scope = 'User'; Key = 'HKCU:\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources' }
)

$odbcFound = $false
foreach ($odbcScope in $odbcPaths) {
    if (Test-Path $odbcScope.Key) {
        $dsns = Get-ItemProperty $odbcScope.Key -ErrorAction SilentlyContinue
        $dsnNames = $dsns.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' }
        foreach ($dsn in $dsnNames) {
            $odbcFound = $true
            $dsnName   = $dsn.Name
            $dsnDriver = $dsn.Value
            $dsnKey    = ($odbcScope.Key -replace 'ODBC Data Sources', $dsnName)
            $dsnProps  = Get-ItemProperty $dsnKey -ErrorAction SilentlyContinue

            Write-Out "  DSN             : $dsnName  [$($odbcScope.Scope)]"
            Write-Out "  Driver          : $dsnDriver"
            if ($dsnProps) {
                if ($dsnProps.Server)   { Write-Out "  Server          : $($dsnProps.Server)" }
                if ($dsnProps.Database) { Write-Out "  Database        : $($dsnProps.Database)" }
                if ($dsnProps.Port)     { Write-Out "  Port            : $($dsnProps.Port)" }
                if ($dsnProps.Servername) { Write-Out "  Server Name     : $($dsnProps.Servername)" }
            }
            Write-Out '  ---'
            Write-Log "RESULT: ODBC DSN=$dsnName driver=$dsnDriver server=$($dsnProps.Server)"
        }
    }
}
if (-not $odbcFound) {
    Write-Out '  No ODBC DSNs configured'
    Write-Log 'RESULT: No ODBC DSNs found'
}
Write-Out ''

# ---- Active connections to remote database servers ----
Write-Log 'QUERY: Get-NetTCPConnection - active outbound connections to database ports (1433, 5432, 3306, 1521, 27017)'
Write-Out '[Remote Database Connections]'
$dbPorts    = @(1433, 5432, 3306, 1521, 27017)
$dbPortMap  = @{ 1433 = 'SQL Server'; 5432 = 'PostgreSQL'; 3306 = 'MySQL/MariaDB'; 1521 = 'Oracle'; 27017 = 'MongoDB' }
$remoteConns = Get-NetTCPConnection -ErrorAction SilentlyContinue |
               Where-Object { $dbPorts -contains $_.RemotePort -and $_.State -eq 'Established' }

if ($remoteConns) {
    foreach ($conn in $remoteConns) {
        $dbType = $dbPortMap[[int]$conn.RemotePort]
        Write-Out "  Remote          : $($conn.RemoteAddress):$($conn.RemotePort) ($dbType) - $($conn.State)"
        Write-Log "RESULT: Remote DB connection to $($conn.RemoteAddress):$($conn.RemotePort) ($dbType)"
    }
} else {
    Write-Out '  No active outbound connections to known database ports'
    Write-Log 'RESULT: No remote database connections active'
}


# ===========================================================
# SECTION 8 - Print Server Role and Configuration
# ===========================================================
Write-Section 'Print Server'

# --- Windows Print Server role ---
Write-Log 'QUERY: Get-WindowsFeature Print-Server, Print-Management (role installation check)'
$printServerRole = Get-WindowsFeature -Name 'Print-Server'     -ErrorAction SilentlyContinue
$printMgmtTool   = Get-WindowsFeature -Name 'Print-Management' -ErrorAction SilentlyContinue

if ($printServerRole) {
    $roleStatus = if ($printServerRole.Installed) { 'Installed' } else { 'Not Installed' }
    Write-Field 'Print Server Role'       $roleStatus
    Write-Field 'Print Server Role State' $printServerRole.InstallState
    Write-Log "RESULT: Print-Server role = $roleStatus"
} else {
    Write-Field 'Print Server Role' 'Not available (may not be Windows Server)'
    Write-Log 'RESULT: Get-WindowsFeature not available on this OS'
}

if ($printMgmtTool) {
    $mgmtStatus = if ($printMgmtTool.Installed) { 'Installed' } else { 'Not Installed' }
    Write-Field 'Print Management Tools' $mgmtStatus
    Write-Log "RESULT: Print-Management = $mgmtStatus"
}

# --- Spooler service ---
Write-Log 'QUERY: Get-Service Spooler'
$spooler = Get-Service -Name 'Spooler' -ErrorAction SilentlyContinue
if ($spooler) {
    Write-Field 'Print Spooler Service' $spooler.Status
    Write-Field 'Spooler Startup Type'  $spooler.StartType
    Write-Log "RESULT: Spooler = $($spooler.Status), StartType = $($spooler.StartType)"
}

# --- Spool folder location ---
Write-Log 'QUERY: HKLM DefaultSpoolDirectory (spool folder path)'
$spoolFolder = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -Name DefaultSpoolDirectory -EA SilentlyContinue).DefaultSpoolDirectory
if (-not $spoolFolder) { $spoolFolder = 'C:\Windows\System32\spool\PRINTERS (default)' }
Write-Field 'Spool Folder' $spoolFolder
Write-Log "RESULT: Spool folder = $spoolFolder"
Write-Out ''

# --- Load all WMI printers and x86 drivers once ---
Write-Log 'QUERY: Get-Printer, Get-CimInstance Win32_Printer, Get-PrinterDriver'
$printers    = Get-Printer -ErrorAction SilentlyContinue
$wmiPrinters = Get-CimInstance Win32_Printer -ErrorAction SilentlyContinue
$allDrivers  = Get-PrinterDriver -ErrorAction SilentlyContinue
# Load drivers split by architecture using -PrinterEnvironment
$driversX64 = @(Get-PrinterDriver -PrinterEnvironment 'Windows x64'     -EA SilentlyContinue)
$driversX86 = @(Get-PrinterDriver -PrinterEnvironment 'Windows NT x86'  -EA SilentlyContinue)
$x86DriverNames = $driversX86 | ForEach-Object { $_.Name }

# Duplex map
$duplexMap = @{
    1 = 'One-sided (Simplex)'
    2 = 'Two-sided Long Edge (Portrait)'
    3 = 'Two-sided Short Edge (Landscape)'
}
# Color map
$colorMap = @{ 1 = 'Black & White (Monochrome)'; 2 = 'Color'; 3 = 'Color' }

# ICM (dmColor) map
$icmMap = @{
    1 = 'ICM Disabled (use dmColor from application)'
    2 = 'ICM handled by OS'
    3 = 'ICM handled by device'
    4 = 'ICM handled by host'
}

# Printer attribute bit flags (Win32_Printer.Attributes)
$ATTR_QUEUED      = 0x00000001   # spooled (not direct)
$ATTR_SHARED      = 0x00000008
$ATTR_RAW_ONLY    = 0x00001000   # RAW only = advanced features OFF
$ATTR_PUBLISHED   = 0x00002000   # published in AD directory
$ATTR_BIDI        = 0x00000800   # enable bidirectional
$ATTR_RENDER_CLI  = 0x00040000   # render print jobs on client

if ($printers) {
    Write-Out "Printer Queues ($($printers.Count) total):"
    foreach ($p in $printers) {

        # Match WMI object for extended attributes
        $wmi   = $wmiPrinters | Where-Object { $_.Name -eq $p.Name } | Select-Object -First 1
        $attrs = if ($wmi) { [int]$wmi.Attributes } else { 0 }

        # Printing defaults via Get-PrintConfiguration
        Write-Log "QUERY: Get-PrintConfiguration -PrinterName '$($p.Name)'"
        $cfg = Get-PrintConfiguration -PrinterName $p.Name -ErrorAction SilentlyContinue

        # Printing Defaults from PrintTicket XML
        $stapleText    = 'Unknown'
        $offsetText    = 'Unknown'
        $colorMgmtText = 'Unknown'
        if ($cfg -and $cfg.PrintTicketXML) {
            [xml]$pt = $cfg.PrintTicketXML
            $ns = @{ psf = 'http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework';
                     psk = 'http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords' }

            # PageColorManagement - maps to FujiFilm "Use the dmColor specified by the application":
            #   psk:None   = On  (driver passes application's dmColor through unchanged)
            #   psk:System = Off (Windows ICM manages colour conversion)
            #   psk:Driver = Off (driver handles ICM internally)
            #   psk:Device = Off (device hardware handles ICM)
            $colorMgmtNode = Select-Xml -Xml $pt -XPath "//psf:Feature[contains(@name,'PageColorManagement')]//psf:Option" -Namespace $ns |
                              Select-Object -First 1
            if ($colorMgmtNode) {
                $cmVal = $colorMgmtNode.Node.GetAttribute('name')
                $colorMgmtText = switch -Wildcard ($cmVal) {
                    '*:None'   { 'On  (passes application dmColor through; no driver override)' }
                    '*:System' { 'Off (Windows system ICM manages colour)' }
                    '*:Driver' { 'Off (driver manages ICM internally)' }
                    '*:Device' { 'Off (device hardware manages ICM)' }
                    default    { $cmVal }
                }
            }

            # Staple (job-level default; FujiFilm finisher installable options live in
            # driver private data, not the standard PrintTicket - see finisher note below)
            $stapleNode = Select-Xml -Xml $pt -XPath "//psf:Feature[contains(@name,'Staple') or contains(@name,'staple') or contains(@name,'Finishing')]//psf:Option" -Namespace $ns |
                          Select-Object -First 1
            if ($stapleNode) {
                $stapleVal  = $stapleNode.Node.GetAttribute('name')
                $stapleText = if ($stapleVal -like '*None*' -or $stapleVal -like '*none*') { 'Disabled (default job)' } else { "Enabled - $($stapleVal -replace '^.*:','')" }
            }

            # Output bin / offset stacking (match OutputBin only, not InputBin)
            $binNode = Select-Xml -Xml $pt -XPath "//psf:Feature[contains(@name,'OutputBin') or (contains(@name,'Bin') and not(contains(@name,'Input')))]//psf:Option" -Namespace $ns |
                       Select-Object -First 1
            if ($binNode) {
                $binVal     = $binNode.Node.GetAttribute('name')
                if     ($binVal -like '*OffsetEachSet*' -or $binVal -like '*PerSet*') { $offsetText = 'Offset per Set' }
                elseif ($binVal -like '*OffsetEachJob*' -or $binVal -like '*PerJob*') { $offsetText = 'Offset per Job' }
                elseif ($binVal -like '*AutoSelect*')                                 { $offsetText = 'Auto Select (no offset)' }
                elseif ($binVal -like '*Standard*' -or $binVal -like '*Main*')        { $offsetText = 'Standard Bin (no offset)' }
                else   { $offsetText = "No offset ($($binVal -replace '^.*:',''))" }
            }
        }

        # Flags
        $advFeatures    = if (($attrs -band $ATTR_RAW_ONLY) -eq 0) { 'Enabled' } else { 'Disabled (RAW only)' }
        $renderOnClient = if (($attrs -band $ATTR_RENDER_CLI) -ne 0) { 'Yes' } else { 'No' }
        $isShared       = $p.Shared
        $shareName      = if ($isShared) { $p.ShareName } else { '' }
        $isPublished    = if ($wmi) { ($attrs -band $ATTR_PUBLISHED) -ne 0 } else { $p.Published }

        # x86 additional driver installed for this queue?
        $has32bitDriver = $x86DriverNames -contains $p.DriverName

        # ICM method - read from Default DevMode binary at byte offset 188 (DWORD)
        # Win32_Printer.ICMMethod is unreliable for third-party drivers; registry is authoritative
        $icmText = 'Not available'
        try {
            $dmRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$($p.Name)"
            $dmBytes   = (Get-ItemProperty -Path $dmRegPath -Name 'Default DevMode' -EA Stop).'Default DevMode'
            if ($dmBytes -and $dmBytes.Count -ge 192) {
                $icmRaw  = [BitConverter]::ToInt32($dmBytes, 188)
                $icmText = if ($icmMap.ContainsKey($icmRaw)) { $icmMap[$icmRaw] } else { "Method $icmRaw (driver-defined)" }
            }
        } catch { $icmText = 'Not available (registry read error)' }

        # Paper size
        $paperText = if ($cfg -and $cfg.PaperSize) { [string]$cfg.PaperSize } else { 'Unknown' }

        # Duplex  (DuplexingMode is an enum; OneSided = 0 which is falsy, so check $null explicitly)
        $duplexText = 'Unknown'
        if ($cfg -and ($cfg.DuplexingMode -ne $null)) {
            $dm = [string]$cfg.DuplexingMode
            $duplexText = switch ($dm) {
                'OneSided'              { 'One-sided (Simplex)' }
                'TwoSidedLongEdge'      { 'Two-sided Long Edge (Portrait / Flip on Long)' }
                'TwoSidedShortEdge'     { 'Two-sided Short Edge (Landscape / Flip on Short)' }
                default                 { $dm }
            }
        }

        # Color  (Color is a boolean in Get-PrintConfiguration)
        $colorText = 'Unknown'
        if ($cfg) {
            if ($cfg.Color -eq $true)  { $colorText = 'Color' }
            elseif ($cfg.Color -eq $false) { $colorText = 'Black & White (Monochrome)' }
        }

        # Write output
        Write-Out ''
        Write-Out "  Queue Name              : $($p.Name)"
        Write-Out "  Driver                  : $($p.DriverName)"
        Write-Out "  Port                    : $($p.PortName)"
        Write-Out "  Status                  : $($p.PrinterStatus)"
        Write-Out ''
        Write-Out "  -- Queue Settings --"
        Write-Out "  Advanced Print Features : $advFeatures"
        Write-Out "  Render on Client        : $renderOnClient"
        Write-Out ''
        Write-Out "  -- Printing Defaults --"
        Write-Out "  Paper Size              : $paperText"
        Write-Out "  2-sided Print (Duplex)  : $duplexText"
        Write-Out "  Output Color            : $colorText"
        Write-Out "  Staple                  : $stapleText"
        Write-Out "  Offset Stacking         : $offsetText"
        Write-Out ''
        Write-Out "  -- Advanced Settings --"
        Write-Out "  Use Application Color   : $colorMgmtText"
        Write-Out "  ICM Method              : $icmText"
        # Finisher note for FF/FujiFilm: installable options are in driver private data
        # Installable options (hardware config) via Get-PrinterProperty
        $printerProps = $null
        try { $printerProps = Get-PrinterProperty -PrinterName $p.Name -EA Stop } catch {}
        if ($printerProps -and $printerProps.Count -gt 0) {
            $pp = @{}
            $printerProps | ForEach-Object { $pp[$_.PropertyName] = $_.Value }

            # Finisher units installed
            $finInst = @('A','B','C','D') | ForEach-Object {
                $v = $pp["Config:OP_Finisher$_"]
                if ($v -and $v -ne 'No') { "Finisher $_ ($v)" }
            }
            $finText = if ($finInst) { $finInst -join ', ' } else { 'None' }

            # Staple capability
            $stapleInstall = if ($pp['Config:DC_FIN_Staple'] -eq 'Yes') {
                $extras = @()
                if ($pp['Config:DC_FIN_FreeStaple'] -eq 'Yes') { $extras += 'FreeStaple' }
                if ($pp['Config:DC_FIN_4Staple']    -eq 'Yes') { $extras += '4-staple'   }
                'Yes' + $(if ($extras) { ' (' + ($extras -join ', ') + ')' } else { '' })
            } else { 'No' }

            # Punch capability
            $punchInstall = if ($pp['Config:DC_FIN_Punch'] -eq 'Yes') {
                $holes = @()
                if ($pp['Config:OP_Punch_2_3'] -and $pp['Config:OP_Punch_2_3'] -ne 'No') { $holes += '2/3-hole' }
                if ($pp['Config:OP_Punch_2_4'] -and $pp['Config:OP_Punch_2_4'] -ne 'No') { $holes += '2/4-hole' }
                'Yes' + $(if ($holes) { ' (' + ($holes -join ', ') + ')' } else { '' })
            } else { 'No' }

            # Booklet maker
            $bookletVal     = $pp['Config:OP_Booklet']
            $bookletInstall = if ($bookletVal -and $bookletVal -ne 'No') {
                $bExtra = @()
                if ($pp['Config:DC_FIN_BiFold'] -eq 'Yes') { $bExtra += 'BiFold' }
                if ($pp['Config:DC_FIN_CZFold'] -eq 'Yes') { $bExtra += 'CZFold' }
                $bookletVal + $(if ($bExtra) { ' (' + ($bExtra -join ', ') + ')' } else { '' })
            } else { 'No' }

            # Offset stacking
            $offsetInstall = if ($pp['Config:DC_OffsetStacking'] -eq 'Yes') { 'Yes' } else { 'No' }

            Write-Out ''
            Write-Out '  -- Installable Options --'
            Write-Out "  Finisher Installed      : $finText"
            Write-Out "  Staple Capability       : $stapleInstall"
            Write-Out "  Punch Capability        : $punchInstall"
            Write-Out "  Booklet Maker           : $bookletInstall"
            Write-Out "  Offset Stacking         : $offsetInstall"
        } elseif ($p.DriverName -like 'FF *' -or $p.DriverName -like '*FujiFilm*' -or $p.DriverName -like '*Apeos*') {
            Write-Out ''
            Write-Out '  -- Installable Options --'
            Write-Out '  N/A (software/virtual driver - no hardware options)'
        }
        Write-Out ''
        Write-Out "  -- Sharing & Directory --"
        Write-Out "  Shared                  : $(if ($isShared) { "Yes (Share name: $shareName)" } else { 'No' })"
        Write-Out "  Listed in Directory     : $(if ($isPublished) { 'Yes (Published in AD)' } else { 'No' })"
        Write-Out ''
        Write-Out "  -- Driver Availability --"
        Write-Out "  x64 Driver              : Installed ($($p.DriverName))"
        Write-Out "  x86 (32-bit) Driver     : $(if ($has32bitDriver) { 'Installed' } else { 'Not installed' })"
        Write-Out '  ============================================'

        Write-Log "RESULT: Queue=$($p.Name) driver=$($p.DriverName) duplex=$duplexText color=$colorText paper=$paperText staple=$stapleText shared=$isShared published=$isPublished x86=$has32bitDriver"
    }
    $ffCount = ($printers | Where-Object {
        $_.DriverName -like '*FujiFilm*'  -or $_.DriverName -like '*Fuji Xerox*' -or
        $_.DriverName -like '*FUJIFILM*'  -or $_.DriverName -like 'FF *'          -or
        $_.DriverName -like '*Apeos*'
    }).Count
    $fxCount = ($printers | Where-Object { $_.DriverName -like '*Fuji Xerox*' -or $_.DriverName -like '*FX*' }).Count
    Write-Log "RESULT: $($printers.Count) queues total, FujiFilm/FF=$ffCount, Fuji Xerox=$fxCount"
} else {
    Write-Out 'Printer Queues: None detected'
    Write-Log 'RESULT: No print queues found'
}
Write-Out ''

# --- Printer Drivers ---
Write-Log 'QUERY: Get-PrinterDriver (all installed drivers)'
# Use the already-loaded per-arch driver lists
$drivers = $driversX64 + $driversX86
if ($drivers -and $drivers.Count -gt 0) {
    $x64Count = $driversX64.Count; $x86Count = $driversX86.Count
    Write-Out "Installed Printer Drivers ($x64Count x64, $x86Count x86):"
    Write-Out "  --- x64 (64-bit) ---"
    foreach ($drv in ($driversX64 | Sort-Object Name)) {
        Write-Out "  Driver  : $($drv.Name)"
        Write-Out "  Version : $($drv.MajorVersion)"
        Write-Out "  Print Processor : $($drv.PrintProcessor)"
        Write-Out "  ---"
    }
    if ($driversX86.Count -gt 0) {
        Write-Out "  --- x86 (32-bit) ---"
        foreach ($drv in ($driversX86 | Sort-Object Name)) {
            Write-Out "  Driver  : $($drv.Name)"
            Write-Out "  Version : $($drv.MajorVersion)"
            Write-Out "  Print Processor : $($drv.PrintProcessor)"
            Write-Out "  ---"
        }
    }
    Write-Log "RESULT: $x64Count x64 drivers, $x86Count x86 drivers installed"
} else {
    Write-Out 'Printer Drivers: None detected'
    Write-Log 'RESULT: No printer drivers found'
}
Write-Out ''

# --- Printer Ports ---
Write-Log 'QUERY: Get-PrinterPort (all configured ports)'
$ports = Get-PrinterPort -ErrorAction SilentlyContinue
if ($ports) {
    Write-Out "Configured Printer Ports ($($ports.Count) total):"
    foreach ($port in ($ports | Sort-Object Name)) {
        $portDetail = "  Port: $($port.Name)"
        if ($port.PrinterHostAddress) { $portDetail += " | IP: $($port.PrinterHostAddress)" }
        if ($port.PortNumber)         { $portDetail += " | TCP Port: $($port.PortNumber)" }
        if ($port.Description)        { $portDetail += " | Type: $($port.Description)" }
        Write-Out $portDetail
    }
    Write-Log "RESULT: $($ports.Count) ports configured"
} else {
    Write-Out 'Printer Ports: None detected'
    Write-Log 'RESULT: No printer ports found'
}


# ===========================================================
# SECTION 8 - Device Counts (FujiFilm / Fuji Xerox summary)
# ===========================================================
Write-Section 'Device Counts'
Write-Log 'QUERY: Summarising FujiFilm/Fuji Xerox queues from printer list'
if ($printers) {
    $ffCount = ($printers | Where-Object {
        $_.DriverName -like '*FujiFilm*'  -or $_.DriverName -like '*Fuji Xerox*' -or
        $_.DriverName -like '*FUJIFILM*'  -or $_.DriverName -like 'FF *'          -or
        $_.DriverName -like '*Apeos*'
    }).Count
    $fxCount = ($printers | Where-Object { $_.DriverName -like '*Fuji Xerox*' -or $_.DriverName -like '*FX*' }).Count
    Write-Field 'FujiFilm Devices' ([string]$ffCount)
    Write-Field 'FF Devices'       ([string]$ffCount)
    Write-Field 'Fuji Xerox'       ([string]$fxCount)
    Write-Field 'FX Devices'       ([string]$fxCount)
} else {
    Write-Field 'FujiFilm Devices' '0'
    Write-Field 'FF Devices'       '0'
    Write-Field 'Fuji Xerox'       '0'
    Write-Field 'FX Devices'       '0'
}


# ===========================================================
# FOOTER
# ===========================================================
Write-Out ''
Write-Out '============================================================'
Write-Out "End of Export - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Out '============================================================'

Write-Log '--- Export complete ---'
Write-Log "Data file : $outFile"
Write-Log "Log file  : $logFile"

Write-Trace 'Script complete'
Write-Trace "Data file  : $outFile"
Write-Trace "Log file   : $logFile"
Write-Trace "Trace file : $traceFile"

Write-Host ''
Write-Host '============================================================'
Write-Host '  Export complete.'
Write-Host ''
Write-Host "  Data  : $outFile"
Write-Host "  Log   : $logFile"
Write-Host "  Trace : $traceFile"
Write-Host ''
Write-Host '  Copy the DATA FILE to:'
Write-Host '  ConsultantApp\data\Deployment\PMS\<CustomerName>\'
Write-Host '============================================================'
Write-Host ''





