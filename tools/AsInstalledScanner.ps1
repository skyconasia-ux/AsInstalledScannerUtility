#Requires -Version 5.1
# AsInstalledScanner.ps1
# Equitrac / ControlSuite configuration scanner.
#
# Usage (interactive menu):
#   powershell -NoProfile -ExecutionPolicy Bypass -File AsInstalledScanner.ps1
#
# Usage (scripted / SSH):
#   powershell -NoProfile -ExecutionPolicy Bypass -File AsInstalledScanner.ps1 -Mode After
#
# Modes:
#   Before  - capture baseline snapshot (raw data only, no report)
#   After   - capture current state + FULL.txt + SUMMARY.txt + REPORT.html + metadata.json
#   Compare - diff latest Before vs After folder -> COMPARE_DIFF.txt + COMPARE_REPORT.html
#   Full    - After + Compare in one run
#
# Output: .\Output\HOSTNAME_YYYYMMDD_HHMMSS_MODE\
#
# NOTE: ASCII-compatible, PS 5.1. READ-ONLY - makes no system changes.
# Supersedes Export-EquitracConfig.ps1 for the After/Full modes.

param(
    [ValidateSet('Before','After','Compare','Full','')]
    [string]$Mode = ''
)

$ErrorActionPreference = 'SilentlyContinue'

# ============================================================
# CONFIG - edit for each installation
# ============================================================
$SqlServer  = '.\SQLExpress'
$SqlDb      = 'eqcas'
$SqlUser    = 'sa'
$SqlPass    = 'FujiFilm_11111'
$Sqlite3    = 'C:\Windows\System32\sqlite3.exe'
$AppName    = 'ControlSuite'
$DbBase     = 'C:\Windows\System32\config\systemprofile\AppData\Local\Equitrac\Equitrac Platform Component'

# ============================================================
# WATCHED EQVAR KEYS
# To add a new setting: append one entry here.
# Section values: auth | smtp | job | quota | currency | license | device
# IsXml: $true = value is an XML blob (truncated in summary, expanded in full)
# ============================================================
$WatchedKeys = @(
    # Authentication — Access Permissions
    @{ Key='cas||securitypolicysids';                Label='Security Policy SIDs (raw)';       Section='auth'; IsXml=$false },
    # Authentication — External Auth
    @{ Key='cas||authequitracwcws';                  Label='Workstation Equitrac Auth';         Section='auth'; IsXml=$false;
       ValueMap=@{'0'='Disabled';'1'='Enabled with password prompt';'2'='Enabled without password prompt'} },
    @{ Key='cas||identityproviderwcws';              Label='Workstation Identity Provider';     Section='auth'; IsXml=$false;
       ValueMap=@{'0'='Disabled';'1'='Enabled'} },
    @{ Key='cas||useraccautosync';                   Label='LDAP Sync User Attrs on Login';     Section='auth'; IsXml=$false;
       ValueMap=@{'0'='Disabled';'1'='Enabled'} },
    @{ Key='cas||clientauthconfig';                  Label='Auth Method Config (raw XML)';      Section='auth'; IsXml=$true  },
    @{ Key='ads||settingsdoc';                       Label='AD Sync Settings (raw XML)';        Section='dirsync'; IsXml=$true  },
    # Authentication — Device Clients
    @{ Key='dce||enableswipe';                       Label='Allow Card Swipe Login';            Section='auth'; IsXml=$false;
       ValueMap=@{'0'='Off';'1'='On'} },
    @{ Key='dce||nosecondaryidwithswipe';            Label='Req. Extra Auth with Swipe';        Section='auth'; IsXml=$false;
       ValueMap=@{'0'='Enabled';'1'='Disabled'} },
    @{ Key='dce||registerpin';                       Label='Allow Keyboard Login';              Section='auth'; IsXml=$false;
       ValueMap=@{'0'='Off';'1'='On'} },
    @{ Key='dce||requirepasswordifavailablekeyboard';Label='Req. Password if PIN2 Available';   Section='auth'; IsXml=$false;
       ValueMap=@{'0'='No';'1'='Yes'} },
    @{ Key='dce||denypromptemptypassword';           Label='Deny Login with Empty Password';    Section='auth'; IsXml=$false;
       ValueMap=@{'0'='No';'1'='Yes'} },
    @{ Key='dce||defaultfunction';                   Label='Default Function at Device';        Section='auth'; IsXml=$false },
    # Authentication — Workstation / Web Client
    @{ Key='cas||caspin1pcmanager';                  Label='PIN1 in PC Manager';                Section='auth'; IsXml=$false },
    @{ Key='cas||caspin2pcmanager';                  Label='PIN2 in PC Manager';                Section='auth'; IsXml=$false },
    # Authentication — Card Registration
    @{ Key='dce||authequitraccardreg';               Label='Card Reg: Equitrac Auth';           Section='auth'; IsXml=$false;
       ValueMap=@{'0'='Off';'1'='On'} },
    @{ Key='dce||authidentityprovidercardreg';       Label='Card Reg: Identity Provider';       Section='auth'; IsXml=$false;
       ValueMap=@{'0'='Off';'1'='On'} },
    @{ Key='dce||registerpinasalternate';            Label='Card Storage: Alternate PIN';       Section='auth'; IsXml=$false;
       ValueMap=@{'0'='Store as Primary PIN';'1'='Store as Alternate PIN'} },
    @{ Key='dce||registertwocards';                  Label='Dual Card Reader Mode';             Section='auth'; IsXml=$false;
       ValueMap=@{'0'='Disabled';'1'='Enabled'} },
    # Authentication — Encryption
    @{ Key='cas||encryptsecondarypin';               Label='Store Secondary PIN Encrypted';     Section='auth'; IsXml=$false;
       ValueMap=@{'N'='Disabled';'Y'='Enabled'} },
    # Authentication — Card Setup
    @{ Key='cas||casstartpinposition';               Label='Card Data: From Position';          Section='auth'; IsXml=$false },
    @{ Key='cas||casendpinposition';                 Label='Card Data: To Position';            Section='auth'; IsXml=$false },
    @{ Key='dce||hidparameters';                     Label='HID Decoding Parameters (raw XML)'; Section='auth'; IsXml=$true  },
    # Authentication — CAS Offline Behavior
    @{ Key='dce||cacheuserloginofflinettldays';      Label='Login Cache TTL (days, 0=disabled)';Section='auth'; IsXml=$false },
    @{ Key='dce||supportofflinemode';                Label='Login Caching';                     Section='auth'; IsXml=$false;
       ValueMap=@{'0'='Disabled';'1'='Enabled'} },
    @{ Key='cas||casdownaction';                     Label='Print Behavior';                    Section='auth'; IsXml=$false;
       ValueMap=@{'PrintChargeLater'='Print, charge accounts later';'AutoSelect'='Auto select';'Auto'='Auto select';'DoNotPrint'='Do not print'} },
    # Authentication — Misc
    @{ Key='dce||adminpin';                          Label='Admin PIN';                         Section='auth'; IsXml=$false },
    @{ Key='dce||maxpinlength';                      Label='Max PIN Length';                    Section='auth'; IsXml=$false },
    @{ Key='cas||loginexpiry';                       Label='Login Expiry (seconds)';            Section='auth'; IsXml=$false },
    @{ Key='cas||pin1minlen';                        Label='PIN1 Min Length';                   Section='auth'; IsXml=$false },
    @{ Key='cas||pin1maxlen';                        Label='PIN1 Max Length';                   Section='auth'; IsXml=$false },
    # Job Management
    @{ Key='dre|sdr|jobexpirytime';              Label='Job Expiry Time (minutes)';    Section='job';      IsXml=$false },
    @{ Key='dre|sdr|distributionlistjobexpirytime'; Label='Dist. List Expiry (minutes)'; Section='job';   IsXml=$false },
    @{ Key='cas||precision';                    Label='Accounting Precision';         Section='job';      IsXml=$false },
    @{ Key='dce||offlinelifetime';              Label='Offline Lifetime';             Section='job';      IsXml=$false },
    @{ Key='dce||requeuereleasedjobsonlogout';  Label='Requeue Jobs on Logout';       Section='job';      IsXml=$false },
    @{ Key='dce||releasebehaviour';             Label='Release Behaviour';            Section='job';      IsXml=$false },
    @{ Key='cas||escrowcfg';                    Label='Escrow Config';                Section='job';      IsXml=$true  },
    # Quotas and Messages
    @{ Key='cas||colourquota';                  Label='Color Quota Mode';             Section='quota';    IsXml=$false;
       ValueMap=@{'0'='Disabled';'1'='Enabled (track only)';'2'='Enabled + block color copy on exceed';'3'='Enabled (custom message only)'} },
    @{ Key='cas||autousercolorquotalimit';       Label='Auto User Color Quota Limit';  Section='quota';    IsXml=$false },
    @{ Key='dce||facaccesscolorcopy';           Label='Color Copy Feature Access';    Section='quota';    IsXml=$false;
       ValueMap=@{'0'='Restricted';'1'='Unrestricted'} },
    @{ Key='cas||accenforcelimit';              Label='Color Quota Account Hard Limit'; Section='quota';  IsXml=$false },
    @{ Key='cas||insufficientfundsmsg';         Label='Insufficient Funds Message';   Section='quota';    IsXml=$false },
    @{ Key='dre||colorquotamessage';            Label='Color Quota Message';          Section='quota';    IsXml=$false },
    # Currency and Accounting
    @{ Key='cas||currencyiso4217';              Label='Currency (ISO 4217)';          Section='currency'; IsXml=$false },
    @{ Key='dce||costpreview';                  Label='Cost Preview Mode';            Section='currency'; IsXml=$false },
    @{ Key='dce||colourmultiplier';             Label='Colour Multiplier';            Section='currency'; IsXml=$false },
    @{ Key='dce||oversizemultiplier';           Label='Oversize Multiplier';          Section='currency'; IsXml=$false },
    @{ Key='dce||displaybalanceinfo';           Label='Display Balance Info';         Section='currency'; IsXml=$false },
    @{ Key='dce||displaycostinfo';              Label='Display Cost Info';            Section='currency'; IsXml=$false },
    @{ Key='dce||chargebeforecopying';          Label='Charge Before Copying';        Section='currency'; IsXml=$false },
    # License Server
    @{ Key='cas||fneserverhost';                Label='License Server Host';          Section='license';  IsXml=$false },
    @{ Key='cas||fneserverport';                Label='License Server Port';          Section='license';  IsXml=$false },
    @{ Key='cas||fneserverprotocol';            Label='License Server Protocol';      Section='license';  IsXml=$false },
    # Device Settings
    @{ Key='dce||defaultpagesize';              Label='Default Page Size';            Section='device';   IsXml=$false },
    @{ Key='dce||copiertimeout';                Label='Copier Timeout (ms)';          Section='device';   IsXml=$false },
    @{ Key='dce||enablekeypad';                 Label='Enable Keypad';                Section='device';   IsXml=$false },
    @{ Key='dce||deviceconnecttimeout';         Label='Device Connect Timeout (ms)';  Section='device';   IsXml=$false },
    @{ Key='dce||enablebillablefeature';        Label='Billable Feature Enabled';     Section='device';   IsXml=$false },
    @{ Key='dce||displayaccountinfo';           Label='Display Account Info';         Section='device';   IsXml=$false },
    @{ Key='dce||promptforbillingcode';         Label='Prompt for Billing Code';      Section='device';   IsXml=$false }
)

# Section display names (ordered)
$SectionMeta = [ordered]@{
    auth     = 'Authentication'
    dirsync  = 'Directory Services Synchronization'
    network  = 'Global Network Settings'
    job      = 'Job Management'
    quota    = 'Quotas and Messages'
    currency = 'Currency and Accounting'
    license  = 'License Server'
    device   = 'Device Settings'
}

# EQVar keys that change every service cycle - suppress from diffs
$EQVarNoise = @(
    'cas||workflowfolderslastupdatetime',
    'dce||workflowfolderslastupdatetime',
    'dce||casconfigorworkflowslastupdatetime',
    'cas||lastupdatetime',
    'dce||lastupdatetime',
    'cas||servicestarttimestamp',
    'dce||servicestarttimestamp'
)

# Workflow type map
$WfTypeMap = @{
    '5'='Copy'; '6'='Scan to Email'; '7'='Fax'; '8'='Print to Me';
    '9'='Release All'; '10'='Scan to Folder'; '11'='Scan to FTP'; '12'='Scan to USB'
}

# ============================================================
# PATHS
# ============================================================
$hostname   = $env:COMPUTERNAME
$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputRoot = Join-Path $scriptDir 'Output'
$TmpDir     = 'C:\Temp\EQ_AIS_Tmp'

function Ensure-Dir { param([string]$p)
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
}

# ============================================================
# INTERACTIVE MENU
# ============================================================
function Show-Menu {
    Clear-Host
    Write-Host ''
    Write-Host '  ================================================' -ForegroundColor Cyan
    Write-Host "    $AppName Config Scanner" -ForegroundColor White
    Write-Host "    Server : $hostname" -ForegroundColor Gray
    Write-Host '  ================================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '    1.  BEFORE   Capture baseline (before UI changes)' -ForegroundColor Yellow
    Write-Host '    2.  AFTER    Capture current state + HTML report'   -ForegroundColor Green
    Write-Host '    3.  COMPARE  Diff latest Before vs After'           -ForegroundColor Magenta
    Write-Host '    4.  FULL     After + Compare in one run'            -ForegroundColor Cyan
    Write-Host ''
    Write-Host '    Q.  Quit' -ForegroundColor DarkGray
    Write-Host ''
    $c = Read-Host '  Select [1-4, Q]'
    switch ($c.Trim().ToUpper()) {
        '1' { return 'Before'  }
        '2' { return 'After'   }
        '3' { return 'Compare' }
        '4' { return 'Full'    }
        'Q' { Write-Host 'Bye.' -ForegroundColor Gray; exit 0 }
        default { Write-Host "  Invalid choice '$c'" -ForegroundColor Red; return Show-Menu }
    }
}

# ============================================================
# UTILITIES
# ============================================================
function HE {
    param([string]$s)
    if ($s -eq $null) { return '' }
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function FV {
    param([string]$v, [int]$Max = 200)
    if ($v -eq $null -or [string]::IsNullOrWhiteSpace($v)) { return '(not set)' }
    $v = $v.Trim()
    if ($v -eq '') { return '(not set)' }
    if ($v.Length -gt $Max) { return $v.Substring(0, $Max) + '...' }
    return $v
}

function XV {
    param([string]$xml, [string]$tag)
    if ([string]::IsNullOrEmpty($xml)) { return '' }
    $m = [regex]::Match($xml, "<$tag>([^<]*)</$tag>")
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return ''
}

function Get-BcpLines {
    param([string]$Table)
    $f = "$TmpDir\bcp_$Table.txt"
    if (Test-Path $f) { Remove-Item $f -Force }
    $a = @("$SqlDb..$Table", 'out', $f, '-S', $SqlServer, '-U', $SqlUser, '-P', $SqlPass, '-c', '-t', '|')
    & bcp @a 2>$null | Out-Null
    if (-not (Test-Path $f)) { return @() }
    return [System.IO.File]::ReadAllLines($f)
}

function Get-EQVarMap {
    param([string]$DbPath)
    $map = @{}
    if (-not (Test-Path $DbPath)) { return $map }
    $tmp = "$TmpDir\eq_q.db3"
    Copy-Item $DbPath $tmp -Force
    if (Test-Path "$DbPath-wal") { Copy-Item "$DbPath-wal" "$tmp-wal" -Force }
    if (Test-Path "$DbPath-shm") { Copy-Item "$DbPath-shm" "$tmp-shm" -Force }
    $q = 'SELECT SubSystem,Class,Name,Value FROM EQVar ORDER BY SubSystem,Class,Name;'
    (& $Sqlite3 -separator "`t" $tmp $q 2>$null) | ForEach-Object {
        $parts = $_ -split "`t", 4
        if ($parts.Count -ge 3) {
            $key = "$($parts[0])|$($parts[1])|$($parts[2])"
            $map[$key] = if ($parts.Count -ge 4) { $parts[3] } else { '' }
        }
    }
    Remove-Item "$tmp*" -Force -ErrorAction SilentlyContinue
    return $map
}

function Load-TsvMap {
    param([string]$Path)
    $map = @{}
    if (-not (Test-Path $Path)) { return $map }
    foreach ($line in [System.IO.File]::ReadAllLines($Path)) {
        $parts = $line -split "`t", 4
        if ($parts.Count -lt 3) { continue }
        $key = "$($parts[0])|$($parts[1])|$($parts[2])"
        $map[$key] = if ($parts.Count -ge 4) { $parts[3] } else { '' }
    }
    return $map
}

# ============================================================
# WINDOWS DATA COLLECTION
# ============================================================
function Collect-WindowsData {
    Write-Host '  [W1/7] System identity...'      -ForegroundColor Cyan
    $cs   = Get-CimInstance Win32_ComputerSystem        -EA SilentlyContinue
    $csp  = Get-CimInstance Win32_ComputerSystemProduct -EA SilentlyContinue
    $bios = Get-CimInstance Win32_BIOS                  -EA SilentlyContinue
    $mfr  = if ($cs)  { $cs.Manufacturer }        else { '' }
    $model= if ($cs)  { $cs.Model }               else { '' }
    $uuid = if ($csp) { $csp.UUID }               else { '' }
    $biosVer    = if ($bios) { $bios.SMBIOSBIOSVersion } else { '' }
    $biosSerial = if ($bios) { $bios.SerialNumber }       else { '' }
    $platform = 'Physical Server'; $hypervisor = ''
    $vmChecks = @(
        @{P='VMware';          N='VMware';       F=@($mfr,$model,$biosVer)},
        @{P='Virtual Machine'; N='Hyper-V';      F=@($model)},
        @{P='Microsoft.*Hyper';N='Hyper-V';      F=@($mfr,$model)},
        @{P='VirtualBox';      N='VirtualBox';   F=@($mfr,$model)},
        @{P='innotek';         N='VirtualBox';   F=@($mfr)},
        @{P='QEMU';            N='KVM/QEMU';     F=@($mfr,$model,$biosVer)},
        @{P='KVM';             N='KVM/QEMU';     F=@($model)},
        @{P='Xen';             N='Xen';          F=@($mfr,$model,$biosVer)},
        @{P='Parallels';       N='Parallels';    F=@($mfr,$model)},
        @{P='Amazon EC2';      N='AWS EC2';      F=@($mfr,$model)},
        @{P='Google';          N='Google Cloud'; F=@($mfr,$model)}
    )
    foreach ($vc in $vmChecks) {
        foreach ($f in $vc.F) { if ($f -match $vc.P) { $platform='Virtual Machine'; $hypervisor=$vc.N; break } }
        if ($hypervisor) { break }
    }
    if (-not $hypervisor -and (Test-Path 'HKLM:\SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters' -EA SilentlyContinue)) {
        $platform='Virtual Machine'; $hypervisor='Hyper-V'
    }

    Write-Host '  [W2/7] OS and hardware...'     -ForegroundColor Cyan
    $os  = Get-CimInstance Win32_OperatingSystem -EA SilentlyContinue
    $cpu = Get-CimInstance Win32_Processor       -EA SilentlyContinue | Select-Object -First 1
    $ram = if ($cs) { [math]::Round($cs.TotalPhysicalMemory/1GB,1) } else { 0 }
    $osInstall = if ($os -and $os.InstallDate) { $os.InstallDate.ToString('yyyy-MM-dd') } else { '' }
    $drives = [System.Collections.Generic.List[object]]::new()
    Get-PSDrive -PSProvider FileSystem -EA SilentlyContinue | Where-Object { $_.Used -ne $null } | ForEach-Object {
        $drives.Add([PSCustomObject]@{
            Name  = $_.Name
            Total = [math]::Round(($_.Used+$_.Free)/1GB,1)
            Used  = [math]::Round($_.Used/1GB,1)
            Free  = [math]::Round($_.Free/1GB,1)
        })
    }

    Write-Host '  [W3/7] Network...'             -ForegroundColor Cyan
    $nbMap   = @{0='Default (via DHCP)';1='Enabled';2='Disabled'}
    $adpObjs = Get-CimInstance Win32_NetworkAdapter -EA SilentlyContinue | Where-Object { $_.NetEnabled -ne $null }
    $nics    = [System.Collections.Generic.List[object]]::new()
    Get-CimInstance Win32_NetworkAdapterConfiguration -EA SilentlyContinue | Where-Object { $_.IPEnabled } | ForEach-Object {
        $a    = $_
        $phys = $adpObjs | Where-Object { $_.InterfaceIndex -eq $a.InterfaceIndex } | Select-Object -First 1
        $ips  = if ($a.IPAddress)           { ($a.IPAddress  | Where-Object { $_ -notmatch ':' }) -join ', ' } else { '' }
        $subs = if ($a.IPSubnet)            { ($a.IPSubnet   | Where-Object { $_ -notmatch ':' }) -join ', ' } else { '' }
        $gw   = if ($a.DefaultIPGateway)    { $a.DefaultIPGateway -join ', ' }    else { '' }
        $dns  = if ($a.DNSServerSearchOrder){ $a.DNSServerSearchOrder -join ', ' } else { '' }
        $dhcp = if ($a.DHCPEnabled)         { "Yes (server: $($a.DHCPServer))" }  else { 'No (static)' }
        $nbRaw= if ($a.TcpipNetbiosOptions -ne $null) { [int]$a.TcpipNetbiosOptions } else { -1 }
        $nb   = if ($nbMap.ContainsKey($nbRaw)) { $nbMap[$nbRaw] } else { "Unknown ($nbRaw)" }
        $spd  = if ($phys -and $phys.Speed) { "$([math]::Round($phys.Speed/1MB)) Mbps" } else { '' }
        $conn = if ($phys) { switch ($phys.NetConnectionStatus) { 2{'Connected'} 7{'Disconnected'} default{"Code $($phys.NetConnectionStatus)"} } } else { '' }
        $nics.Add([PSCustomObject]@{
            Desc=$a.Description; Mac=$a.MACAddress; IP=$ips; Subnet=$subs
            Gateway=$gw; DNS=$dns; DHCP=$dhcp; NetBIOS=$nb; Speed=$spd; Status=$conn
            DNSDomain=$a.DNSDomain
        })
    }

    Write-Host '  [W4/7] Installed software...'  -ForegroundColor Cyan
    $allReg  = Get-ItemProperty @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    ) -EA SilentlyContinue | Where-Object { $_.DisplayName }
    $allSvcs = Get-Service -EA SilentlyContinue
    $pmDefs  = @(
        @{L='PaperCut MF/NG';           R=@('*PaperCut MF*','*PaperCut NG*');              S=@('*PaperCut*');      D=@('C:\Program Files\PaperCut MF','C:\Program Files\PaperCut NG')},
        @{L='Equitrac / ControlSuite';  R=@('*Equitrac*','*ControlSuite*','*Kofax Control*','*Nuance Control*','*Tungsten*'); S=@('*Equitrac*','*ControlSuite*','*Kofax*','*Nuance*','*Tungsten*'); D=@('C:\Program Files\Kofax\ControlSuite','C:\Program Files\Nuance\ControlSuite','C:\Program Files\Equitrac')},
        @{L='YSoft SafeQ';              R=@('*SafeQ*','*YSoft*');                           S=@('*SafeQ*','*YSoft*');D=@('C:\SafeQ6','C:\SafeQ5','C:\Program Files\Y Soft')},
        @{L='AWMS2';                    R=@('*AWMS*','*ApeosWare Management*');              S=@('*AWMS*','*ApeosWare*'); D=@('C:\Program Files\FujiFilm\AWMS','C:\AWMS2')}
    )
    $pmApps=[System.Collections.Generic.List[object]]::new()
    foreach ($def in $pmDefs) {
        $fnd=$false; $ver=''; $pth=''; $sst=''
        foreach ($pat in $def.R) { $r=$allReg|Where-Object{$_.DisplayName -like $pat}|Select-Object -First 1; if($r){$fnd=$true;$ver=$r.DisplayVersion;$pth=$r.InstallLocation;break} }
        foreach ($pat in $def.S) { $s=$allSvcs|Where-Object{$_.DisplayName -like $pat}|Select-Object -First 1; if($s){$fnd=$true;$sst=[string]$s.Status;break} }
        if (-not $fnd) { foreach ($d in $def.D) { if(Test-Path $d -EA SilentlyContinue){$fnd=$true;if(-not $pth){$pth=$d};break} } }
        if ($fnd) { $pmApps.Add([PSCustomObject]@{Name=$def.L;Version=$ver;Path=$pth;SvcStatus=$sst}) }
    }
    foreach ($kw in @('PrinterLogic','Printix','Pharos','UniFlow','ThinPrint','MyQ','PrintManager')) {
        $m=$allReg|Where-Object{$_.DisplayName -like "*$kw*"}|Select-Object -First 1
        if ($m) { $pmApps.Add([PSCustomObject]@{Name=$m.DisplayName;Version=$m.DisplayVersion;Path='';SvcStatus=''}) }
    }

    Write-Host '  [W5/7] Windows roles and features...' -ForegroundColor Cyan
    $printRole=''; $spoolerSt=''; $dotNet35=''; $dotNet48=''
    $roles=[System.Collections.Generic.List[string]]::new()
    $feats=[System.Collections.Generic.List[string]]::new()
    try {
        $allFeat = Get-WindowsFeature -EA Stop
        function GF($n){$allFeat|Where-Object{$_.Name -eq $n}|Select-Object -First 1}
        $printRole = if ((GF 'Print-Server').Installed)          { 'Installed' } else { 'Not Installed' }
        $dotNet35  = if ((GF 'NET-Framework-Core').Installed)    { 'Installed' } else { 'Not Installed' }
        $dotNet48  = if ((GF 'NET-Framework-45-Core').Installed) { 'Installed' } else { 'Not Installed' }
        $allFeat | Where-Object { $_.Installed -and $_.FeatureType -eq 'Role'    } | ForEach-Object { $roles.Add($_.DisplayName) }
        $allFeat | Where-Object { $_.Installed -and $_.FeatureType -eq 'Feature' } | ForEach-Object { $feats.Add($_.DisplayName) }
    } catch {}
    $spooler  = Get-Service Spooler -EA SilentlyContinue
    $spoolerSt= if ($spooler) { "$($spooler.Status)  StartType=$($spooler.StartType)" } else { 'Not found' }

    Write-Host '  [W6/7] SQL Server...'          -ForegroundColor Cyan
    $sqlInsts=[System.Collections.Generic.List[object]]::new()
    $sqlKey='HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL'
    if (Test-Path $sqlKey -EA SilentlyContinue) {
        $ip = Get-ItemProperty $sqlKey -EA SilentlyContinue
        if ($ip) {
            $ip.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object {
                $iname=$_.Name; $folder=$_.Value
                $sk  = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$folder\Setup"
                $msk = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$folder\MSSQLServer"

                $ed          = try{(Get-ItemProperty $sk -EA Stop).Edition}catch{''}
                $ver2        = try{(Get-ItemProperty $sk -EA Stop).Version}catch{''}
                $installPath = try{(Get-ItemProperty $sk -EA Stop).SQLPath}catch{''}
                $dataRoot    = try{(Get-ItemProperty $sk -EA Stop).SQLDataRoot}catch{''}

                $svcN    = if($iname -eq 'MSSQLSERVER'){'MSSQLSERVER'}else{"MSSQL`$$iname"}
                $svc2    = Get-Service $svcN -EA SilentlyContinue
                $svcSt2  = if($svc2){[string]$svc2.Status}else{'Not running'}
                $svcAcct = try{(Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\$svcN" -EA Stop).ObjectName}catch{''}

                $loginMode = try{[int](Get-ItemProperty $msk -EA Stop).LoginMode}catch{0}
                $authMode  = switch($loginMode){1{'Windows Only'} 2{'Mixed Mode (SQL + Windows)'} default{"Unknown ($loginMode)"}}

                $backupDir = try{(Get-ItemProperty $msk -EA Stop).BackupDirectory}catch{''}
                $defData   = try{(Get-ItemProperty $msk -EA Stop).DefaultData}catch{''}
                $defLog    = try{(Get-ItemProperty $msk -EA Stop).DefaultLog}catch{''}

                $errLog = ''
                if (Test-Path "$msk\Parameters" -EA SilentlyContinue) {
                    (Get-ItemProperty "$msk\Parameters" -EA SilentlyContinue).PSObject.Properties |
                        Where-Object { $_.Name -match '^SqlArg' -and $_.Value -match '^-e(.+)' } |
                        Select-Object -First 1 | ForEach-Object { $errLog = $_.Value -replace '^-e','' }
                }

                $npKey    = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$folder\MSSQLServer\SuperSocketNetLib\Np"
                $npEnabled = (Test-Path $npKey -EA SilentlyContinue) -and ((try{(Get-ItemProperty $npKey -EA Stop).Enabled}catch{0}) -eq 1)

                $tcpKey = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$folder\MSSQLServer\SuperSocketNetLib\Tcp"
                $tcpEn=$false; $tcpPt=''; $tcpKeepAlive=''; $tcpListenAll=''
                $tcpIpList=[System.Collections.Generic.List[object]]::new()
                if (Test-Path $tcpKey -EA SilentlyContinue) {
                    $tcpP = Get-ItemProperty $tcpKey -EA SilentlyContinue
                    $tcpEn        = ($tcpP.Enabled -eq 1)
                    $tcpKeepAlive = "$($tcpP.KeepAlive)"
                    $tcpListenAll = if($tcpP.ListenOnAllIPs -eq 1){'Yes'}else{'No'}
                    $ipAllKey = "$tcpKey\IPAll"
                    if (Test-Path $ipAllKey -EA SilentlyContinue) {
                        $ia = Get-ItemProperty $ipAllKey -EA SilentlyContinue
                        $tcpPt = if($ia.TcpPort){$ia.TcpPort}elseif($ia.TcpDynamicPorts){"dyn:$($ia.TcpDynamicPorts)"}else{''}
                    }
                    Get-ChildItem $tcpKey -EA SilentlyContinue | Where-Object { $_.PSChildName -match '^IP\d+$' } | ForEach-Object {
                        $ipP = Get-ItemProperty $_.PSPath -EA SilentlyContinue
                        if ($ipP -and $ipP.IpAddress) {
                            $tcpIpList.Add([PSCustomObject]@{
                                Name    = $_.PSChildName
                                Active  = if($ipP.Active  -eq 1){'Yes'}else{'No'}
                                Enabled = if($ipP.Enabled -eq 1){'Yes'}else{'No'}
                                IP      = $ipP.IpAddress
                                Port    = if($ipP.TcpPort){$ipP.TcpPort}elseif($ipP.TcpDynamicPorts){"dyn:$($ipP.TcpDynamicPorts)"}else{''}
                            })
                        }
                    }
                }

                $fsKey   = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$folder\MSSQLSERVER\Filestream"
                $fsLevel = try{[int](Get-ItemProperty $fsKey -EA Stop).EnableLevel}catch{-1}
                $fsDesc  = switch($fsLevel){
                    0{'Disabled'} 1{'T-SQL access only'} 2{'T-SQL + local file system'} 3{'T-SQL + local + remote file system'} default{'Not configured'}
                }

                $sqlSrv = if($iname -eq 'MSSQLSERVER'){'.'}else{".\$iname"}
                $sqA = @('-S',$sqlSrv,'-U',$SqlUser,'-P',$SqlPass,'-W','-s','|','-h','-1')

                $maxMem=''; $minMem=''; $userInst=''
                & sqlcmd @sqA -Q "SET NOCOUNT ON; SELECT name+'|'+CAST(value_in_use AS VARCHAR(20)) FROM sys.configurations WHERE name IN ('max server memory (MB)','min server memory (MB)','user instances enabled') ORDER BY name;" 2>$null | ForEach-Object {
                    $t=$_.Trim()
                    if ($t -match '\|') {
                        $kk,$vv=$t -split '\|',2
                        switch -Wildcard ($kk.Trim()) {
                            'max*' { $maxMem  = "$($vv.Trim()) MB" }
                            'min*' { $minMem  = "$($vv.Trim()) MB" }
                            'user*'{ $userInst = if($vv.Trim() -eq '1'){'Enabled'}else{'Disabled'} }
                        }
                    }
                }

                $sysAdmins=[System.Collections.Generic.List[string]]::new()
                & sqlcmd @sqA -Q "SET NOCOUNT ON; SELECT sp.name+'|'+sp.type_desc FROM sys.server_role_members srm JOIN sys.server_principals sp ON srm.member_principal_id=sp.principal_id JOIN sys.server_principals r ON srm.role_principal_id=r.principal_id WHERE r.name='sysadmin';" 2>$null | ForEach-Object {
                    $t=$_.Trim(); if($t -and $t -match '\|'){$sysAdmins.Add($t)}
                }

                $tempdbFiles=[System.Collections.Generic.List[object]]::new()
                & sqlcmd @sqA -Q "SET NOCOUNT ON; SELECT name+'|'+type_desc+'|'+physical_name+'|'+CAST(size*8/1024 AS VARCHAR)+'|'+CASE WHEN is_percent_growth=1 THEN CAST(growth AS VARCHAR)+'%' ELSE CAST(growth*8/1024 AS VARCHAR)+' MB' END FROM sys.master_files WHERE database_id=2;" 2>$null | ForEach-Object {
                    $t=$_.Trim()
                    if ($t -and $t -match '\|') {
                        $p=$t -split '\|',5
                        if($p.Count -ge 5){$tempdbFiles.Add([PSCustomObject]@{Name=$p[0];Type=$p[1];Path=$p[2];SizeMB=$p[3];Growth=$p[4]})}
                    }
                }

                $sqlInsts.Add([PSCustomObject]@{
                    Name=if($iname -eq 'MSSQLSERVER'){'Default (MSSQLSERVER)'}else{$iname}
                    Edition=$ed; Version=$ver2; Status=$svcSt2; ServiceAcct=$svcAcct; AuthMode=$authMode
                    InstallPath=$installPath; DataRoot=$dataRoot; DefaultData=$defData; DefaultLog=$defLog
                    BackupDir=$backupDir; ErrorLog=$errLog; NpEnabled=$npEnabled
                    TcpEnabled=$tcpEn; TcpPort=$tcpPt; TcpKeepAlive=$tcpKeepAlive; TcpListenAll=$tcpListenAll
                    TcpIpAddrs=$tcpIpList; FsDesc=$fsDesc
                    MaxMemoryMB=$maxMem; MinMemoryMB=$minMem; UserInstances=$userInst
                    SysAdmins=$sysAdmins; TempDbFiles=$tempdbFiles
                })
            }
        }
    }

    Write-Host '  [W7/7] Print queues and drivers...' -ForegroundColor Cyan
    $spoolFld = try{(Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -Name DefaultSpoolDirectory -EA Stop).DefaultSpoolDirectory}catch{'(default)'}
    $printers  = Get-Printer -EA SilentlyContinue
    $wmiPrt    = Get-CimInstance Win32_Printer -EA SilentlyContinue
    $drvX64    = @(Get-PrinterDriver -PrinterEnvironment 'Windows x64'    -EA SilentlyContinue)
    $drvX86    = @(Get-PrinterDriver -PrinterEnvironment 'Windows NT x86' -EA SilentlyContinue)
    $x86N      = $drvX86 | ForEach-Object { $_.Name }
    $icmM      = @{1='ICM Disabled';2='ICM by OS';3='ICM by device';4='ICM by host'}
    $RAWONLY   = 0x00001000; $PUBL=0x00002000; $RCLI=0x00040000
    $pQueues   = [System.Collections.Generic.List[object]]::new()
    foreach ($p in $printers) {
        $wmi2  = $wmiPrt|Where-Object{$_.Name -eq $p.Name}|Select-Object -First 1
        $attrs = if($wmi2){[int]$wmi2.Attributes}else{0}
        $cfg2  = Get-PrintConfiguration -PrinterName $p.Name -EA SilentlyContinue
        $stTxt='Unknown'; $ofTxt='Unknown'; $cmTxt='Unknown'
        if ($cfg2 -and $cfg2.PrintTicketXML) {
            try {
                [xml]$pt2=$cfg2.PrintTicketXML
                $ns2=@{psf='http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework';psk='http://schemas.microsoft.com/windows/2003/08/printing/printschemakeywords'}
                $cmN=Select-Xml -Xml $pt2 -XPath "//psf:Feature[contains(@name,'PageColorManagement')]//psf:Option" -Namespace $ns2|Select-Object -First 1
                if($cmN){$v2=$cmN.Node.GetAttribute('name');$cmTxt=switch -Wildcard($v2){'*:None'{'On (pass-through)'}'*:System'{'Off (Windows ICM)'}'*:Driver'{'Off (Driver ICM)'}'*:Device'{'Off (Device ICM)'}default{$v2}}}
                $stN=Select-Xml -Xml $pt2 -XPath "//psf:Feature[contains(@name,'Staple') or contains(@name,'staple') or contains(@name,'Finishing')]//psf:Option" -Namespace $ns2|Select-Object -First 1
                if($stN){$v2=$stN.Node.GetAttribute('name');$stTxt=if($v2 -like '*None*'-or $v2 -like '*none*'){'Disabled'}else{$v2 -replace '^.*:',''}}
                $bnN=Select-Xml -Xml $pt2 -XPath "//psf:Feature[contains(@name,'OutputBin') or (contains(@name,'Bin') and not(contains(@name,'Input')))]//psf:Option" -Namespace $ns2|Select-Object -First 1
                if($bnN){$v2=$bnN.Node.GetAttribute('name');$ofTxt=if($v2 -like '*OffsetEachSet*'){'Offset/Set'}elseif($v2 -like '*OffsetEachJob*'){'Offset/Job'}else{$v2 -replace '^.*:',''}}
            } catch {}
        }
        $icmTxt='N/A'
        try{$dmB=(Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$($p.Name)" -Name 'Default DevMode' -EA Stop).'Default DevMode';if($dmB -and $dmB.Count-ge 192){$r=[BitConverter]::ToInt32($dmB,188);$icmTxt=if($icmM.ContainsKey($r)){$icmM[$r]}else{"Method $r"}}}catch{}
        $dxTxt='Unknown'
        if($cfg2 -and $cfg2.DuplexingMode -ne $null){$dxTxt=switch([string]$cfg2.DuplexingMode){'OneSided'{'One-sided'}'TwoSidedLongEdge'{'Two-sided Long Edge'}'TwoSidedShortEdge'{'Two-sided Short Edge'}default{[string]$cfg2.DuplexingMode}}}
        $clrTxt=if($cfg2){if($cfg2.Color){'Color'}else{'Monochrome'}}else{'Unknown'}
        $ppTxt =if($cfg2 -and $cfg2.PaperSize){[string]$cfg2.PaperSize}else{'Unknown'}
        $finTxt=''; $staI=''; $pnI=''; $bkI=''; $ofI=''
        $pp2=$null; try{$pp2=Get-PrinterProperty -PrinterName $p.Name -EA Stop}catch{}
        if($pp2 -and $pp2.Count-gt 0){
            $pm=@{};$pp2|ForEach-Object{$pm[$_.PropertyName]=$_.Value}
            $fl=@('A','B','C','D')|ForEach-Object{$v2=$pm["Config:OP_Finisher$_"];if($v2 -and $v2-ne'No'){"Fin$_ ($v2)"}}
            $finTxt=if($fl){$fl -join ', '}else{'None'}
            $staI=if($pm['Config:DC_FIN_Staple']-eq'Yes'){'Yes'}else{'No'}
            $pnI =if($pm['Config:DC_FIN_Punch'] -eq'Yes'){'Yes'}else{'No'}
            $bkI =if($pm['Config:OP_Booklet']-and $pm['Config:OP_Booklet']-ne'No'){$pm['Config:OP_Booklet']}else{'No'}
            $ofI =if($pm['Config:DC_OffsetStacking']-eq'Yes'){'Yes'}else{'No'}
        } elseif($p.DriverName -like 'FF *' -or $p.DriverName -like '*FujiFilm*' -or $p.DriverName -like '*Apeos*'){
            $finTxt='N/A (virtual)';$staI='N/A';$pnI='N/A';$bkI='N/A';$ofI='N/A'
        }
        $pQueues.Add([PSCustomObject]@{
            Name=$p.Name; Driver=$p.DriverName; Port=$p.PortName; Status=[string]$p.PrinterStatus
            AdvFeatures=if(($attrs-band $RAWONLY)-eq 0){'Enabled'}else{'Disabled (RAW only)'}
            RenderOnClient=if(($attrs-band $RCLI)-ne 0){'Yes'}else{'No'}
            PaperSize=$ppTxt; Duplex=$dxTxt; Color=$clrTxt; Staple=$stTxt; OffsetStack=$ofTxt
            ColorMgmt=$cmTxt; IcmMethod=$icmTxt
            Shared=$p.Shared; ShareName=if($p.Shared){$p.ShareName}else{''}
            Published=($attrs-band $PUBL)-ne 0; Has32bit=$x86N -contains $p.DriverName
            FinisherInst=$finTxt; StapleInst=$staI; PunchInst=$pnI; BookletInst=$bkI; OffsetInst=$ofI
        })
    }
    $pPorts=[System.Collections.Generic.List[object]]::new()
    Get-PrinterPort -EA SilentlyContinue | Sort-Object Name | ForEach-Object {
        $pPorts.Add([PSCustomObject]@{Name=$_.Name;IP=$_.PrinterHostAddress;Port=$_.PortNumber;Type=$_.Description})
    }
    $ffCnt=($printers|Where-Object{$_.DriverName -like '*FujiFilm*'-or $_.DriverName -like '*Fuji Xerox*'-or $_.DriverName -like '*FUJIFILM*'-or $_.DriverName -like 'FF *'-or $_.DriverName -like '*Apeos*'}).Count

    return [PSCustomObject]@{
        ComputerName=$env:COMPUTERNAME; Manufacturer=$mfr; Model=$model
        Platform=$platform; Hypervisor=$hypervisor; UUID=$uuid
        BiosVersion=$biosVer; BiosSerial=$biosSerial
        OsName     =if($os){$os.Caption}else{''}
        OsVersion  =if($os){$os.Version}else{''}
        OsArch     =if($os){$os.OSArchitecture}else{''}
        OsInstall  =$osInstall
        Ram        ="$ram GB"
        CpuName    =if($cpu){$cpu.Name}else{''}
        CpuCores   =if($cpu){[string]$cpu.NumberOfCores}else{''}
        CpuThreads =if($cpu){[string]$cpu.NumberOfLogicalProcessors}else{''}
        Drives     =$drives; Nics=$nics
        DomainName =if($cs){$cs.Domain}else{''}
        DomainJoined=if($cs){$cs.PartOfDomain}else{$false}
        PmApps     =$pmApps
        PrintRoleStatus=$printRole; SpoolerStatus=$spoolerSt
        DotNet35=$dotNet35; DotNet48=$dotNet48
        InstalledRoles=$roles; InstalledFeatures=$feats
        SqlInstances=$sqlInsts; SpoolFolder=$spoolFld
        PrintQueues=$pQueues; DriversX64=$drvX64; DriversX86=$drvX86
        PrinterPorts=$pPorts; FfCount=$ffCnt
    }
}

function New-HtmlSection {
    param([string]$Id, [string]$Title, [string]$Body)
    $sid = "s-$Id"
    return "<div class='sec' id='$sid'>" +
           "<div class='sec-hdr' onclick=""tog('$sid')"">" +
           "<h2>$(HE $Title)</h2><span class='tog' id='t-$sid'>-</span></div>" +
           "<div class='sec-body' id='b-$sid'>$Body</div></div>"
}

function New-KVTable {
    param([System.Collections.Generic.List[object]]$Rows)
    $html = "<table class='kv'>"
    foreach ($r in $Rows) {
        $cls = if ($r.V -eq '(not set)' -or $r.V -eq '(empty)') { " e" } else { "" }
        $html += "<tr><td class='k'>$(HE $r.K)</td><td class='v$cls'>$(HE $r.V)</td></tr>"
    }
    $html += "</table>"
    return $html
}

# ============================================================
# DATA COLLECTION
# ============================================================
function Collect-Data {
    param([string]$SaveRawTo = '')

    Ensure-Dir $TmpDir

    Write-Host '  Collecting Windows server data...' -ForegroundColor White
    $winData = Collect-WindowsData

    Write-Host '  [1/5] SQLite EQVar databases...' -ForegroundColor Cyan
    $dcePath = "$DbBase\EQDCESrv\Cache\DCE_config.db3"
    $drePath = "$DbBase\EQDRESrv\EQSpool\DREEQVar.db3"
    $script:dce = Get-EQVarMap -DbPath $dcePath
    $script:dre = Get-EQVarMap -DbPath $drePath

    # Save raw TSVs for Compare if requested
    if ($SaveRawTo) {
        foreach ($pair in @(@{Src=$dcePath;Dst='eqvar_dce.tsv'},@{Src=$drePath;Dst='eqvar_dre.tsv'})) {
            if (Test-Path $pair.Src) {
                $tmp = "$TmpDir\raw_q.db3"
                Copy-Item $pair.Src $tmp -Force
                if (Test-Path "$($pair.Src)-wal") { Copy-Item "$($pair.Src)-wal" "$tmp-wal" -Force }
                # Replace embedded newlines so multi-line XML values don't break TSV line format
                $q = "SELECT SubSystem,Class,Name,replace(replace(Value,char(10),' '),char(13),'') FROM EQVar ORDER BY SubSystem,Class,Name;"
                & $Sqlite3 -separator "`t" $tmp $q 2>$null |
                    Out-File (Join-Path $SaveRawTo $pair.Dst) -Encoding ASCII
                Remove-Item "$tmp*" -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Write-Host '  [2/5] SQL table BCP dumps...' -ForegroundColor Cyan
    $bcpTables = @('cas_installedsoftware','cat_pricelist','cas_scan_alias',
                   'cas_workflow_folders','cas_pullgroups','cas_user_ext',
                   'cas_prq_device_ext','cas_config','cat_validation')
    $bcp = @{}
    foreach ($t in $bcpTables) {
        Write-Host "        $t" -ForegroundColor DarkCyan
        $bcp[$t] = Get-BcpLines -Table $t
        if ($SaveRawTo -and (Test-Path "$TmpDir\bcp_$t.txt")) {
            Copy-Item "$TmpDir\bcp_$t.txt" (Join-Path $SaveRawTo "bcp_$t.txt") -Force
        }
    }

    Write-Host '  [3/5] Parsing components...' -ForegroundColor Cyan
    $components = [System.Collections.Generic.List[object]]::new()
    foreach ($line in $bcp['cas_installedsoftware']) {
        $f = $line -split '\|'
        $comp = $f[0].Trim()
        if ($comp -eq '') { continue }
        $components.Add([PSCustomObject]@{
            Name    = $comp
            Desc    = if ($f.Count -gt 3) { $f[3].Trim() } else { '' }
            Version = if ($f.Count -gt 5) { $f[5].Trim() } else { '' }
            Date    = if ($f.Count -gt 6) { $f[6].Trim() } else { '' }
        })
    }

    Write-Host '  [4/5] Parsing price lists, workflows, users...' -ForegroundColor Cyan

    # Price lists
    $priceLists = [System.Collections.Generic.List[object]]::new()
    $cId=''; $cName=''; $cType=''; $cXml=[System.Text.StringBuilder]::new(); $inR=$false
    $FlushPL = {
        if ($inR -and $cName -ne '') {
            $xml  = $cXml.ToString()
            $rates = ([regex]::Matches($xml, 'rate="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }) |
                     Select-Object -Unique | Sort-Object { [double]$_ }
            $priceLists.Add([PSCustomObject]@{ Id=$cId; Name=$cName; Type=$cType; Rates=$rates })
        }
    }
    foreach ($line in $bcp['cat_pricelist']) {
        if ($line -match '^(\d+)\|') {
            & $FlushPL
            $f=$line -split '\|',5
            $cId=$f[0].Trim(); $cName=if($f.Count-gt 1){$f[1].Trim()}else{''}
            $cType=if($f.Count-gt 3){$f[3].Trim()}else{''}
            $cXml=[System.Text.StringBuilder]::new()
            if($f.Count-gt 4){[void]$cXml.Append($f[4])}
            $inR=$true
        } elseif ($inR) { [void]$cXml.AppendLine($line) }
    }
    & $FlushPL

    # Workflows
    $workflows = [System.Collections.Generic.List[object]]::new()
    foreach ($line in $bcp['cas_scan_alias']) {
        if (-not ($line -match '^\d+\|')) { continue }
        $f=$line -split '\|',5
        $st = XV (if($f.Count-gt 4){$f[4]}else{''}) 'scan_alias_type'
        $workflows.Add([PSCustomObject]@{
            Id       = $f[0].Trim()
            Name     = if($f.Count-gt 1){$f[1].Trim()}else{''}
            Type     = if($WfTypeMap.ContainsKey($st)){$WfTypeMap[$st]}elseif($st){"Type $st"}else{''}
            Modified = if($f.Count-gt 2){$f[2].Trim()}else{''}
        })
    }

    # Workflow folders
    $wfFolders = [System.Collections.Generic.List[object]]::new()
    foreach ($line in $bcp['cas_workflow_folders']) {
        $f=$line -split '\|'
        if($f[0].Trim()){$wfFolders.Add([PSCustomObject]@{Id=$f[0].Trim();Name=if($f.Count-gt 1){$f[1].Trim()}else{''}})}
    }

    # Pull groups
    $pullGroups = [System.Collections.Generic.List[object]]::new()
    foreach ($line in $bcp['cas_pullgroups']) {
        $f=$line -split '\|'
        if($f[0].Trim()){$pullGroups.Add([PSCustomObject]@{Id=$f[0].Trim();Name=if($f.Count-gt 1){$f[1].Trim()}else{''}})}
    }

    # Users
    $users = [System.Collections.Generic.List[object]]::new()
    foreach ($line in $bcp['cas_user_ext']) {
        $f=$line -split '\|'
        if($f[0].Trim()){$users.Add([PSCustomObject]@{Id=$f[0].Trim();Domain=if($f.Count-gt 1){$f[1].Trim()}else{''}; Name=if($f.Count-gt 2){$f[2].Trim()}else{''}})}
    }

    Write-Host '  [5/5] Building key-value map...' -ForegroundColor Cyan

    # EV helper using script-scoped dce/dre
    function EV { param([string]$Key)
        if ($script:dce.ContainsKey($Key)) { return $script:dce[$Key] }
        if ($script:dre.ContainsKey($Key)) { return $script:dre[$Key] }
        return ''
    }

    $keyValues = @{}
    foreach ($wk in $WatchedKeys) { $keyValues[$wk.Key] = EV $wk.Key }
    # Keys that live authoritatively in DRE — override DCE-first EV result
    foreach ($dk in @('cas||casdownaction')) {
        if ($script:dre.ContainsKey($dk)) { $keyValues[$dk] = $script:dre[$dk] }
    }

    # SMTP parse
    $smtpXml  = EV 'cas||smtpauthenticationsec'
    $smtpAddr = XV $smtpXml 'address'
    $smtpSrv  = ''; $smtpPort = ''
    if ($smtpAddr -match '^([^:]+):(\d+)$') { $smtpSrv=$Matches[1]; $smtpPort=$Matches[2] }

    # Auth data — parsed from EQVar values for structured HTML rendering
    $authData = @{}

    # Resolve security policy SIDs to account names
    $rawSids = EV 'cas||securitypolicysids'
    $resolvedSids = [System.Collections.Generic.List[string]]::new()
    if ($rawSids) {
        foreach ($sid in ($rawSids -split ';')) {
            $sid = $sid.Trim()
            if (-not $sid) { continue }
            try {
                $name = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount]).Value
                $resolvedSids.Add("$name  ($sid)")
            } catch {
                $resolvedSids.Add($sid)
            }
        }
    }
    $authData['SecurityPolicySids'] = $resolvedSids

    # Parse clientauthconfig XML — domains, LDAP servers, auth method flags
    # type=0 items = Windows domains; type=2 items = LDAP servers
    $cacXml = EV 'cas||clientauthconfig'
    $authData['AuthEquitracPINS']         = ''
    $authData['AuthExternalUserIDPassword'] = ''
    $authData['AuthExternalPassword']     = ''
    $authData['Domains']                  = [System.Collections.Generic.List[string]]::new()
    $authData['LdapServers']              = [System.Collections.Generic.List[object]]::new()
    $authData['UserAccAutoSync']          = EV 'cas||useraccautosync'
    if ($cacXml) {
        try {
            $xdoc = [xml]$cacXml
            $root = $xdoc.DocumentElement
            $authData['AuthEquitracPINS']           = if ($root.EquitracPINS)           { $root.EquitracPINS }           else { '' }
            $authData['AuthExternalUserIDPassword'] = if ($root.ExternalUserIDPassword) { $root.ExternalUserIDPassword } else { '' }
            $authData['AuthExternalPassword']       = if ($root.ExternalPassword)       { $root.ExternalPassword }       else { '' }
            $ldapTypeMap  = @{'0'='First try AD-style, then try simple';'1'='Try AD-style only';'2'='Try simple only'}
            $ldapUidMap   = @{'0'='Do not change';'1'='Remove email domain'}
            $i = 0
            while ($true) {
                $item = $root.SelectSingleNode("item$i")
                if (-not $item) { break }
                if ($item.type -eq '0') {
                    $authData['Domains'].Add($item.default)
                } elseif ($item.type -eq '2') {
                    $sm  = if ($item.submethod)   { $item.submethod }   else { '' }
                    $uid = if ($item.useridmod)   { $item.useridmod }   else { '' }
                    $authData['LdapServers'].Add([PSCustomObject]@{
                        Server    = if ($item.server)      { $item.server }      else { '' }
                        Port      = if ($item.port)        { $item.port }        else { '' }
                        LdapType  = if ($ldapTypeMap.ContainsKey($sm))  { $ldapTypeMap[$sm] }  else { $sm }
                        ForceSSL  = if ($item.forcessl -eq '-1') { 'Enabled' } else { 'Disabled' }
                        LdapV3    = if ($item.versiontouse -eq '3') { 'Enabled' } else { 'Disabled' }
                        DirectBind= if ($item.dnprefix -or $item.dnsuffix) { 'true' } else { 'false' }
                        DnPrefix  = if ($item.dnprefix)    { $item.dnprefix }    else { '' }
                        DnSuffix  = if ($item.dnsuffix)    { $item.dnsuffix }    else { '' }
                        UserIdMod = if ($ldapUidMap.ContainsKey($uid))  { $ldapUidMap[$uid] }  else { $uid }
                    })
                }
                $i++
            }
        } catch { }
    }

    # Parse hidparameters XML — HID card groups
    # Structure: HIDParameters > Groups > item0..N > { GroupName, IsDefault, Items > item0..N }
    $hidXml = EV 'dce||hidparameters'
    $authData['HidGroups'] = [System.Collections.Generic.List[object]]::new()
    if ($hidXml) {
        try {
            $xdoc = [xml]$hidXml
            $groupsNode = $xdoc.DocumentElement.SelectSingleNode('Groups')
            if ($groupsNode) {
                foreach ($grpNode in $groupsNode.ChildNodes) {
                    if ($grpNode.NodeType -ne 1) { continue }
                    $grpName    = if ($grpNode.GroupName) { $grpNode.GroupName } else { $grpNode.Name }
                    $isDefault  = $grpNode.IsDefault -eq '-1'
                    $items      = [System.Collections.Generic.List[object]]::new()
                    $itemsNode  = $grpNode.SelectSingleNode('Items')
                    if ($itemsNode) {
                        foreach ($itemNode in $itemsNode.ChildNodes) {
                            if ($itemNode.NodeType -ne 1) { continue }
                            $facPIN = if ($itemNode.UseAsPIN) { $itemNode.UseAsPIN } else { '' }
                            $items.Add([PSCustomObject]@{
                                FacCode  = "$($itemNode.FacStart),$($itemNode.FacEnd),$($itemNode.FacWidth)"
                                FacMatch = if ($itemNode.FacMatch) { $itemNode.FacMatch } else { '' }
                                FacInPIN = if ($facPIN -eq 'FacID') { 'Yes' } elseif ($facPIN) { $facPIN } else { 'No' }
                                IDCode   = "$($itemNode.IDStart),$($itemNode.IDEnd),$($itemNode.IDWidth)"
                            })
                        }
                    }
                    $authData['HidGroups'].Add([PSCustomObject]@{
                        GroupName = $grpName
                        IsDefault = $isDefault
                        Items     = $items
                    })
                }
            }
        } catch { }
    }

    # --- Network data collection ---
    $netData = @{
        UpgradeMode    = EV 'cas||csupgrademode'
        DomainQualif   = EV 'cas||upnisdomainqualified'
        DefaultDomain  = EV 'cas||domainqualif'
        SmtpList       = [System.Collections.Generic.List[object]]::new()
        SnmpPollSecs   = EV 'dre||dmepolltime'
        SnmpList       = [System.Collections.Generic.List[object]]::new()
    }

    # Parse default SMTP server (cas||smtpauthenticationsec)
    function Parse-SmtpXml { param([string]$xml,[string]$fallbackName)
        if (-not $xml) { return $null }
        try {
            $x = [xml]$xml
            $r = $x.DocumentElement
            $addr = if ($r.address) { $r.address } else { '' }
            $srv=''; $port=''
            if ($addr -match '^(.+):(\d+)$') { $srv=$Matches[1]; $port=$Matches[2] }
            else { $srv=$addr }
            $title = if ($r.title -and $r.title.Trim()) { $r.title.Trim() } else { $fallbackName }
            return [PSCustomObject]@{
                Name     = $title
                Server   = $srv
                Port     = $port
                TLS      = ($r.isenablessl -eq '-1')
                MailFrom = if ($r.mailfrom)  { $r.mailfrom }  else { '' }
                AuthUser = if ($r.user)      { $r.user }      else { '' }
                AuthPass = if ($r.password -and $r.password.Trim()) { 'Configured' } else { 'Not configured' }
                BasicAuth= ($r.user -or ($r.password -and $r.password.Trim()))
            }
        } catch { return $null }
    }
    $defSmtp = Parse-SmtpXml (EV 'cas||smtpauthenticationsec') '<Default>'
    if ($defSmtp) { $netData['SmtpList'].Add($defSmtp) }
    # Additional SMTP servers (cas||smtpservers)
    $smtpSrvXml = EV 'cas||smtpservers'
    if ($smtpSrvXml -and $smtpSrvXml -notmatch '^\s*<servers\s*/>\s*$' -and $smtpSrvXml -notmatch '<servers>\s*</servers>') {
        try {
            $sx = [xml]$smtpSrvXml
            $i = 0
            while ($true) {
                $item = $sx.DocumentElement.SelectSingleNode("item$i")
                if (-not $item) { break }
                $inner = $item.InnerXml
                $parsed = Parse-SmtpXml "<smtpdefault>$inner</smtpdefault>" "Server $($i+1)"
                if ($parsed) { $netData['SmtpList'].Add($parsed) }
                $i++
            }
        } catch { }
    }

    # Parse SNMP configs (cas||snmpconfigsets)
    $snmpXml = EV 'cas||snmpconfigsets'
    if ($snmpXml) {
        try {
            $sx = [xml]$snmpXml
            $root = $sx.DocumentElement
            # Version is XML schema version (always 0), not polling interval — skipped
            $csNode = $root.SelectSingleNode('ConfigSets')
            if ($csNode) {
                foreach ($cfgNode in $csNode.ChildNodes) {
                    $eq = $cfgNode.SelectSingleNode('EQSNMP')
                    if (-not $eq) { continue }
                    $setName = if ($eq.SetName) { $eq.SetName } else { '' }
                    # Decode HTML entities in name (&lt;Default&gt; → <Default>)
                    $setName = $setName -replace '&lt;','<' -replace '&gt;','>' -replace '&amp;','&'
                    $proto = $eq.SelectSingleNode('Protocol/item0')
                    $type = if ($proto -and $proto.Type) { $proto.Type } else { '1' }
                    if ($type -eq '1') {
                        $commRead  = if ($proto.Community -and $proto.Community.Read  -and $proto.Community.Read.Trim())  { 'Configured' } else { 'Not configured' }
                        $commWrite = if ($proto.Community -and $proto.Community.Write -and $proto.Community.Write.Trim()) { 'Configured' } else { 'Not configured' }
                        $netData['SnmpList'].Add([PSCustomObject]@{
                            Name=''; SetName=$setName; Version='v1/v2c'
                            CommunityRead=$commRead; CommunityWrite=$commWrite
                            SecName=''; CtxName=''; AuthProto=''; AuthKey=''; PrivProto=''; PrivKey=''
                        })
                    } else {
                        $secName  = if ($proto.SecurityName)  { $proto.SecurityName  } else { '' }
                        $ctxName  = if ($proto.ContextName)   { $proto.ContextName   } else { '' }
                        $authProtRaw = if ($proto.AuthProtocol) { $proto.AuthProtocol } else { '' }
                        $authProt = switch ($authProtRaw) { '1' { 'MD5' } '2' { 'SHA1' } default { $authProtRaw } }
                        $authKey  = if ($proto.AuthKey  -and $proto.AuthKey.Trim())  { 'Configured' } else { 'Not configured' }
                        $privProtRaw = if ($proto.PrivProtocol) { $proto.PrivProtocol } else { '' }
                        $privProt = switch ($privProtRaw) { '1' { 'AES-128' } '2' { 'DES' } default { $privProtRaw } }
                        $privKey  = if ($proto.PrivKey  -and $proto.PrivKey.Trim())  { 'Configured' } else { 'Not configured' }
                        $netData['SnmpList'].Add([PSCustomObject]@{
                            Name=''; SetName=$setName; Version='v3'
                            CommunityRead=''; CommunityWrite=''
                            SecName=$secName; CtxName=$ctxName
                            AuthProto=$authProt; AuthKey=$authKey
                            PrivProto=$privProt; PrivKey=$privKey
                        })
                    }
                }
            }
        } catch { }
    }

    return [PSCustomObject]@{
        Hostname    = $hostname
        AppName     = $AppName
        Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Components  = $components
        KeyValues   = $keyValues
        PriceLists  = $priceLists
        Workflows   = $workflows
        WFolders    = $wfFolders
        PullGroups  = $pullGroups
        Users       = $users
        SmtpServer  = $smtpSrv
        SmtpPort    = $smtpPort
        Currency    = (EV 'cas||currencyiso4217')
        LicenseHost = (EV 'cas||fneserverhost')
        DceMap      = $script:dce
        DreMap      = $script:dre
        WinData     = $winData
        AuthData    = $authData
        NetData     = $netData
    }
}

# ============================================================
# WRITE FULL TXT
# ============================================================
function Write-FullTxt {
    param($data, [string]$OutDir, [string]$mode)
    $buf = [System.Collections.Generic.List[string]]::new()
    function W  { param([string]$l='') $buf.Add($l) }
    function WH { param([string]$h) W ''; W ('='*70); W "  $h"; W ('='*70) }
    function WS { param([string]$s) W ''; W "  --- $s ---" }
    function WF { param([string]$l,[string]$v) W "  ${l}: $v" }

    W "$($data.AppName) Configuration Export"
    W "Server    : $($data.Hostname)"
    W "Generated : $($data.Timestamp)"
    W "Mode      : $mode"

    # ---- Windows Server sections ----
    if ($data.WinData) {
        $w2 = $data.WinData
        WH "SYSTEM IDENTITY"
        WF 'Computer Name'  $w2.ComputerName
        WF 'Manufacturer'   $w2.Manufacturer
        WF 'Model'          $w2.Model
        WF 'Platform'       $w2.Platform
        if ($w2.Hypervisor) { WF 'Hypervisor' $w2.Hypervisor }
        WF 'UUID'           $w2.UUID
        WF 'BIOS Version'   $w2.BiosVersion
        WF 'BIOS Serial'    $w2.BiosSerial

        WH "OPERATING SYSTEM"
        WF 'OS Name'        $w2.OsName
        WF 'OS Version'     $w2.OsVersion
        WF 'Architecture'   $w2.OsArch
        WF 'Install Date'   $w2.OsInstall

        WH "HARDWARE AND STORAGE"
        WF 'RAM'            $w2.Ram
        WF 'CPU'            $w2.CpuName
        WF 'Cores'          $w2.CpuCores
        WF 'Threads'        $w2.CpuThreads
        WS "Drives"
        foreach ($d in $w2.Drives) { WF "  Drive $($d.Name)" "$($d.Total) GB total, $($d.Used) GB used, $($d.Free) GB free" }

        WH "NETWORK"
        foreach ($n in $w2.Nics) {
            WS $n.Desc
            WF '  IP'        $n.IP
            WF '  Subnet'    $n.Subnet
            WF '  Gateway'   $n.Gateway
            WF '  DNS'       $n.DNS
            WF '  DHCP'      $n.DHCP
            WF '  MAC'       $n.Mac
            WF '  Speed'     $n.Speed
            WF '  NetBIOS'   $n.NetBIOS
        }

        WH "DOMAIN"
        WF 'Domain'         $w2.DomainName
        WF 'Domain Joined'  $(if ($w2.DomainJoined) { "Yes ($($w2.DomainName))" } else { 'No' })

        WH "INSTALLED PRINT MANAGEMENT SOFTWARE"
        if ($w2.PmApps.Count -eq 0) { W '  (none detected)' }
        else {
            foreach ($a in $w2.PmApps) {
                WF "  Product"  $a.Name
                if ($a.Version)   { WF "  Version"   $a.Version }
                if ($a.SvcStatus) { WF "  Service"   $a.SvcStatus }
                if ($a.Path)      { WF "  Path"      $a.Path }
                W ''
            }
        }

        WH "WINDOWS ROLES AND FEATURES"
        WF '.NET 3.5'       $w2.DotNet35
        WF '.NET 4.8'       $w2.DotNet48
        WF 'Print Server Role' $w2.PrintRoleStatus
        WF 'Spooler'        $w2.SpoolerStatus
        WF 'Spool Folder'   $w2.SpoolFolder
        WS "Installed Roles"
        foreach ($r in $w2.InstalledRoles)    { W "  [Role]    $r" }
        WS "Installed Features"
        foreach ($f in $w2.InstalledFeatures) { W "  [Feature] $f" }

        WH "SQL SERVER INSTANCES"
        if ($w2.SqlInstances.Count -eq 0) { W '  None detected' }
        else {
            foreach ($si in $w2.SqlInstances) {
                WS "Instance: $($si.Name)"
                WF '  Edition'          $si.Edition
                WF '  Version'          $si.Version
                WF '  Service'          $si.Status
                WF '  Service Account'  $si.ServiceAcct
                WF '  Auth Mode'        $si.AuthMode
                WF '  Install Path'     $si.InstallPath
                WF '  Data Root'        $si.DataRoot
                WF '  Default Data Dir' $si.DefaultData
                WF '  Default Log Dir'  $si.DefaultLog
                WF '  Backup Dir'       $si.BackupDir
                WF '  Error Log'        $si.ErrorLog
                WF '  Named Pipes'      "$(if($si.NpEnabled){'Enabled'}else{'Disabled'})"
                WF '  TCP/IP'           "$(if($si.TcpEnabled){'Enabled'}else{'Disabled'})  Port=$($si.TcpPort)"
                WF '  TCP Keep Alive'   $si.TcpKeepAlive
                WF '  TCP Listen All'   $si.TcpListenAll
                if ($si.TcpIpAddrs.Count -gt 0) {
                    W '  TCP IP Addresses:'
                    foreach ($ip in $si.TcpIpAddrs) {
                        W "    $($ip.Name) [$($ip.IP)]: Active=$($ip.Active) Enabled=$($ip.Enabled) Port=$($ip.Port)"
                    }
                }
                WF '  FILESTREAM'       $si.FsDesc
                WF '  Max Memory'       $si.MaxMemoryMB
                WF '  Min Memory'       $si.MinMemoryMB
                WF '  User Instances'   $si.UserInstances
                if ($si.SysAdmins.Count -gt 0) {
                    W '  Sysadmin Members:'
                    foreach ($a in $si.SysAdmins) { W "    $a" }
                }
                if ($si.TempDbFiles.Count -gt 0) {
                    W '  TempDB Files:'
                    foreach ($tf in $si.TempDbFiles) {
                        W "    [$($tf.Type)] $($tf.Name): $($tf.Path)  Size=$($tf.SizeMB) MB  Growth=$($tf.Growth)"
                    }
                }
            }
        }

        WH "PRINT QUEUES ($($w2.PrintQueues.Count) total)"
        WF 'Spool Folder'   $w2.SpoolFolder
        foreach ($q in $w2.PrintQueues) {
            WS "Queue: $($q.Name)"
            WF '  Driver'           $q.Driver
            WF '  Port'             $q.Port
            WF '  Status'           $q.Status
            WF '  Paper Size'       $q.PaperSize
            WF '  Duplex'           $q.Duplex
            WF '  Color'            $q.Color
            WF '  Advanced Features'$q.AdvFeatures
            WF '  Render on Client' $q.RenderOnClient
            WF '  Use App Color'    $q.ColorMgmt
            WF '  ICM Method'       $q.IcmMethod
            WF '  Shared'           $(if ($q.Shared) { "Yes (ShareName: $($q.ShareName))" } else { 'No' })
            WF '  Published in AD'  $(if ($q.Published) { 'Yes' } else { 'No' })
            WF '  x86 Driver'       $(if ($q.Has32bit) { 'Installed' } else { 'Not installed' })
            if ($q.FinisherInst) {
                WF '  Finisher'         $q.FinisherInst
                WF '  Staple Installed' $q.StapleInst
                WF '  Punch Installed'  $q.PunchInst
                WF '  Booklet Maker'    $q.BookletInst
                WF '  Offset Stacking'  $q.OffsetInst
            }
        }

        WH "PRINTER DRIVERS"
        W "  x64 drivers ($($w2.DriversX64.Count)):"
        $w2.DriversX64 | Sort-Object Name | ForEach-Object { W "    $($_.Name)" }
        W "  x86 drivers ($($w2.DriversX86.Count)):"
        $w2.DriversX86 | Sort-Object Name | ForEach-Object { W "    $($_.Name)" }

        WH "PRINTER PORTS ($($w2.PrinterPorts.Count) total)"
        foreach ($pt in $w2.PrinterPorts) {
            $line = "  $($pt.Name)"
            if ($pt.IP)   { $line += "  IP=$($pt.IP)" }
            if ($pt.Port) { $line += "  Port=$($pt.Port)" }
            W $line
        }
    }

    WH "INSTALLED COMPONENTS"
    if ($data.Components.Count -eq 0) { W "  (none)" }
    else {
        foreach ($c in $data.Components) {
            $s = "  $($c.Name)"
            if ($c.Desc)    { $s += "  ($($c.Desc))" }
            if ($c.Version) { $s += "  v$($c.Version)" }
            if ($c.Date)    { $s += "  [$($c.Date)]" }
            W $s
        }
    }

    foreach ($sec in $SectionMeta.Keys) {
        WH $SectionMeta[$sec].ToUpper()
        foreach ($wk in ($WatchedKeys | Where-Object { $_.Section -eq $sec })) {
            $raw = FV $data.KeyValues[$wk.Key] 300
            $disp = if ($wk.ValueMap -and $wk.ValueMap.ContainsKey($raw)) { "$raw ($($wk.ValueMap[$raw]))" } else { $raw }
            WF $wk.Label $disp
        }
    }

    WH "PRICING (cat_pricelist)"
    if ($data.PriceLists.Count -eq 0) { W "  (none)" }
    else {
        foreach ($pl in $data.PriceLists) {
            WS "Price List: $($pl.Name)  [ID=$($pl.Id)  Type=$($pl.Type)]"
            W  "  Rates: $(if($pl.Rates){$pl.Rates -join ' | '}else{'(none)'})"
        }
    }

    WH "WORKFLOWS (cas_scan_alias)"
    foreach ($wf in $data.Workflows) {
        $t = if ($wf.Type) { "  [$($wf.Type)]" } else { '' }
        W "  [$($wf.Id)] $($wf.Name)$t"
    }
    WS "Workflow Folders"
    foreach ($wff in $data.WFolders) { W "  [$($wff.Id)] $($wff.Name)" }

    WH "PULL PRINT GROUPS (cas_pullgroups)"
    if ($data.PullGroups.Count -eq 0) { W "  (none)" } else {
        foreach ($pg in $data.PullGroups) { W "  [$($pg.Id)] $($pg.Name)" }
    }

    WH "USER ACCOUNTS (cas_user_ext)"
    $shown = 0
    foreach ($u in $data.Users) {
        W "  [$($u.Id)] $($u.Domain)\$($u.Name)"
        if ((++$shown) -ge 50) { W "  ... ($($data.Users.Count - 50) more - see raw BCP)"; break }
    }

    WH "FULL EQVAR DUMP (DCE_config.db3)"
    if ($data.DceMap.Count -eq 0) { W "  (not available)" } else {
        $bySub = @{}
        $data.DceMap.GetEnumerator() | ForEach-Object {
            $sub = ($_.Key -split '\|', 3)[0]; if (-not $sub) { $sub = '(blank)' }
            if (-not $bySub[$sub]) { $bySub[$sub] = [System.Collections.Generic.List[string]]::new() }
            $val = $_.Value
            if ([string]::IsNullOrWhiteSpace($val)) { $val = '(empty)' }
            elseif ($val.Length -gt 160) { $val = $val.Substring(0, 160) + '...' }
            $bySub[$sub].Add("  $($_.Key) = $val")
        }
        foreach ($sub in ($bySub.Keys | Sort-Object)) {
            WS "SubSystem: $sub"
            $bySub[$sub] | Sort-Object | ForEach-Object { W $_ }
        }
    }

    $outFile = Join-Path $OutDir "${mode}_FULL.txt"
    [System.IO.File]::WriteAllLines($outFile, $buf, [System.Text.Encoding]::ASCII)
    Write-Host "  -> $outFile ($($buf.Count) lines)" -ForegroundColor Green
    return $outFile
}

# ============================================================
# WRITE SUMMARY TXT
# ============================================================
function Write-SummaryTxt {
    param($data, [string]$OutDir, [string]$mode)
    $buf = [System.Collections.Generic.List[string]]::new()
    function W { param([string]$l='') $buf.Add($l) }

    W "$($data.AppName) - Configuration Summary"
    W "Server    : $($data.Hostname)"
    W "Generated : $($data.Timestamp)"
    W "Mode      : $mode"
    W ''
    W '--- KEY COUNTS ---'
    W "Components  : $($data.Components.Count)"
    W "Price Lists : $($data.PriceLists.Count)"
    W "Workflows   : $($data.Workflows.Count)"
    W "Pull Groups : $($data.PullGroups.Count)"
    W "Users       : $($data.Users.Count)"
    W ''
    W '--- QUICK SETTINGS ---'
    W "SMTP        : $(if($data.SmtpServer){"$($data.SmtpServer):$($data.SmtpPort)"}else{'(not configured)'})"
    W "Currency    : $($data.Currency)"
    W "License     : $($data.LicenseHost)"
    W "Card Swipe  : $(FV ($data.KeyValues['dce||enableswipe']))"
    W "PIN Alt     : $(FV ($data.KeyValues['dce||registerpinasalternate']))"
    W "Precision   : $(FV ($data.KeyValues['cas||precision']))"
    W ''
    W '--- INSTALLED COMPONENTS ---'
    foreach ($c in $data.Components) { W "  $($c.Name)  v$($c.Version)" }
    W ''
    W '--- PRICE LISTS ---'
    foreach ($pl in $data.PriceLists) {
        W "  [$($pl.Id)] $($pl.Name)  ($($pl.Type))  ->  $(if($pl.Rates){$pl.Rates -join ' | '}else{'no rates'})"
    }
    W ''
    W '--- WORKFLOWS ---'
    foreach ($wf in $data.Workflows) { W "  [$($wf.Id)] $($wf.Name)  [$($wf.Type)]" }
    W ''
    W '--- PULL PRINT GROUPS ---'
    foreach ($pg in $data.PullGroups) { W "  [$($pg.Id)] $($pg.Name)" }

    $outFile = Join-Path $OutDir "${mode}_SUMMARY.txt"
    [System.IO.File]::WriteAllLines($outFile, $buf, [System.Text.Encoding]::ASCII)
    Write-Host "  -> $outFile" -ForegroundColor Green
}

# ============================================================
# WRITE METADATA JSON
# ============================================================
function Write-MetadataJson {
    param($data, [string]$OutDir, [string]$mode)
    $smtp = if ($data.SmtpServer) { "$($data.SmtpServer):$($data.SmtpPort)" } else { '' }
    # Build component array
    $compArr = ($data.Components | ForEach-Object {
        "    {`"name`":`"$($_.Name)`",`"version`":`"$($_.Version)`"}"
    }) -join ",`n"
    $plArr = ($data.PriceLists | ForEach-Object {
        $rateStr = ($_.Rates | ForEach-Object { "`"$_`"" }) -join ','
        "    {`"id`":`"$($_.Id)`",`"name`":`"$($_.Name)`",`"type`":`"$($_.Type)`",`"rates`":[$rateStr]}"
    }) -join ",`n"
    $wfArr = ($data.Workflows | ForEach-Object {
        "    {`"id`":`"$($_.Id)`",`"name`":`"$($_.Name)`",`"type`":`"$($_.Type)`"}"
    }) -join ",`n"

    $json = "{`n" +
            "  `"app`": `"$($data.AppName)`",`n" +
            "  `"server`": `"$($data.Hostname)`",`n" +
            "  `"timestamp`": `"$($data.Timestamp)`",`n" +
            "  `"mode`": `"$mode`",`n" +
            "  `"currency`": `"$($data.Currency)`",`n" +
            "  `"smtpServer`": `"$smtp`",`n" +
            "  `"licenseHost`": `"$($data.LicenseHost)`",`n" +
            "  `"components`": [`n$compArr`n  ],`n" +
            "  `"priceLists`": [`n$plArr`n  ],`n" +
            "  `"workflows`": [`n$wfArr`n  ],`n" +
            "  `"counts`": {`"components`":$($data.Components.Count),`"priceLists`":$($data.PriceLists.Count),`"workflows`":$($data.Workflows.Count),`"pullGroups`":$($data.PullGroups.Count),`"users`":$($data.Users.Count)}`n" +
            "}`n"

    $outFile = Join-Path $OutDir 'metadata.json'
    [System.IO.File]::WriteAllText($outFile, $json, [System.Text.Encoding]::ASCII)
    Write-Host "  -> $outFile" -ForegroundColor Green
}

# ============================================================
# BUILD AUTH HTML (structured Authentication section)
# ============================================================
function AG     { param([string]$t) "<div class='auth-grp'>$t</div>" }
function NG     { param([string]$t) "<div class='auth-grp'>$t</div>" }
function NRow   { param([string]$k,[string]$v)
    $ec = if (-not $v -or $v -eq '(not set)') { ' e' } else { '' }
    "<tr><td class='k'>$(HE $k)</td><td class='v$ec'>$(HE $v)</td></tr>"
}
function NTable { param([string]$body) "<table class='kv'>$body</table>" }
function NRS1   { param([string]$k,[string]$v) "<tr><td class='k sub1'>$(HE $k)</td><td class='v'>$v</td></tr>" }
function NOn    { param([string]$v,[string]$t='-1') if ($v -eq $t -or $v -eq '1') { "<span class='dc-on'>Enabled</span>" } else { "<span class='dc-off'>Disabled</span>" } }
function NCB    { param([bool]$on) if ($on) { "<span class='eq-on'>&#9745; Yes</span>" } else { "<span class='eq-off'>&#9744; No</span>" } }
function NMask  { param([string]$v) if ($v -eq 'Configured') { "<span style='color:#555'>&#11044;&#11044;&#11044;&#11044;&#11044;&#11044; <small style='color:#888'>(configured)</small></span>" } else { "<span class='dc-off'>Not configured</span>" } }
function ARow   { param([string]$k,[string]$v)
    $ec = if ($v -eq '(not set)') { ' e' } else { '' }
    "<tr><td class='k'>$(HE $k)</td><td class='v$ec'>$(HE $v)</td></tr>"
}
function ARowH  { param([string]$k,[string]$v) "<tr><td class='k'>$(HE $k)</td><td class='v'>$v</td></tr>" }
function ARS1   { param([string]$k,[string]$v) "<tr><td class='k sub1'>$(HE $k)</td><td class='v'>$v</td></tr>" }
function ARS2   { param([string]$k,[string]$v) "<tr><td class='k sub2'>$(HE $k)</td><td class='v'>$v</td></tr>" }
function SecLbl { param([string]$t) "<tr><td colspan='2' class='sec-lbl'>$(HE $t)</td></tr>" }
function ATable { param([string]$body) "<table class='kv'>$body</table>" }
function DCOn   { param([string]$v) if ($v -eq '1') { "<span class='dc-on'>On</span>" } else { "<span class='dc-off'>Off</span>" } }
function DCEnI  { param([string]$v) if ($v -eq '0') { "<span class='dc-on'>Enabled</span>" } else { "<span class='dc-off'>Disabled</span>" } }
function YNo    { param([string]$v) if ($v -eq '1') { "<span class='eq-on'>&#9745; Yes</span>" } else { "<span class='eq-off'>&#9744; No</span>" } }
function CBx    { param([string]$v,[string]$t='Y') $on=$v-eq$t-or$v-eq'1'; if($on){"<span class='eq-on'>&#9745; Checked</span>"}else{"<span class='eq-off'>&#9744; Unchecked</span>"} }
function AVMap  { param([string]$raw,[hashtable]$map)
    $r = FV $raw 300
    if ($map -and $map.ContainsKey($raw)) { return "$raw ($($map[$raw]))" }
    return $r
}
function AVMapH { param([string]$raw,[hashtable]$map)
    if ($map -and $map.ContainsKey($raw)) { return $map[$raw] }
    if ($raw) { return $raw }
    return '(not set)'
}

function Build-AuthHtml {
    param($data)
    $ad = $data.AuthData
    $kv = $data.KeyValues
    $html = ''

    # --- Access Permissions ---
    # Positions in cas||securitypolicysids: 0=Admin, 1=Accounts, 2=Print Distribution, 3=Reports
    $html += AG 'Access Permissions'
    $sidRows = ''
    $sids = if ($ad -and $ad['SecurityPolicySids']) { $ad['SecurityPolicySids'] } else { @() }
    $permLabels = @('Admin', 'Accounts', 'Reports', 'Print Distribution')
    $permIdx    = @(0, 1, 3, 4)
    for ($pi = 0; $pi -lt 4; $pi++) {
        $idx = $permIdx[$pi]
        $val = if ($sids.Count -gt $idx) { $sids[$idx] } else { '(not set)' }
        $sidRows += ARow $permLabels[$pi] $val
    }
    $html += ATable $sidRows

    # --- External Authentication ---
    $html += AG 'External Authentication'

    # Windows Authentication
    $html += "<div class='sub'>Windows Authentication</div>"
    $winRows = ''
    $winAuthState = if ($ad -and $ad['Domains'] -and $ad['Domains'].Count -gt 0) { 'Enabled' } else { 'Disabled' }
    $winRows += ARow 'Enable Windows Authentication' $winAuthState
    if ($ad -and $ad['Domains'] -and $ad['Domains'].Count -gt 0) {
        $j = 1
        foreach ($d in $ad['Domains']) { $winRows += ARow "Domain $j" $d; $j++ }
    } else { $winRows += ARow 'Domains' '(not set)' }
    $html += ATable $winRows

    # LDAP Authentication
    $html += "<div class='sub'>LDAP Authentication</div>"
    $ldapEnabled = if ($ad -and $ad['LdapServers'] -and $ad['LdapServers'].Count -gt 0) { 'Enabled' } else { 'Disabled' }
    $syncVal = if ($ad -and $ad['UserAccAutoSync']) { $ad['UserAccAutoSync'] } else { $kv['cas||useraccautosync'] }
    $ldapTopRows = ''
    $ldapTopRows += ARow 'Enable LDAP Authentication' $ldapEnabled
    $ldapTopRows += ARow 'Synchronize user attributes on login' (AVMap $syncVal @{'0'='Disabled';'1'='Enabled'})
    $html += ATable $ldapTopRows

    if ($ad -and $ad['LdapServers'] -and $ad['LdapServers'].Count -gt 0) {
        $html += "<div class='sub' style='margin-top:8px'>LDAP Server Configuration</div>"
        foreach ($srv in $ad['LdapServers']) {
            $html += "<table class='kv' style='margin-bottom:8px'>"
            $html += "<tr><td class='k'>Server Name</td><td class='v'>$(HE $srv.Server)</td></tr>"
            $html += "<tr><td class='k'>Port</td><td class='v'>$(HE $srv.Port)</td></tr>"
            $html += "<tr><td class='k'>Type</td><td class='v'>$(HE $srv.LdapType)</td></tr>"
            $html += "<tr><td class='k'>Force SSL</td><td class='v'>$(HE $srv.ForceSSL)</td></tr>"
            $html += "<tr><td class='k'>Use LDAP version 3</td><td class='v'>$(HE $srv.LdapV3)</td></tr>"
            $html += "<tr><td class='k'>Authentication Method</td><td class='v'>Direct bind</td></tr>"
            $html += "<tr><td class='k'>Direct bind</td><td class='v'>$(HE $srv.DirectBind)</td></tr>"
            $html += "<tr><td class='k'>DN Prefix</td><td class='v'>$(HE $srv.DnPrefix)</td></tr>"
            $html += "<tr><td class='k'>DN Suffix</td><td class='v'>$(HE $srv.DnSuffix)</td></tr>"
            $html += "<tr><td class='k'>User ID Modification</td><td class='v'>$(HE $srv.UserIdMod)</td></tr>"
            $html += "</table>"
        }
    }

    # --- Device Clients ---
    $html += AG 'Device Clients'
    $swipeOn  = $kv['dce||enableswipe']
    $reqExtra = $kv['dce||nosecondaryidwithswipe']   # 0=Enabled (inverted key)
    $reqPin2  = $kv['dce||requirepasswordifavailablekeyboard']
    $kbOn     = $kv['dce||registerpin']
    $dcHtml   = "<table class='kv'>"
    # A. Card swipe block
    $dcHtml  += ARowH 'Allow card swipe login'  (DCOn $swipeOn)
    $dcHtml  += ARS1  'Require additional authentication with Equitrac password'  (DCEnI $reqExtra)
    if ($reqExtra -eq '0') {
        $dcHtml += ARS2 'Require password only if PIN2 available (legacy devices only)' (YNo $reqPin2)
    }
    # Separator between card swipe and keyboard blocks
    $dcHtml  += "<tr><td colspan='2' style='height:6px;border:none;background:none'></td></tr>"
    # B. Keyboard block
    $dcHtml  += ARowH 'Allow keyboard login'  (DCOn $kbOn)
    # Keyboard Equitrac auth mode — locked to "with password prompt" when card swipe secondary auth is enabled
    $kbMode   = if ($reqExtra -eq '0') { 'Enabled with password prompt' } else { 'Enabled without password prompt' }
    $dcHtml  += ARS1 'Equitrac authentication'  $kbMode
    $dcHtml  += ARS2 'Require password only if PIN2 available (legacy devices only)'  (YNo $reqPin2)
    $dcHtml  += ARS2 'Deny login with empty password'  (YNo $kv['dce||denypromptemptypassword'])
    # Default function (bottom of Device Clients in UI)
    $dcHtml  += "<tr><td colspan='2' style='height:6px;border:none;background:none'></td></tr>"
    $dcHtml  += ARow 'Default function at device'  (FV $kv['dce||defaultfunction'] 100)
    $dcHtml  += '</table>'
    $html    += $dcHtml

    # --- Workstation Clients and Web Client ---
    $html += AG 'Workstation Clients and Web Client'
    $wcAuthMap = @{'0'='Disabled';'1'='Enabled with password prompt';'2'='Enabled without password prompt'}
    $wcIdpMap  = @{'0'='Disabled';'1'='Enabled'}
    $wcRows  = ARow 'Equitrac authentication'  (AVMapH $kv['cas||authequitracwcws']  $wcAuthMap)
    $wcRows += ARow 'Identity provider'         (AVMapH $kv['cas||identityproviderwcws'] $wcIdpMap)
    $html   += ATable $wcRows

    # --- Equitrac Authentication ---
    $html += AG 'Equitrac Authentication'
    $vPins  = if ($ad -and $ad['AuthEquitracPINS'])           { $ad['AuthEquitracPINS'] }           else { '' }
    $vExtUP = if ($ad -and $ad['AuthExternalUserIDPassword']) { $ad['AuthExternalUserIDPassword'] } else { '' }
    $vExtP  = if ($ad -and $ad['AuthExternalPassword'])       { $ad['AuthExternalPassword'] }       else { '' }
    $eqRows  = ARowH 'Equitrac primary or alternate PIN with secondary PIN'      (CBx $vPins)
    $eqRows += ARowH 'External username and password'                             (CBx $vExtUP)
    $eqRows += ARowH 'Equitrac primary or alternate PIN with external password'   (CBx $vExtP)
    $html   += ATable $eqRows

    # --- Card Registration ---
    $html += AG 'Card Registration'
    $crHtml  = "<table class='kv'>"
    $crHtml += SecLbl 'Registration login'
    $crHtml += ARS1 'Equitrac authentication'  (DCOn $kv['dce||authequitraccardreg'])
    $crHtml += ARS1 'Identity provider'        (DCOn $kv['dce||authidentityprovidercardreg'])
    $crHtml += "<tr><td colspan='2' style='height:6px;border:none;background:none'></td></tr>"
    $crHtml += ARow 'Card number storage'  (AVMap $kv['dce||registerpinasalternate'] @{'0'='Store as primary PIN';'1'='Store as alternate PIN'})
    $crHtml += ARow 'Dual card reader mode'  (AVMap $kv['dce||registertwocards'] @{'0'='Disabled';'1'='Enabled'})
    $crHtml += '</table>'
    $html   += $crHtml

    # --- Encryption ---
    $html += AG 'Encryption'
    $encRows = ARow 'Store Secondary PIN Encrypted' (AVMap $kv['cas||encryptsecondarypin'] @{'N'='Disabled';'Y'='Enabled'})
    $html += ATable $encRows

    # --- Card Setup (PIN Positions + HID Groups) ---
    $html += AG 'Card Setup'
    $csRows = ''
    $csRows += ARow 'Card Data: From Position' (FV $kv['cas||casstartpinposition'] 50)
    $csRows += ARow 'Card Data: To Position'   (FV $kv['cas||casendpinposition']   50)
    $csRows += ARow 'Admin PIN'                (FV $kv['dce||adminpin']            50)
    $csRows += ARow 'Max PIN Length'           (FV $kv['dce||maxpinlength']        50)
    $csRows += ARow 'PIN1 Min Length'          (FV $kv['cas||pin1minlen']          50)
    $csRows += ARow 'PIN1 Max Length'          (FV $kv['cas||pin1maxlen']          50)
    $html += ATable $csRows

    # HID Groups sub-table
    if ($ad -and $ad['HidGroups'] -and $ad['HidGroups'].Count -gt 0) {
        foreach ($grp in $ad['HidGroups']) {
            $defLabel = if ($grp.IsDefault) { ' &nbsp;<small style="font-weight:normal;color:#888">(Default)</small>' } else { '' }
            $html += "<div class='sub'>HID Group: $(HE $grp.GroupName)$defLabel</div>"
            if ($grp.Items.Count -gt 0) {
                $html += "<table class='kv hid-tbl'><thead><tr>"
                $html += "<th>Facility Code<br><small>start,end,width</small></th>"
                $html += "<th>Matching Facility Code</th>"
                $html += "<th>Include Fac. in PIN</th>"
                $html += "<th>ID Code<br><small>start,end,width</small></th>"
                $html += "</tr></thead><tbody>"
                foreach ($item in $grp.Items) {
                    $pinCls = if ($item.FacInPIN -eq 'Yes') { ' style="color:#27ae60;font-weight:500"' } else { '' }
                    $html += "<tr>"
                    $html += "<td>$(HE $item.FacCode)</td>"
                    $html += "<td>$(HE $item.FacMatch)</td>"
                    $html += "<td$pinCls>$(HE $item.FacInPIN)</td>"
                    $html += "<td>$(HE $item.IDCode)</td>"
                    $html += "</tr>"
                }
                $html += "</tbody></table>"
            } else {
                $html += "<p style='color:#aaa;font-size:11px;margin:4px 0 8px'>No decoding items configured.</p>"
            }
        }
    } else {
        $html += "<div class='sub'>HID Decoding Groups</div><p style='color:#aaa;font-size:11px;margin:4px 0 8px'>(not set)</p>"
    }

    # --- CAS Offline Behavior ---
    $html += AG 'CAS Offline Behavior'
    $offRows = ''
    $offRows += ARow 'Login Caching'             (AVMap $kv['dce||supportofflinemode']         @{'0'='Disabled';'1'='Enabled'})
    $offRows += ARow 'Login Cache TTL (days)'    (FV $kv['dce||cacheuserloginofflinettldays']  50)
    $offRows += ARow 'Login Expiry (seconds)'    (FV $kv['cas||loginexpiry']                   50)
    $vPrint = if ($kv['cas||casdownaction']) { $kv['cas||casdownaction'] } else { '' }
    $printMap = @{'PrintChargeLater'='Print, charge accounts later';'AutoSelect'='Auto select';'Auto'='Auto select';'DoNotPrint'='Do not print'}
    $offRows += ARow 'Print Behavior'            (AVMap $vPrint $printMap)
    $html += ATable $offRows

    return $html
}

# ============================================================
# GLOBAL NETWORK SETTINGS HTML
# ============================================================
function Build-NetworkHtml {
    param([object]$data)
    $nd  = $data.NetData
    $html = ''

    # 1. Backward Compatibility
    $html += NG 'Backward Compatibility'
    $vUpgrade = if ($nd['UpgradeMode'] -eq '1') { 'Enabled' } else { 'Disabled' }
    $html += NTable (NRow 'Enable "upgrade mode"' $vUpgrade)

    # 2. Domain Qualification
    $html += NG 'Domain Qualification'
    $dqOn = ($nd['DomainQualif'] -eq '-1' -or $nd['DomainQualif'] -eq '1')
    $vDqOn = if ($dqOn) { "<span class='dc-on'>Enabled</span>" } else { "<span class='dc-off'>Disabled</span>" }
    $dqHtml = "<table class='kv'>"
    $dqHtml += "<tr><td class='k'>Qualify all user IDs with NT domain information</td><td class='v'>$vDqOn</td></tr>"
    if ($dqOn) {
        $dd = if ($nd['DefaultDomain']) { $nd['DefaultDomain'] } else { '(not set)' }
        $dqHtml += NRS1 'Default domain for unqualified user IDs' (HE $dd)
    }
    $dqHtml += '</table>'
    $html += $dqHtml

    # 3. SMTP Mail Server
    $html += NG 'SMTP Mail Server'
    if ($nd['SmtpList'].Count -eq 0) {
        $html += "<p style='color:#aaa;padding:8px;font-size:12px'>No SMTP servers configured.</p>"
    } else {
        # Summary table
        $smtpTbl  = "<table class='net-tbl'><thead><tr><th>Connection name</th><th>SMTP Mail Server</th></tr></thead><tbody>"
        foreach ($s in $nd['SmtpList']) { $smtpTbl += "<tr><td>$(HE $s.Name)</td><td>$(HE $s.Server)</td></tr>" }
        $smtpTbl += '</tbody></table>'
        $html += $smtpTbl
        # Detail cards
        foreach ($s in $nd['SmtpList']) {
            $vBasicAuth = if ($s.BasicAuth) { "<span class='dc-on'>Enabled</span>" }  else { "<span class='dc-off'>Disabled</span>" }
            $vAuthUser  = if ($s.AuthUser) { HE $s.AuthUser } else { '(not set)' }
            $html += "<details class='net-detail'><summary>$(HE $s.Name)</summary>"
            $html += "<table class='kv'>"
            $html += NRow 'Connection name'                                    $s.Name
            $html += NRow 'SMTP Email server (DNS name or IP address)'         $s.Server
            $html += NRow 'SMTP Email server port'                             $s.Port
            $html += "<tr><td class='k'>Enable TLS</td><td class='v'>$(NCB $s.TLS)</td></tr>"
            $html += NRow 'Mail From address for system generated messages'    $s.MailFrom
            $html += "<tr><td class='k'>Basic Authentication</td><td class='v'>$vBasicAuth</td></tr>"
            if ($s.BasicAuth) {
                $html += NRS1 'User Name' $vAuthUser
                $html += "<tr><td class='k sub1'>Password</td><td class='v'>$(NMask $s.AuthPass)</td></tr>"
            }
            $html += '</table></details>'
        }
    }

    # 4. SNMP Configuration
    $html += NG 'SNMP Configuration'
    $pollSecs = $nd['SnmpPollSecs']
    if ($pollSecs -match '^\d+$') {
        $pollMin = [int]([int]$pollSecs / 60)
        $html += NTable (NRow 'Polling interval' "$pollMin minute$(if ($pollMin -ne 1) {'s'})")
    }
    if ($nd['SnmpList'].Count -eq 0) {
        $html += "<p style='color:#aaa;padding:8px;font-size:12px'>No SNMP configurations found.</p>"
    } else {
        $snmpTbl  = "<table class='net-tbl'><thead><tr><th>SNMP configuration name</th><th>Version</th></tr></thead><tbody>"
        foreach ($s in $nd['SnmpList']) { $snmpTbl += "<tr><td>$(HE $s.SetName)</td><td>$(HE $s.Version)</td></tr>" }
        $snmpTbl += '</tbody></table>'
        $html += $snmpTbl
        foreach ($s in $nd['SnmpList']) {
            $html += "<details class='net-detail'><summary>$(HE $s.SetName)</summary>"
            $html += "<table class='kv'>"
            if ($s.Version -eq 'v1/v2c') {
                $html += "<tr><td class='k'>Enable SNMP v1/2c</td><td class='v'><span class='dc-on'>Enabled</span></td></tr>"
                $html += "<tr><td colspan='2' class='sec-lbl'>Community</td></tr>"
                $html += NRS1 'Get (read)'  (NMask $s.CommunityRead)
                $html += NRS1 'Set (write)' (NMask $s.CommunityWrite)
            } else {
                $vSecName  = if ($s.SecName)  { HE $s.SecName  } else { '(not set)' }
                $vCtxName  = if ($s.CtxName)  { HE $s.CtxName  } else { '(not set)' }
                $vAuthProt = if ($s.AuthProto) { HE $s.AuthProto } else { '(not set)' }
                $vPrivProt = if ($s.PrivProto) { HE $s.PrivProto } else { '(not set)' }
                $html += "<tr><td class='k'>Enable SNMP v3</td><td class='v'><span class='dc-on'>Enabled</span></td></tr>"
                $html += "<tr><td colspan='2' class='sec-lbl'>Security</td></tr>"
                $html += NRS1 'Name' $vSecName
                $html += "<tr><td colspan='2' class='sec-lbl'>Context</td></tr>"
                $html += NRS1 'Name' $vCtxName
                $html += "<tr><td colspan='2' class='sec-lbl'>Authentication</td></tr>"
                $html += "<tr><td class='k sub1'>Key</td><td class='v'>$(NMask $s.AuthKey)</td></tr>"
                $html += NRS1 'Protocol' $vAuthProt
                $html += "<tr><td colspan='2' class='sec-lbl'>Privacy</td></tr>"
                $html += "<tr><td class='k sub1'>Key</td><td class='v'>$(NMask $s.PrivKey)</td></tr>"
                $html += NRS1 'Protocol' $vPrivProt
            }
            $html += '</table></details>'
        }
    }

    return $html
}

function Build-DirSyncHtml {
    param($data)
    $html = ''
    $xml = $data.KeyValues['ads||settingsdoc']
    if (-not $xml) {
        return "<p style='color:#aaa;padding:8px;font-size:12px'>No directory sync configuration found.</p>"
    }
    try { $doc = [xml]$xml } catch { return "<p style='color:#aaa;padding:8px;font-size:12px'>Could not parse AD sync settings.</p>" }

    $html += AG 'Active Directory'

    $servers = $doc.ADSyncSettings.ChildNodes
    if (-not $servers -or $servers.Count -eq 0) {
        $html += "<p style='color:#aaa;padding:8px;font-size:12px'>No AD servers configured.</p>"
        return $html
    }

    foreach ($srv in $servers) {
        $dc        = if ($srv.DomainController) { $srv.DomainController } else { '' }
        $partition = if ($srv.ApplicationPartition -and $srv.ApplicationPartition.Trim()) { $srv.ApplicationPartition } `
                     elseif ($srv.Domain -and $srv.Domain.Trim()) { $srv.Domain } else { '(not set)' }
        $authUser  = if ($srv.AuthUserName -and $srv.AuthUserName.Trim()) { '●●●●●● (configured)' } else { 'Not configured' }
        $autoSync  = NOn $srv.DoAutoSync
        $freqMin   = if ($srv.UpdateFrequency -match '^\d+$') { "$($srv.UpdateFrequency) minutes" } else { '(not set)' }

        $html += "<details class='net-detail' open><summary>$(HE $dc)</summary>"
        $html += "<table class='kv'>"
        $html += ARowH 'Domain Controller' (HE $dc)
        $html += ARowH 'Naming Context / Partition' (HE $partition)
        $html += ARowH 'Auth username' $authUser
        $html += ARowH 'Auto sync' $autoSync
        $html += ARowH 'Update frequency' (HE $freqMin)

        # Containers
        $containers = $srv.Containers
        if ($containers -and $containers.ChildNodes.Count -gt 0) {
            $html += "<tr><td colspan='2' class='sec-lbl'>Containers</td></tr>"
            foreach ($c in $containers.ChildNodes) {
                $dn = if ($c.ContainerDN) { $c.ContainerDN } else { '' }
                $html += ARS1 'DN' (HE $dn)
            }
        }
        $html += '</table></details>'
    }

    return $html
}

# ============================================================
# WRITE HTML REPORT
# ============================================================
function Write-HtmlReport {
    param($data, [string]$OutDir, [string]$mode, [string]$fullTxtPath)

    $rawText = if ($fullTxtPath -and (Test-Path $fullTxtPath)) {
        [System.IO.File]::ReadAllText($fullTxtPath)
    } else { '' }

    $badgeColor = switch ($mode) {
        'Before'  { '#e67e22' }
        'After'   { '#27ae60' }
        'Compare' { '#8e44ad' }
        'Full'    { '#2980b9' }
        default   { '#555'   }
    }

    # --- Sidebar nav ---
    $nav  = "<div class='grp'>Windows Server</div>"
    $nav += "<a href='#s-win-system'>System Identity</a>"
    $nav += "<a href='#s-win-os'>Operating System</a>"
    $nav += "<a href='#s-win-hw'>Hardware &amp; Storage</a>"
    $nav += "<a href='#s-win-net'>Network</a>"
    $nav += "<a href='#s-win-domain'>Domain</a>"
    $nav += "<a href='#s-win-sw'>Installed Software</a>"
    $nav += "<a href='#s-win-roles'>Windows Roles</a>"
    $nav += "<a href='#s-win-sql'>SQL Server</a>"
    $nav += "<a href='#s-win-queues'>Print Queues</a>"
    $nav += "<a href='#s-win-drivers'>Drivers &amp; Ports</a>"
    $nav += "<div class='grp'>ControlSuite</div>"
    $nav += "<a href='#s-components'>Installed Components</a>"
    foreach ($sec in $SectionMeta.Keys) {
        $nav += "<a href='#s-$sec'>$($SectionMeta[$sec])</a>"
    }
    $nav += "<div class='grp'>Configuration</div>"
    $nav += "<a href='#s-pricing'>Pricing</a><a href='#s-workflows'>Workflows</a>"
    $nav += "<a href='#s-pullgroups'>Pull Print Groups</a><a href='#s-users'>Users</a>"
    $nav += "<div class='grp'>Full Data</div><a href='#s-eqvar'>EQVar Full Dump</a>"

    # --- Summary cards ---
    $smtp = if ($data.SmtpServer) { "$($data.SmtpServer):$($data.SmtpPort)" } else { 'Not set' }
    $winCards  = "<div class='scard sc-win'><div class='lbl'>Platform</div><div class='val'>$(HE $data.WinData.Platform)</div></div>"
    $winCards += "<div class='scard sc-win'><div class='lbl'>OS</div><div class='val'>$(HE ($data.WinData.OsName -replace 'Windows Server ','WS '))</div></div>"
    $winCards += "<div class='scard sc-win'><div class='lbl'>RAM</div><div class='val'>$(HE $data.WinData.Ram)</div></div>"
    $winCards += "<div class='scard sc-win'><div class='lbl'>Print Queues</div><div class='val'>$($data.WinData.PrintQueues.Count)</div></div>"
    # Build eq cards directly
    $eqCards  = "<div class='scard sc-eq'><div class='lbl'>Components</div><div class='val'>$($data.Components.Count)</div></div>"
    $eqCards += "<div class='scard sc-eq'><div class='lbl'>Price Lists</div><div class='val'>$($data.PriceLists.Count)</div></div>"
    $eqCards += "<div class='scard sc-eq'><div class='lbl'>Workflows</div><div class='val'>$($data.Workflows.Count)</div></div>"
    $eqCards += "<div class='scard sc-eq'><div class='lbl'>Pull Groups</div><div class='val'>$($data.PullGroups.Count)</div></div>"
    $eqCards += "<div class='scard sc-eq'><div class='lbl'>Users</div><div class='val'>$($data.Users.Count)</div></div>"
    $eqCards += "<div class='scard sc-eq'><div class='lbl'>SMTP</div><div class='val'>$(HE $smtp)</div></div>"
    $eqCards += "<div class='scard sc-eq'><div class='lbl'>Currency</div><div class='val'>$(HE $data.Currency)</div></div>"
    $eqCards += "<div class='scard sc-eq'><div class='lbl'>License Host</div><div class='val'>$(HE $data.LicenseHost)</div></div>"
    $cardsHtml = "<div class='srow-lbl'>Windows Server</div><div class='sgrid'>$winCards</div>" +
                 "<div class='srow-lbl' style='margin-top:8px'>ControlSuite</div><div class='sgrid'>$eqCards</div>"

    # ============================================================
    # WINDOWS HTML SECTIONS
    # ============================================================
    $winHtml = ''
    if ($data.WinData) {
        $wd = $data.WinData

        # System Identity
        $sysRows = [System.Collections.Generic.List[object]]::new()
        @(
            @{K='Computer Name';  V=$wd.ComputerName},
            @{K='Manufacturer';   V=$wd.Manufacturer},
            @{K='Model';          V=$wd.Model},
            @{K='Platform';       V=if($wd.Hypervisor){"$($wd.Platform) ($($wd.Hypervisor))"}else{$wd.Platform}},
            @{K='System UUID';    V=$wd.UUID},
            @{K='BIOS Version';   V=$wd.BiosVersion},
            @{K='BIOS Serial';    V=$wd.BiosSerial}
        ) | ForEach-Object { $sysRows.Add([PSCustomObject]@{K=$_.K;V=(FV $_.V)}) }
        $winHtml += New-HtmlSection 'win-system' 'System Identity' (New-KVTable $sysRows)

        # OS
        $osRows = [System.Collections.Generic.List[object]]::new()
        @(
            @{K='OS Name';       V=$wd.OsName},
            @{K='OS Version';    V=$wd.OsVersion},
            @{K='Architecture';  V=$wd.OsArch},
            @{K='Install Date';  V=$wd.OsInstall}
        ) | ForEach-Object { $osRows.Add([PSCustomObject]@{K=$_.K;V=(FV $_.V)}) }
        $winHtml += New-HtmlSection 'win-os' 'Operating System' (New-KVTable $osRows)

        # Hardware & Storage
        $hwRows = [System.Collections.Generic.List[object]]::new()
        $hwRows.Add([PSCustomObject]@{K='RAM';     V=(FV $wd.Ram)})
        $hwRows.Add([PSCustomObject]@{K='CPU';     V=(FV $wd.CpuName)})
        $hwRows.Add([PSCustomObject]@{K='Cores';   V=(FV $wd.CpuCores)})
        $hwRows.Add([PSCustomObject]@{K='Threads'; V=(FV $wd.CpuThreads)})
        foreach ($d2 in $wd.Drives) {
            $hwRows.Add([PSCustomObject]@{K="Drive $($d2.Name):"; V="$($d2.Total) GB total  |  $($d2.Used) GB used  |  $($d2.Free) GB free"})
        }
        $winHtml += New-HtmlSection 'win-hw' 'Hardware and Storage' (New-KVTable $hwRows)

        # Network
        $netBody = ''
        foreach ($nic in $wd.Nics) {
            $nicRows = [System.Collections.Generic.List[object]]::new()
            @(
                @{K='IP Address';  V=$nic.IP},
                @{K='Subnet';      V=$nic.Subnet},
                @{K='Gateway';     V=$nic.Gateway},
                @{K='DNS';         V=$nic.DNS},
                @{K='DHCP';        V=$nic.DHCP},
                @{K='MAC Address'; V=$nic.Mac},
                @{K='Link Speed';  V=$nic.Speed},
                @{K='Status';      V=$nic.Status},
                @{K='NetBIOS';     V=$nic.NetBIOS},
                @{K='DNS Domain';  V=$nic.DNSDomain}
            ) | Where-Object { $_.V } | ForEach-Object { $nicRows.Add([PSCustomObject]@{K=$_.K;V=(FV $_.V)}) }
            $netBody += "<div class='sub'>$(HE $nic.Desc)</div>" + (New-KVTable $nicRows)
        }
        $winHtml += New-HtmlSection 'win-net' 'Network' $netBody

        # Domain
        $domRows = [System.Collections.Generic.List[object]]::new()
        $domRows.Add([PSCustomObject]@{K='Domain Joined'; V=if($wd.DomainJoined){"Yes ($($wd.DomainName))"}else{'No'}})
        $domRows.Add([PSCustomObject]@{K='Domain Name';   V=(FV $wd.DomainName)})
        $winHtml += New-HtmlSection 'win-domain' 'Domain' (New-KVTable $domRows)

        # Installed PM Software
        $swBody = ''
        if ($wd.PmApps.Count -eq 0) {
            $swBody = "<p style='color:#aaa;padding:8px'>No print management software detected.</p>"
        } else {
            foreach ($app in $wd.PmApps) {
                $appRows = [System.Collections.Generic.List[object]]::new()
                if ($app.Version)   { $appRows.Add([PSCustomObject]@{K='Version';  V=$app.Version}) }
                if ($app.SvcStatus) { $appRows.Add([PSCustomObject]@{K='Service';  V=$app.SvcStatus}) }
                if ($app.Path)      { $appRows.Add([PSCustomObject]@{K='Path';     V=$app.Path}) }
                $swBody += "<div class='sub'>$(HE $app.Name)</div>" + (New-KVTable $appRows)
            }
        }
        $winHtml += New-HtmlSection 'win-sw' 'Installed Print Management Software' $swBody

        # Windows Roles & Features
        $roleRows = [System.Collections.Generic.List[object]]::new()
        $roleRows.Add([PSCustomObject]@{K='.NET 3.5';        V=(FV $wd.DotNet35)})
        $roleRows.Add([PSCustomObject]@{K='.NET 4.8';        V=(FV $wd.DotNet48)})
        $roleRows.Add([PSCustomObject]@{K='Print Server Role';V=(FV $wd.PrintRoleStatus)})
        $roleRows.Add([PSCustomObject]@{K='Spooler Service'; V=(FV $wd.SpoolerStatus)})
        $roleRows.Add([PSCustomObject]@{K='Spool Folder';    V=(FV $wd.SpoolFolder)})
        $roleBody = (New-KVTable $roleRows)
        if ($wd.InstalledRoles.Count -gt 0) {
            $rr=[System.Collections.Generic.List[object]]::new()
            $wd.InstalledRoles | Sort-Object | ForEach-Object { $rr.Add([PSCustomObject]@{K='Role';V=$_}) }
            $roleBody += "<div class='sub'>Installed Roles</div>" + (New-KVTable $rr)
        }
        if ($wd.InstalledFeatures.Count -gt 0) {
            $fr=[System.Collections.Generic.List[object]]::new()
            $wd.InstalledFeatures | Sort-Object | ForEach-Object { $fr.Add([PSCustomObject]@{K='Feature';V=$_}) }
            $roleBody += "<div class='sub'>Installed Features</div>" + (New-KVTable $fr)
        }
        $winHtml += New-HtmlSection 'win-roles' 'Windows Roles and Features' $roleBody

        # SQL Server
        $sqlBody = ''
        if ($wd.SqlInstances.Count -eq 0) {
            $sqlBody = "<p style='color:#aaa;padding:8px'>No SQL Server instances detected.</p>"
        } else {
            foreach ($si in $wd.SqlInstances) {
                $sr=[System.Collections.Generic.List[object]]::new()
                @(
                    @{K='Edition';           V=$si.Edition},
                    @{K='Version';           V=$si.Version},
                    @{K='Service';           V=$si.Status},
                    @{K='Service Account';   V=$si.ServiceAcct},
                    @{K='Auth Mode';         V=$si.AuthMode},
                    @{K='Install Path';      V=$si.InstallPath},
                    @{K='Data Root';         V=$si.DataRoot},
                    @{K='Default Data Dir';  V=$si.DefaultData},
                    @{K='Default Log Dir';   V=$si.DefaultLog},
                    @{K='Backup Dir';        V=$si.BackupDir},
                    @{K='Error Log';         V=$si.ErrorLog},
                    @{K='Named Pipes';       V=if($si.NpEnabled){'Enabled'}else{'Disabled'}},
                    @{K='TCP/IP';            V=if($si.TcpEnabled){'Enabled'}else{'Disabled'}},
                    @{K='TCP Port (IPAll)';  V=$si.TcpPort},
                    @{K='TCP Keep Alive';    V=$si.TcpKeepAlive},
                    @{K='TCP Listen All';    V=$si.TcpListenAll},
                    @{K='FILESTREAM';        V=$si.FsDesc},
                    @{K='Max Server Memory'; V=$si.MaxMemoryMB},
                    @{K='Min Server Memory'; V=$si.MinMemoryMB},
                    @{K='User Instances';    V=$si.UserInstances}
                ) | ForEach-Object { $sr.Add([PSCustomObject]@{K=$_.K;V=(FV $_.V)}) }
                $sqlBody += "<div class='inst-hdr'>&#9670; Instance: $(HE $si.Name)</div>" + (New-KVTable $sr)
                if ($si.TcpIpAddrs.Count -gt 0) {
                    $ipr=[System.Collections.Generic.List[object]]::new()
                    foreach ($ip in $si.TcpIpAddrs) {
                        $ipr.Add([PSCustomObject]@{K="$($ip.Name) ($($ip.IP))"; V="Active=$($ip.Active)  Enabled=$($ip.Enabled)  Port=$(if($ip.Port){$ip.Port}else{'(none)'})"})
                    }
                    $sqlBody += "<div class='sub' style='margin-top:6px'>TCP/IP Addresses</div>" + (New-KVTable $ipr)
                }
                if ($si.SysAdmins.Count -gt 0) {
                    $ar=[System.Collections.Generic.List[object]]::new()
                    foreach ($a in $si.SysAdmins) {
                        $p=$a -split '\|',2
                        $ar.Add([PSCustomObject]@{K=$p[0];V=if($p.Count -gt 1){$p[1]}else{''}})
                    }
                    $sqlBody += "<div class='sub' style='margin-top:6px'>Sysadmin Members</div>" + (New-KVTable $ar)
                }
                if ($si.TempDbFiles.Count -gt 0) {
                    $tr=[System.Collections.Generic.List[object]]::new()
                    foreach ($tf in $si.TempDbFiles) {
                        $tr.Add([PSCustomObject]@{K="[$($tf.Type)] $($tf.Name)"; V="$($tf.Path)  |  Size: $($tf.SizeMB) MB  |  Growth: $($tf.Growth)"})
                    }
                    $sqlBody += "<div class='sub' style='margin-top:6px'>TempDB Files</div>" + (New-KVTable $tr)
                }
            }
        }
        $winHtml += New-HtmlSection 'win-sql' 'SQL Server' $sqlBody

        # Print Queues
        $qBody = "<p style='padding:4px 0 8px;color:#555;font-size:12px'>$($wd.PrintQueues.Count) queues total &nbsp;|&nbsp; FujiFilm/FF: $($wd.FfCount) &nbsp;|&nbsp; Spool: $(HE $wd.SpoolFolder)</p>"
        foreach ($q in $wd.PrintQueues) {
            $qr=[System.Collections.Generic.List[object]]::new()
            @(
                @{K='Driver';           V=$q.Driver},
                @{K='Port';             V=$q.Port},
                @{K='Status';           V=$q.Status},
                @{K='Paper Size';       V=$q.PaperSize},
                @{K='Duplex';           V=$q.Duplex},
                @{K='Color';            V=$q.Color},
                @{K='Adv Print Features';V=$q.AdvFeatures},
                @{K='Render on Client'; V=$q.RenderOnClient},
                @{K='Use App Color';    V=$q.ColorMgmt},
                @{K='ICM Method';       V=$q.IcmMethod},
                @{K='Shared';           V=if($q.Shared){"Yes (Share: $($q.ShareName))"}else{'No'}},
                @{K='Published in AD';  V=if($q.Published){'Yes'}else{'No'}},
                @{K='x86 (32-bit) Driver';V=if($q.Has32bit){'Installed'}else{'Not installed'}},
                @{K='Finisher Installed';V=$q.FinisherInst},
                @{K='Staple Capable';   V=$q.StapleInst},
                @{K='Punch Capable';    V=$q.PunchInst},
                @{K='Booklet Maker';    V=$q.BookletInst},
                @{K='Offset Stacking (hw)'; V=$q.OffsetInst}
            ) | Where-Object { $_.V -ne '' } | ForEach-Object { $qr.Add([PSCustomObject]@{K=$_.K;V=(FV $_.V 200)}) }
            $qBody += "<div class='sub'>$(HE $q.Name)</div>" + (New-KVTable $qr)
        }
        $winHtml += New-HtmlSection 'win-queues' "Print Queues ($($wd.PrintQueues.Count))" $qBody

        # Drivers & Ports
        $drvBody = ''
        if ($wd.DriversX64.Count -gt 0) {
            $dr64=[System.Collections.Generic.List[object]]::new()
            $wd.DriversX64 | Sort-Object Name | ForEach-Object { $dr64.Add([PSCustomObject]@{K='x64';V=$_.Name}) }
            $drvBody += "<div class='sub'>x64 (64-bit) Drivers</div>" + (New-KVTable $dr64)
        }
        if ($wd.DriversX86.Count -gt 0) {
            $dr86=[System.Collections.Generic.List[object]]::new()
            $wd.DriversX86 | Sort-Object Name | ForEach-Object { $dr86.Add([PSCustomObject]@{K='x86';V=$_.Name}) }
            $drvBody += "<div class='sub'>x86 (32-bit) Drivers</div>" + (New-KVTable $dr86)
        }
        if ($wd.PrinterPorts.Count -gt 0) {
            $prRows=[System.Collections.Generic.List[object]]::new()
            foreach ($pt in $wd.PrinterPorts) {
                $pv = $pt.Name
                if ($pt.IP)   { $pv += "  IP=$($pt.IP)" }
                if ($pt.Port) { $pv += "  Port=$($pt.Port)" }
                $prRows.Add([PSCustomObject]@{K='Port';V=$pv})
            }
            $drvBody += "<div class='sub'>Printer Ports ($($wd.PrinterPorts.Count))</div>" + (New-KVTable $prRows)
        }
        $winHtml += New-HtmlSection 'win-drivers' 'Printer Drivers and Ports' $drvBody
    }

    # --- Components section ---
    $compRows = [System.Collections.Generic.List[object]]::new()
    foreach ($c in $data.Components) {
        $lbl = $c.Name; if ($c.Desc) { $lbl += " ($($c.Desc))" }
        $compRows.Add([PSCustomObject]@{ K=$lbl; V="v$($c.Version)  [$($c.Date)]" })
    }
    $compSection = New-HtmlSection 'components' 'Installed Components' (New-KVTable $compRows)

    # --- Watched key sections ---
    $keyHtml = ''
    foreach ($sec in $SectionMeta.Keys) {
        if ($sec -eq 'auth') {
            $keyHtml += New-HtmlSection $sec $SectionMeta[$sec] (Build-AuthHtml $data)
            continue
        }
        if ($sec -eq 'dirsync') {
            $keyHtml += New-HtmlSection $sec $SectionMeta[$sec] (Build-DirSyncHtml $data)
            continue
        }
        if ($sec -eq 'network') {
            $keyHtml += New-HtmlSection $sec $SectionMeta[$sec] (Build-NetworkHtml $data)
            continue
        }
        $rows = [System.Collections.Generic.List[object]]::new()
        foreach ($wk in ($WatchedKeys | Where-Object { $_.Section -eq $sec })) {
            $raw = FV $data.KeyValues[$wk.Key] 300
            $v = if ($wk.ValueMap -and $wk.ValueMap.ContainsKey($raw)) { "$raw ($($wk.ValueMap[$raw]))" } else { $raw }
            $rows.Add([PSCustomObject]@{ K=$wk.Label; V=$v })
        }
        $keyHtml += New-HtmlSection $sec $SectionMeta[$sec] (New-KVTable $rows)
    }

    # --- Pricing ---
    $priceBody = ''
    if ($data.PriceLists.Count -eq 0) {
        $priceBody = "<p style='color:#aaa;padding:8px'>No price lists found.</p>"
    } else {
        foreach ($pl in $data.PriceLists) {
            $rateStr = if ($pl.Rates) { ($pl.Rates | ForEach-Object { HE $_ }) -join ' &nbsp;|&nbsp; ' } else { '(none)' }
            $priceBody += "<div class='sub'>$(HE $pl.Name) &nbsp; <small style='font-weight:normal'>[ID=$(HE $pl.Id) &nbsp; Type=$(HE $pl.Type)]</small></div>"
            $priceBody += "<table class='kv'><tr><td class='k'>Rates</td><td class='v'>$rateStr</td></tr></table>"
        }
    }
    $pricingSection = New-HtmlSection 'pricing' 'Pricing' $priceBody

    # --- Workflows ---
    $wfRows = [System.Collections.Generic.List[object]]::new()
    foreach ($wf in $data.Workflows) {
        $wfRows.Add([PSCustomObject]@{ K="[$($wf.Id)] $(HE $wf.Name)"; V=(HE $wf.Type) })
    }
    $wfBody = (New-KVTable $wfRows)
    if ($data.WFolders.Count -gt 0) {
        $folderRows = [System.Collections.Generic.List[object]]::new()
        foreach ($wff in $data.WFolders) {
            $folderRows.Add([PSCustomObject]@{ K="[$($wff.Id)]"; V=(HE $wff.Name) })
        }
        $wfBody += "<div class='sub'>Workflow Folders</div>" + (New-KVTable $folderRows)
    }
    $wfSection = New-HtmlSection 'workflows' 'Workflows' $wfBody

    # --- Pull groups ---
    $pgRows = [System.Collections.Generic.List[object]]::new()
    foreach ($pg in $data.PullGroups) {
        $pgRows.Add([PSCustomObject]@{ K="[$($pg.Id)]"; V=(HE $pg.Name) })
    }
    $pgSection = New-HtmlSection 'pullgroups' 'Pull Print Groups' (New-KVTable $pgRows)

    # --- Users ---
    $userRows = [System.Collections.Generic.List[object]]::new()
    $shown = 0
    foreach ($u in $data.Users) {
        $userRows.Add([PSCustomObject]@{ K="[$($u.Id)]"; V="$(HE $u.Domain)\$(HE $u.Name)" })
        if ((++$shown) -ge 50) { break }
    }
    $userExtra = if ($data.Users.Count -gt 50) {
        "<p style='color:#aaa;padding:6px 8px;font-size:11px'>... $($data.Users.Count - 50) more users in FULL.txt</p>"
    } else { '' }
    $userSection = New-HtmlSection 'users' "Users ($($data.Users.Count))" ((New-KVTable $userRows) + $userExtra)

    # --- Full EQVar ---
    $eqBody = ''
    $bySub = @{}
    $data.DceMap.GetEnumerator() | ForEach-Object {
        $sub = ($_.Key -split '\|', 3)[0]; if (-not $sub) { $sub = '(blank)' }
        if (-not $bySub[$sub]) { $bySub[$sub] = [System.Collections.Generic.List[object]]::new() }
        $val = $_.Value
        if ([string]::IsNullOrWhiteSpace($val)) { $val = '(empty)' }
        elseif ($val.Length -gt 150) { $val = $val.Substring(0, 150) + '...' }
        $bySub[$sub].Add([PSCustomObject]@{ K=$_.Key; V=$val })
    }
    foreach ($sub in ($bySub.Keys | Sort-Object)) {
        $eqBody += "<div class='sub'>SubSystem: $(HE $sub)</div>" + (New-KVTable $bySub[$sub])
    }
    $eqSection = New-HtmlSection 'eqvar' 'Full EQVar Configuration (DCE_config.db3)' $eqBody

    $formatted = $winHtml + $compSection + $keyHtml + $pricingSection + $wfSection + $pgSection + $userSection + $eqSection

    # Build CSS and JS inline
    $css = @'
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Consolas,'Courier New',monospace;background:#eef0f4;color:#1a1a2e;font-size:13px}
#hdr{background:#1a2744;color:#fff;padding:12px 22px;display:flex;align-items:center;gap:14px;position:sticky;top:0;z-index:200;box-shadow:0 2px 6px rgba(0,0,0,.4)}
#hdr h1{font-size:15px;font-weight:bold;white-space:nowrap}
#hdr .meta{font-size:11px;color:#9ab;flex:1}
.badge{border-radius:4px;padding:3px 10px;font-size:11px;font-weight:bold;color:#fff}
#sbar{background:#253060;padding:7px 22px;display:flex;align-items:center;gap:10px}
#sbar input{width:300px;padding:6px 10px;border-radius:4px;border:none;font-size:13px}
#sbar .hint{font-size:11px;color:#8ab}
#layout{display:flex;height:calc(100vh - 80px)}
#nav{width:200px;background:#fff;border-right:1px solid #d4d8e2;overflow-y:auto;padding:6px 0;flex-shrink:0}
#nav a{display:block;padding:5px 12px;color:#444;text-decoration:none;font-size:12px;border-left:3px solid transparent;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#nav a:hover{background:#eef2ff;color:#1a56db;border-left-color:#1a56db}
#nav a.act{background:#e8f0fe;color:#1a56db;border-left-color:#1a56db;font-weight:bold}
#nav .grp{font-size:10px;font-weight:bold;color:#bbb;padding:10px 12px 2px;text-transform:uppercase;letter-spacing:.5px}
#main{flex:1;overflow-y:auto;padding:14px 22px}
.tabs{display:flex;gap:3px;margin-bottom:12px;border-bottom:2px solid #d0d5e0}
.tab{padding:7px 16px;cursor:pointer;font-size:12px;border-radius:4px 4px 0 0;background:#d8dde8;color:#555;user-select:none}
.tab.act{background:#fff;border-bottom:3px solid #1a56db;color:#1a56db;font-weight:bold;margin-bottom:-2px}
#pf,#pr{display:none}
#pf.on,#pr.on{display:block}
.sum{background:#1a2744;color:#fff;border-radius:6px;padding:12px 16px;margin-bottom:12px}
.sum h2{font-size:11px;color:#8ab;margin-bottom:8px;text-transform:uppercase;letter-spacing:.5px}
.sgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:7px}
.scard{background:rgba(255,255,255,.08);border-radius:4px;padding:7px 10px}
.scard .lbl{font-size:10px;color:#8ab;text-transform:uppercase}
.scard .val{font-size:14px;font-weight:bold;margin-top:2px;word-break:break-all}
.sec{background:#fff;border-radius:6px;margin-bottom:9px;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.sec-hdr{padding:10px 14px;cursor:pointer;display:flex;align-items:center;justify-content:space-between;border-radius:6px;user-select:none}
.sec-hdr:hover{background:#f4f6ff}
.sec-hdr h2{font-size:13px;font-weight:bold;color:#1a2744}
.tog{font-size:15px;color:#aaa;width:18px;text-align:center;flex-shrink:0}
.sec-body{padding:10px 14px 14px;border-top:1px solid #eaecf0}
.sec-body.hide{display:none}
table.kv{width:100%;border-collapse:collapse}
table.kv td{padding:5px 8px;vertical-align:top;border-bottom:1px solid #f0f2f6;word-break:break-all}
table.kv tr:last-child td{border-bottom:none}
table.kv tr:hover td{background:#f8f9fc}
table.kv .k{color:#555;width:240px;min-width:180px;font-weight:500;word-break:normal}
table.kv .v{color:#111}
table.kv .v.e{color:#bbb;font-style:italic}
.sub{margin-top:10px;margin-bottom:5px;font-size:11px;font-weight:bold;color:#666;text-transform:uppercase;letter-spacing:.3px;border-bottom:1px solid #eee;padding-bottom:3px}
.inst-hdr{margin-top:14px;margin-bottom:6px;padding:7px 12px;background:#1a2744;color:#fff;border-radius:5px;font-size:12px;font-weight:bold;letter-spacing:.4px}
.auth-grp{margin-top:12px;margin-bottom:6px;padding:5px 10px;background:#eef2ff;border-left:3px solid #1a56db;border-radius:0 4px 4px 0;font-size:11px;font-weight:bold;color:#1a3a8a;text-transform:uppercase;letter-spacing:.3px}
table.hid-tbl{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:6px}
table.hid-tbl th{background:#f0f4f8;color:#555;font-size:10px;text-transform:uppercase;letter-spacing:.3px;padding:5px 10px;border:1px solid #e0e4ea;text-align:left;font-weight:600}
table.hid-tbl td{padding:5px 10px;border:1px solid #e0e4ea;color:#111;vertical-align:top}
table.hid-tbl tr:nth-child(even) td{background:#fafbfc}
table.kv td.sub1{padding-left:28px;background:#f9fbff;color:#444;border-left:3px solid #c8d8f4}
table.kv td.sub2{padding-left:48px;background:#fcfcff;color:#555;border-left:3px solid #e4eaf8;font-style:italic}
table.kv td.sec-lbl{font-size:10px;font-weight:700;color:#3a4a6a;text-transform:uppercase;letter-spacing:.4px;padding:10px 8px 4px;border:none;background:#eef3fb;border-bottom:1px solid #d4dff0}
.dc-on{color:#1a6d1a;font-weight:600}
.dc-off{color:#888}
.eq-on{color:#1a6d1a}
.eq-off{color:#9ea5b0}
table.net-tbl{width:100%;border-collapse:collapse;font-size:12px;margin:6px 0 10px}
table.net-tbl th{background:#eef3fb;color:#3a4a6a;font-size:10px;text-transform:uppercase;letter-spacing:.3px;padding:6px 10px;border:1px solid #d4dff0;text-align:left;font-weight:700}
table.net-tbl td{padding:5px 10px;border:1px solid #e8edf6;color:#222;font-size:12px}
table.net-tbl tr:nth-child(even) td{background:#f9fbff}
details.net-detail{margin:4px 0;border:1px solid #dde4f0;border-radius:4px;overflow:hidden}
details.net-detail summary{padding:7px 12px;cursor:pointer;font-size:12px;font-weight:600;color:#2a3a6a;background:#f4f7fd;user-select:none}
details.net-detail summary:hover{background:#eaeff9}
details.net-detail[open] summary{border-bottom:1px solid #dde4f0}
details.net-detail table.kv{margin:0}
pre#rawpre{background:#1e1e1e;color:#d4d4d4;padding:14px;border-radius:6px;font-size:11px;line-height:1.6;white-space:pre-wrap;word-break:break-all}
mark{background:#ffd600;color:#000;border-radius:2px}
.hi{display:none!important}
.srow-lbl{font-size:10px;font-weight:bold;color:#6ab;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}
.sc-win{background:rgba(100,180,255,.12)!important;border-left:2px solid #4fa8e8}
.sc-eq{background:rgba(255,200,100,.10)!important;border-left:2px solid #e8a84f}
'@

    $js = @'
function tog(id){
  var b=document.getElementById('b-'+id),t=document.getElementById('t-'+id);
  var c=b.classList.toggle('hide'); t.textContent=c?'+':'-';
}
function setTab(t){
  document.querySelectorAll('.tab').forEach(function(e,i){e.classList.toggle('act',(t==='f'&&i===0)||(t==='r'&&i===1));});
  document.getElementById('pf').classList.toggle('on',t==='f');
  document.getElementById('pr').classList.toggle('on',t==='r');
}
var qi=document.getElementById('q');
qi.addEventListener('input',function(){ doSearch(this.value.trim().toLowerCase()); });
function doSearch(q){
  clearMarks();
  var secs=document.querySelectorAll('.sec');
  if(!q){ secs.forEach(function(s){s.classList.remove('hi'); var b=s.querySelector('.sec-body'); if(b) b.classList.remove('hide');}); return; }
  secs.forEach(function(s){
    var m=s.textContent.toLowerCase().indexOf(q)>=0;
    s.classList.toggle('hi',!m);
    if(m){ var b=s.querySelector('.sec-body'); if(b) b.classList.remove('hide'); }
  });
  markText(document.getElementById('pf'),q);
}
function clearMarks(){ document.querySelectorAll('mark').forEach(function(m){ var p=m.parentNode; p.replaceChild(document.createTextNode(m.textContent),m); p.normalize(); }); }
function markText(node,q){
  if(node.nodeType===3){ var i=node.textContent.toLowerCase().indexOf(q); if(i>=0){ var mk=document.createElement('mark'),a=node.splitText(i); a.splitText(q.length); mk.appendChild(a.cloneNode(true)); a.parentNode.replaceChild(mk,a); } }
  else if(node.nodeType===1&&!/^(SCRIPT|STYLE|PRE)$/.test(node.nodeName)){ Array.from(node.childNodes).forEach(function(c){markText(c,q);}); }
}
var mainEl=document.getElementById('main');
function secTop(el){
  // offset of el's top edge relative to #main's scrollable content
  return el.getBoundingClientRect().top - mainEl.getBoundingClientRect().top + mainEl.scrollTop;
}
document.querySelectorAll('#nav a[href^="#"]').forEach(function(a){
  a.addEventListener('click',function(e){
    e.preventDefault();
    var target=document.getElementById(this.getAttribute('href').slice(1));
    if(!target) return;
    mainEl.scrollTop=secTop(target);
    document.querySelectorAll('#nav a').forEach(function(x){x.classList.remove('act');});
    this.classList.add('act');
  });
});
mainEl.addEventListener('scroll',function(){
  var top=mainEl.scrollTop+60;
  document.querySelectorAll('.sec').forEach(function(s){
    var a=document.querySelector('#nav a[href="#'+s.id+'"]');
    if(!a) return;
    var st=secTop(s);
    a.classList.toggle('act', top>=st && top<st+s.offsetHeight);
  });
});
'@

    $html = "<!DOCTYPE html>`n<html lang='en'>`n<head>`n<meta charset='UTF-8'>`n" +
            "<title>$(HE $data.AppName) Config - $(HE $data.Hostname)</title>`n" +
            "<style>`n$css`n</style>`n</head>`n<body>`n" +
            "<div id='hdr'>" +
              "<h1>&#9881;&nbsp;$(HE $data.AppName) Configuration</h1>" +
              "<div class='meta'>$(HE $data.Hostname) &bull; $(HE $data.Timestamp)</div>" +
              "<span class='badge' style='background:$badgeColor'>$mode</span>" +
            "</div>`n" +
            "<div id='sbar'>" +
              "<input type='text' id='q' placeholder='&#128269; Search settings, keys, values...'>" +
              "<span class='hint'>Searches all sections</span>" +
            "</div>`n" +
            "<div id='layout'>`n<nav id='nav'>$nav</nav>`n<div id='main'>`n" +
            "<div class='tabs'>" +
              "<div class='tab act' onclick=""setTab('f')"">&#9776; Formatted</div>" +
              "<div class='tab' onclick=""setTab('r')"">&#128196; Raw Text</div>" +
            "</div>`n" +
            "<div id='pf' class='on'>`n" +
              "<div class='sum'><h2>Summary</h2>$cardsHtml</div>`n" +
              $formatted + "`n</div>`n" +
            "<div id='pr'><pre id='rawpre'>$(HE $rawText)</pre></div>`n" +
            "</div></div>`n" +
            "<script>`n$js`n</script>`n</body></html>"

    $outFile = Join-Path $OutDir "${mode}_REPORT.html"
    [System.IO.File]::WriteAllText($outFile, $html, [System.Text.Encoding]::UTF8)
    Write-Host "  -> $outFile" -ForegroundColor Green
    return $outFile
}

# ============================================================
# COMPARE LOGIC
# ============================================================
function Find-LatestFolder {
    param([string]$Suffix)
    $all = Get-ChildItem $OutputRoot -Directory -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -match "_${Suffix}$" } |
           Sort-Object LastWriteTime -Descending
    if ($all) { return $all[0].FullName } else { return $null }
}

function Invoke-Compare {
    param([string]$BeforeDir, [string]$AfterDir, [string]$OutDir)

    Write-Host "  Before : $BeforeDir" -ForegroundColor Yellow
    Write-Host "  After  : $AfterDir" -ForegroundColor Green

    $lines = [System.Collections.Generic.List[string]]::new()
    function CW  { param([string]$l='') $lines.Add($l); Write-Host $l }
    function CH  { param([string]$h) CW ''; CW ('='*70); CW "  $h"; CW ('='*70) }
    function CWS { param([string]$s) CW ''; CW "  --- $s ---" }

    CW "Configuration Comparison Report"
    CW "Before : $(Split-Path $BeforeDir -Leaf)"
    CW "After  : $(Split-Path $AfterDir -Leaf)"
    CW "Run at : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    # EQVar diff
    CH "EQVAR CHANGES (DCE_config.db3)"
    $bEQ = Load-TsvMap (Join-Path $BeforeDir 'eqvar_dce.tsv')
    $aEQ = Load-TsvMap (Join-Path $AfterDir  'eqvar_dce.tsv')
    $anyEQ = $false
    ($bEQ.Keys + $aEQ.Keys) | Sort-Object -Unique | ForEach-Object {
        $k = $_
        if ($EQVarNoise -contains $k) { return }
        $inB = $bEQ.ContainsKey($k); $inA = $aEQ.ContainsKey($k)
        if ($inB -and $inA) {
            if ($bEQ[$k] -ne $aEQ[$k]) {
                $bv = $bEQ[$k]; if ($bv.Length -gt 200) { $bv = $bv.Substring(0,200)+'...' }
                $av = $aEQ[$k]; if ($av.Length -gt 200) { $av = $av.Substring(0,200)+'...' }
                CW "  ~ $k"; CW "      WAS: $bv"; CW "      NOW: $av"
                $anyEQ = $true
            }
        } elseif ($inA) {
            $av = $aEQ[$k]; if ($av.Length -gt 200) { $av = $av.Substring(0,200)+'...' }
            CW "  + $k = $av"; $anyEQ = $true
        } else {
            $bv = $bEQ[$k]; if ($bv.Length -gt 200) { $bv = $bv.Substring(0,200)+'...' }
            CW "  - $k = $bv"; $anyEQ = $true
        }
    }
    if (-not $anyEQ) { CW "  No EQVar changes." }

    # BCP content diff
    CH "SQL TABLE CONTENT CHANGES"
    $diffTables = @('cas_config','cas_scan_alias','cat_pricelist','cas_pullgroups',
                    'cas_workflow_folders','cas_user_ext','cat_validation','cas_prq_device_ext')
    $anyBcp = $false
    foreach ($tbl in $diffTables) {
        $bFile = Join-Path $BeforeDir "bcp_$tbl.txt"
        $aFile = Join-Path $AfterDir  "bcp_$tbl.txt"
        if (-not (Test-Path $bFile) -and -not (Test-Path $aFile)) { continue }
        $bLines = if (Test-Path $bFile) { [System.IO.File]::ReadAllLines($bFile) } else { @() }
        $aLines = if (Test-Path $aFile) { [System.IO.File]::ReadAllLines($aFile) } else { @() }
        $bMap = @{}; foreach ($l in $bLines) { $k=($l -split '\|')[0]; $bMap[$k]=$l }
        $aMap = @{}; foreach ($l in $aLines) { $k=($l -split '\|')[0]; $aMap[$k]=$l }
        $added   = $aMap.Keys | Where-Object { -not $bMap.ContainsKey($_) } | Sort-Object
        $removed = $bMap.Keys | Where-Object { -not $aMap.ContainsKey($_) } | Sort-Object
        $updated = $aMap.Keys | Where-Object { $bMap.ContainsKey($_) -and $bMap[$_] -ne $aMap[$_] } | Sort-Object
        if ($added.Count -eq 0 -and $removed.Count -eq 0 -and $updated.Count -eq 0) { continue }
        $anyBcp = $true
        CWS $tbl
        foreach ($k in $updated) {
            $bv=$bMap[$k]; if($bv.Length -gt 200){$bv=$bv.Substring(0,200)+'...'}
            $av=$aMap[$k]; if($av.Length -gt 200){$av=$av.Substring(0,200)+'...'}
            CW "  ~ row[$k]"; CW "      WAS: $bv"; CW "      NOW: $av"
        }
        foreach ($k in $added)   { $v=$aMap[$k]; if($v.Length-gt 200){$v=$v.Substring(0,200)+'...'}; CW "  + row[$k]: $v" }
        foreach ($k in $removed) { $v=$bMap[$k]; if($v.Length-gt 200){$v=$v.Substring(0,200)+'...'}; CW "  - row[$k]: $v" }
    }
    if (-not $anyBcp) { CW "  No SQL content changes." }

    # Write diff text
    $diffFile = Join-Path $OutDir 'COMPARE_DIFF.txt'
    [System.IO.File]::WriteAllLines($diffFile, $lines, [System.Text.Encoding]::ASCII)
    Write-Host "  -> $diffFile" -ForegroundColor Green

    # Write compare HTML
    $diffHtml = ''
    foreach ($line in $lines) {
        $eLine = HE $line
        $cls = if ($line -match '^\s*~') { 'chg' }
               elseif ($line -match '^\s*\+') { 'add' }
               elseif ($line -match '^\s*-')  { 'rem' }
               elseif ($line -match '^=+')    { 'hdr' }
               elseif ($line -match '^  ---') { 'sub' }
               else { '' }
        $diffHtml += if ($cls) { "<div class='$cls'>$eLine</div>" } else { "<div>$eLine</div>" }
    }

    $cHtml = "<!DOCTYPE html>`n<html lang='en'><head><meta charset='UTF-8'>" +
             "<title>Compare - $(HE $hostname)</title><style>" +
             "body{font-family:Consolas,'Courier New',monospace;background:#1a1a2e;color:#cdd;padding:20px;font-size:12px;line-height:1.7}" +
             "h1{font-size:16px;color:#fff;margin-bottom:10px}" +
             ".hdr{color:#61afef;font-weight:bold;margin-top:10px}" +
             ".sub{color:#98c379;margin-top:6px}" +
             ".chg{color:#e5c07b}" +
             ".add{color:#98c379}" +
             ".rem{color:#e06c75}" +
             "div{white-space:pre-wrap;word-break:break-all}" +
             "</style></head><body>" +
             "<h1>&#9881; Configuration Comparison &mdash; $(HE $hostname)</h1>" +
             $diffHtml + "</body></html>"

    $htmlFile = Join-Path $OutDir 'COMPARE_REPORT.html'
    [System.IO.File]::WriteAllText($htmlFile, $cHtml, [System.Text.Encoding]::UTF8)
    Write-Host "  -> $htmlFile" -ForegroundColor Green
}

# ============================================================
# MODE HANDLERS
# ============================================================
function Run-Before {
    Write-Host "`n[BEFORE] Capturing baseline..." -ForegroundColor Yellow
    $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outDir = Join-Path $OutputRoot "${hostname}_${stamp}_Before"
    Ensure-Dir $OutputRoot; Ensure-Dir $outDir

    $data = Collect-Data -SaveRawTo $outDir
    Write-FullTxt -data $data -OutDir $outDir -mode 'Before'

    Write-Host "`n[BEFORE] Complete -> $outDir" -ForegroundColor Yellow
    Write-Host "  Raw data saved. Run AFTER when you've made your UI changes.`n" -ForegroundColor Gray
}

function Run-After {
    Write-Host "`n[AFTER] Capturing current state..." -ForegroundColor Green
    $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outDir = Join-Path $OutputRoot "${hostname}_${stamp}_After"
    Ensure-Dir $OutputRoot; Ensure-Dir $outDir

    $data    = Collect-Data -SaveRawTo $outDir
    $txtFile = Write-FullTxt    -data $data -OutDir $outDir -mode 'After'
               Write-SummaryTxt   -data $data -OutDir $outDir -mode 'After'
               Write-MetadataJson -data $data -OutDir $outDir -mode 'After'
    $null    = Write-HtmlReport   -data $data -OutDir $outDir -mode 'After' -fullTxtPath $txtFile

    Write-Host "`n[AFTER] Complete -> $outDir`n" -ForegroundColor Green
    return $outDir
}

function Run-Compare {
    param([string]$ForcedAfterDir = '')
    Write-Host "`n[COMPARE] Searching for snapshots..." -ForegroundColor Magenta
    $beforeDir = Find-LatestFolder 'Before'
    $afterDir  = if ($ForcedAfterDir) { $ForcedAfterDir } else { Find-LatestFolder 'After' }

    if (-not $beforeDir) { Write-Host "  ERROR: No Before snapshot found in $OutputRoot" -ForegroundColor Red; return }
    if (-not $afterDir)  { Write-Host "  ERROR: No After snapshot found in $OutputRoot"  -ForegroundColor Red; return }

    $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outDir = Join-Path $OutputRoot "${hostname}_${stamp}_Compare"
    Ensure-Dir $outDir

    Invoke-Compare -BeforeDir $beforeDir -AfterDir $afterDir -OutDir $outDir

    Write-Host "`n[COMPARE] Complete -> $outDir`n" -ForegroundColor Magenta
}

function Run-Full {
    $afterDir = Run-After
    Run-Compare -ForcedAfterDir $afterDir
    Write-Host "[FULL] Done.`n" -ForegroundColor Cyan
}

# ============================================================
# MAIN
# ============================================================
if ($Mode -eq '') { $Mode = Show-Menu }

switch ($Mode) {
    'Before'  { Run-Before }
    'After'   { Run-After  }
    'Compare' { Run-Compare }
    'Full'    { Run-Full   }
}
