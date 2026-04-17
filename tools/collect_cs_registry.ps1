# ControlSuite / Equitrac / Nuance / Kofax / Tungsten -- Registry + Service collector
# Output: C:\Temp\ControlSuite_Registry.txt
# NOTE: ASCII-only -- no box-drawing or Unicode chars (PS 5.1 ANSI compat)

$out = [System.Collections.Generic.List[string]]::new()
$ts  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

function W { param([string]$line='') $out.Add($line) }
function WH { param([string]$h) W ''; W ('=' * 72); W "  $h"; W ('=' * 72) }
function WS { param([string]$s) W ''; W "  --- $s ---" }

W "ControlSuite / Equitrac / Nuance / Kofax / Tungsten -- Registry Inspection"
W "Generated : $ts"
W "Host      : $env:COMPUTERNAME"

# Helper: recursively dump registry key, readable values only
function Dump-RegKey {
    param([string]$Path, [int]$Depth = 0)
    if (-not (Test-Path -LiteralPath $Path -EA SilentlyContinue)) { return }
    $indent = '  ' * $Depth
    W "${indent}[$Path]"
    try {
        $props = Get-ItemProperty -LiteralPath $Path -EA SilentlyContinue
        if ($props) {
            $props.PSObject.Properties |
              Where-Object { $_.Name -notlike 'PS*' } |
              ForEach-Object {
                $name = $_.Name
                $val  = $_.Value
                if ($val -is [byte[]]) {
                    if ($val.Count -le 128) {
                        $hex = ($val | ForEach-Object { '{0:X2}' -f $_ }) -join ' '
                        W "${indent}  $name = [BIN $($val.Count)B] $hex"
                    } else {
                        $str = [System.Text.Encoding]::Unicode.GetString($val).TrimEnd([char]0) -replace '[\x00-\x08\x0B\x0E-\x1F]',''
                        if ($str -match '^[\x09\x0A\x0D\x20-\x7E]+$') {
                            W "${indent}  $name = [BIN->UNICODE] $str"
                        } else {
                            $preview = ($val[0..31] | ForEach-Object { '{0:X2}' -f $_ }) -join ' '
                            W "${indent}  $name = [BIN $($val.Count)B] $preview ..."
                        }
                    }
                } elseif ($val -is [string[]]) {
                    W "${indent}  $name = [MULTI_SZ] $($val -join ' | ')"
                } elseif ($val -is [int] -or $val -is [long] -or $val -is [string]) {
                    W "${indent}  $name = $val"
                } else {
                    W "${indent}  $name = [$($val.GetType().Name)] $val"
                }
            }
        }
    } catch {}
    try {
        Get-ChildItem -LiteralPath $Path -EA SilentlyContinue | ForEach-Object {
            Dump-RegKey -Path $_.PSPath -Depth ($Depth + 1)
        }
    } catch {}
}

# ==============================================================================
WH 'SECTION 1: HKLM\SOFTWARE\Equitrac (full recursive)'
# ==============================================================================
Dump-RegKey 'HKLM:\SOFTWARE\Equitrac'

# ==============================================================================
WH 'SECTION 2: HKLM\SOFTWARE\Nuance (full recursive)'
# ==============================================================================
Dump-RegKey 'HKLM:\SOFTWARE\Nuance'

# ==============================================================================
WH 'SECTION 3: HKLM\SOFTWARE\Kofax (full recursive)'
# ==============================================================================
Dump-RegKey 'HKLM:\SOFTWARE\Kofax'

# ==============================================================================
WH 'SECTION 4: WOW6432Node brand keys'
# ==============================================================================
foreach ($brand in @('Equitrac','Nuance','Kofax','Tungsten','ControlSuite')) {
    $p = "HKLM:\SOFTWARE\WOW6432Node\$brand"
    if (Test-Path $p) { Dump-RegKey $p }
    else { W "  absent: $p" }
}

# ==============================================================================
WH 'SECTION 5: Windows Services -- Equitrac / EQ / Nuance / Kofax / ControlSuite'
# ==============================================================================
$svcKeys = Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services' -EA SilentlyContinue |
    Where-Object { $_.PSChildName -match '(?i)Equitrac|^EQ[A-Z]|Nuance|Kofax|ControlSuite|eqcas|eqlog|eqextend|eqpms|eqaud|eqdir' }

if ($svcKeys) {
    foreach ($sk in $svcKeys) {
        WS $sk.PSChildName
        Dump-RegKey $sk.PSPath -Depth 1
        $svc = Get-Service -Name $sk.PSChildName -EA SilentlyContinue
        if ($svc) {
            W "    [LIVE] Status=$($svc.Status)  StartType=$($svc.StartType)  DisplayName=$($svc.DisplayName)"
        }
    }
} else {
    W '  (no matching services found -- trying broader search)'
    # Broader: look for ImagePath containing brand keywords
    Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services' -EA SilentlyContinue | ForEach-Object {
        try {
            $imgPath = (Get-ItemProperty $_.PSPath -Name ImagePath -EA SilentlyContinue).ImagePath
            if ($imgPath -match '(?i)Equitrac|Nuance|Kofax|ControlSuite') {
                W "  Service: $($_.PSChildName)"
                W "  ImagePath: $imgPath"
            }
        } catch {}
    }
}

# ==============================================================================
WH 'SECTION 6: Installed Products -- Uninstall registry'
# ==============================================================================
@(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
) | ForEach-Object {
    Get-ChildItem $_ -EA SilentlyContinue | ForEach-Object {
        try {
            $p = Get-ItemProperty $_.PSPath -EA SilentlyContinue
            if ($p.DisplayName -match '(?i)Equitrac|Nuance|Kofax|ControlSuite|Tungsten|Omtool|Autonomy') {
                W "  Product  : $($p.DisplayName)"
                W "  Version  : $($p.DisplayVersion)"
                W "  Publisher: $($p.Publisher)"
                W "  InstDate : $($p.InstallDate)"
                W "  Location : $($p.InstallLocation)"
                W "  UninstCmd: $($p.UninstallString)"
                W "  GUID     : $($_.PSChildName)"
                W ''
            }
        } catch {}
    }
}

# ==============================================================================
WH 'SECTION 7: ODBC Data Sources (System DSNs)'
# ==============================================================================
Dump-RegKey 'HKLM:\SOFTWARE\ODBC\ODBC.INI'
Dump-RegKey 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI'

# ==============================================================================
WH 'SECTION 8: SQL Server instances on this machine'
# ==============================================================================
Dump-RegKey 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server'
Dump-RegKey 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server'

# ==============================================================================
WH 'SECTION 9: Scheduled Tasks -- brand-related'
# ==============================================================================
try {
    $tasks = Get-ScheduledTask -EA SilentlyContinue |
        Where-Object { $_.TaskName -match '(?i)Equitrac|Nuance|Kofax|ControlSuite|Tungsten|^EQ' }
    if ($tasks) {
        foreach ($t in $tasks) {
            $trigTypes = ($t.Triggers | ForEach-Object { $_.GetType().Name }) -join ', '
            $actions   = ($t.Actions  | ForEach-Object { "$($_.Execute) $($_.Arguments)".Trim() }) -join '; '
            W "  Task    : $($t.TaskPath)$($t.TaskName)"
            W "  State   : $($t.State)"
            W "  Trigger : $trigTypes"
            W "  Action  : $actions"
            W ''
        }
    } else { W '  (none found)' }
} catch { W "  (error: $($_.Exception.Message))" }

# ==============================================================================
WH 'SECTION 10: Firewall Rules -- brand-related'
# ==============================================================================
try {
    $rules = Get-NetFirewallRule -EA SilentlyContinue |
        Where-Object { $_.DisplayName -match '(?i)Equitrac|Nuance|Kofax|ControlSuite|Tungsten' }
    if ($rules) {
        foreach ($r in $rules) {
            $pf = $r | Get-NetFirewallPortFilter -EA SilentlyContinue
            $af = $r | Get-NetFirewallApplicationFilter -EA SilentlyContinue
            W "  Rule    : $($r.DisplayName)"
            W "  Dir/Act : $($r.Direction) / $($r.Action)  Enabled=$($r.Enabled)"
            if ($pf) { W "  Port    : Proto=$($pf.Protocol) Local=$($pf.LocalPort) Remote=$($pf.RemotePort)" }
            if ($af -and $af.Program -ne 'Any') { W "  Program : $($af.Program)" }
            W ''
        }
    } else { W '  (none found)' }
} catch { W "  (error: $($_.Exception.Message))" }

# ==============================================================================
WH 'SECTION 11: File system -- known installation paths'
# ==============================================================================
@(
    'C:\Program Files\Kofax',
    'C:\Program Files\Nuance',
    'C:\Program Files\Equitrac',
    'C:\Program Files (x86)\Kofax',
    'C:\Program Files (x86)\Nuance',
    'C:\Program Files (x86)\Equitrac',
    'C:\ProgramData\Kofax',
    'C:\ProgramData\Nuance',
    'C:\ProgramData\Equitrac'
) | ForEach-Object {
    if (Test-Path $_) {
        W "  EXISTS: $_"
        Get-ChildItem $_ -EA SilentlyContinue | ForEach-Object {
            $type = if ($_.PSIsContainer) { '[DIR] ' } else { '[FILE]' }
            $sz   = if (-not $_.PSIsContainer) { "  $([Math]::Round($_.Length/1KB,1)) KB" } else { '' }
            W "    $type $($_.Name)$sz"
        }
        W ''
    } else { W "  absent: $_" }
}

# ==============================================================================
WH 'SECTION 12: Recent Event Log entries -- brand-related providers'
# ==============================================================================
@('Application','System') | ForEach-Object {
    $log = $_
    try {
        $events = Get-WinEvent -LogName $log -MaxEvents 3000 -EA SilentlyContinue |
            Where-Object { $_.ProviderName -match '(?i)Equitrac|Nuance|Kofax|ControlSuite|^EQ' } |
            Select-Object -First 50
        if ($events) {
            W "  Log: $log  ($($events.Count) matching events shown)"
            $events | Group-Object ProviderName | ForEach-Object {
                W "    Provider: $($_.Name)  Count=$($_.Count)"
                $_.Group | Select-Object -First 5 | ForEach-Object {
                    # Flatten message on a single line safely
                    $msg = ($_.Message -replace '[\r\n]+',' ').Trim()
                    if ($msg.Length -gt 200) { $msg = $msg.Substring(0,197) + '...' }
                    W "      [$($_.TimeCreated -f 'yyyy-MM-dd HH:mm')] Lvl=$($_.LevelDisplayName) Id=$($_.Id) : $msg"
                }
            }
        } else {
            W "  $log : no brand-matching events in last 3000 entries"
        }
    } catch { W "  $log : error -- $($_.Exception.Message)" }
}

# ==============================================================================
WH 'SECTION 13: IIS / Web bindings (ControlSuite web front-end)'
# ==============================================================================
$iisRegPath = 'HKLM:\SOFTWARE\Microsoft\InetStp'
if (Test-Path $iisRegPath) {
    Dump-RegKey $iisRegPath
    try {
        Import-Module WebAdministration -EA Stop
        Get-Website -EA SilentlyContinue | ForEach-Object {
            $site = $_
            W "  Site : $($site.Name)  State=$($site.State)  PhysPath=$($site.PhysicalPath)"
            $site.Bindings.Collection | ForEach-Object {
                W "    Binding: $($_.Protocol) $($_.bindingInformation)"
            }
        }
        Get-WebApplication -EA SilentlyContinue | ForEach-Object {
            W "  App  : $($_.Path)  Phys=$($_.PhysicalPath)"
        }
    } catch {
        W '  (WebAdministration module unavailable -- IIS reg key shown above)'
    }
} else {
    W '  IIS not detected (HKLM:\SOFTWARE\Microsoft\InetStp absent)'
}

# ==============================================================================
WH 'SECTION 14: ControlSuite / Equitrac -- additional known registry paths'
# ==============================================================================
# These paths appear in various CS versions (Equitrac Express, Office, ControlSuite)
@(
    'HKLM:\SOFTWARE\Equitrac\Express',
    'HKLM:\SOFTWARE\Equitrac\Office',
    'HKLM:\SOFTWARE\Nuance\ControlSuite',
    'HKLM:\SOFTWARE\Kofax\ControlSuite',
    'HKLM:\SOFTWARE\Kofax\AutoStore',
    'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
    'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
) | ForEach-Object {
    if (Test-Path $_) { Dump-RegKey $_ }
}

# Also check for any EQ* service parameters (database connections often stored here)
Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services' -EA SilentlyContinue |
    Where-Object { $_.PSChildName -match '^EQ' } |
    ForEach-Object {
        WS "Service params: $($_.PSChildName)"
        Dump-RegKey $_.PSPath -Depth 1
    }

# ==============================================================================
$outPath = 'C:\Temp\ControlSuite_Registry.txt'
$out | Set-Content -Path $outPath -Encoding UTF8
Write-Host "Done. $($out.Count) lines -> $outPath"
