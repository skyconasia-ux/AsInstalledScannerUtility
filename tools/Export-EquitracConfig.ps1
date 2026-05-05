# Export-EquitracConfig.ps1
# Standalone collector: reads all Equitrac/ControlSuite configuration and writes
# a human-readable text report.  Run on the ControlSuite server itself.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File Export-EquitracConfig.ps1
#
# Output: results\EquitracConfig_HOSTNAME_DATE.txt  (next to this script)
#
# Requirements:
#   - PowerShell 5.1+
#   - bcp.exe   (SQL Server tools, usually in PATH with SQL Express install)
#   - sqlite3.exe at C:\Windows\System32\sqlite3.exe
#   - SQL Server Express (.\SQLExpress), database eqcas
#     Edit $SqlServer / $SqlUser / $SqlPass if different.
#
# NOTE: ASCII-only, PS 5.1 compatible.  READ-ONLY - makes no changes.

$ErrorActionPreference = 'SilentlyContinue'

# ---- Connection settings ----
$SqlServer = '.\SQLExpress'
$SqlDb     = 'eqcas'
$SqlUser   = 'sa'
$SqlPass   = 'FujiFilm_11111'

# ---- Paths ----
$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$stamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
$hostname   = $env:COMPUTERNAME
$resultsDir = Join-Path $scriptDir 'results'
if (-not (Test-Path $resultsDir)) { New-Item -ItemType Directory -Path $resultsDir | Out-Null }
$outFile    = Join-Path $resultsDir "EquitracConfig_${hostname}_${stamp}.txt"

$tmpDir = 'C:\Temp\EQ_ConfigExport'
if (-not (Test-Path $tmpDir)) { New-Item -ItemType Directory -Path $tmpDir | Out-Null }

$Sqlite3 = 'C:\Windows\System32\sqlite3.exe'
$DbBase  = 'C:\Windows\System32\config\systemprofile\AppData\Local\Equitrac\Equitrac Platform Component'

# ---- Output buffer ----
$out = [System.Collections.Generic.List[string]]::new()

function W  { param([string]$l='') $out.Add($l) }
function WH { param([string]$h)
    W ''
    W ('=' * 70)
    W "  $h"
    W ('=' * 70)
}
function WS { param([string]$s) W ''; W "  --- $s ---" }
function WF { param([string]$label,[string]$val) W "  ${label}: ${val}" }
function WL { param([string]$l) W "  $l" }

# ============================================================
# Helper: BCP dump a table, return lines array
# ============================================================
function Get-BcpLines {
    param([string]$Table)
    $f = "$tmpDir\$Table.txt"
    if (Test-Path $f) { Remove-Item $f -Force }
    $bcpArgs = @("$SqlDb..$Table", 'out', $f, '-S', $SqlServer, '-U', $SqlUser, '-P', $SqlPass, '-c', '-t', '|')
    & bcp @bcpArgs 2>$null | Out-Null
    if (-not (Test-Path $f)) { return @() }
    return [System.IO.File]::ReadAllLines($f)
}

# ============================================================
# Helper: Query SQLite EQVar database -> hashtable SubSystem|Class|Name -> Value
# ============================================================
function Get-EQVar {
    param([string]$DbPath)
    $map = @{}
    if (-not (Test-Path $DbPath)) { return $map }
    $tmp = "$tmpDir\eqvar_query.db3"
    Copy-Item $DbPath $tmp -Force
    if (Test-Path "$DbPath-wal") { Copy-Item "$DbPath-wal" "$tmp-wal" -Force }
    if (Test-Path "$DbPath-shm") { Copy-Item "$DbPath-shm" "$tmp-shm" -Force }
    $q = "SELECT SubSystem,Class,Name,Value FROM EQVar ORDER BY SubSystem,Class,Name;"
    $rows = & $Sqlite3 -separator "`t" $tmp $q 2>$null
    foreach ($line in $rows) {
        $parts = $line -split "`t", 4
        if ($parts.Count -lt 3) { continue }
        $key = "$($parts[0])|$($parts[1])|$($parts[2])"
        $map[$key] = if ($parts.Count -ge 4) { $parts[3] } else { '' }
    }
    Remove-Item "$tmp*" -Force -ErrorAction SilentlyContinue
    return $map
}

# ============================================================
# Helpers: EQVar lookup and value formatting
# ============================================================
function EV {
    param([string]$Key)
    # Check dce first, fall back to cas equivalent, then dre
    if ($dce.ContainsKey($Key)) { return $dce[$Key] }
    if ($dre.ContainsKey($Key)) { return $dre[$Key] }
    return ''
}

function FV {
    param([string]$v, [int]$Max=120)
    if ($v -eq $null -or [string]::IsNullOrWhiteSpace($v)) { return '(not set)' }
    $v = $v.Trim()
    if ($v -eq '') { return '(not set)' }
    if ($v.Length -gt $Max) { return $v.Substring(0,$Max) + '...' }
    return $v
}

# Extract a named value from a simple XML string: <tagname>VALUE</tagname>
function XV {
    param([string]$xml, [string]$tag)
    if ([string]::IsNullOrEmpty($xml)) { return '' }
    $m = [regex]::Match($xml, "<$tag>([^<]*)</$tag>")
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return ''
}

# ============================================================
# Report header
# ============================================================
W "EquitracConfig Export"
W "Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
W "Computer  : $hostname"
W "User      : $($env:USERNAME)"

# ============================================================
# Load SQLite EQVar databases
# ============================================================
$dcePath = "$DbBase\EQDCESrv\Cache\DCE_config.db3"
$drePath = "$DbBase\EQDRESrv\EQSpool\DREEQVar.db3"

$dce = Get-EQVar -DbPath $dcePath
$dre = Get-EQVar -DbPath $drePath

# ============================================================
# SECTION 1 - Installed Software Versions
# ============================================================
# cas_installedsoftware BCP format:
#   Component | Server | FQDN | Description | (empty) | Version | InstallDate
WH "INSTALLED COMPONENTS"

$swLines = Get-BcpLines -Table 'cas_installedsoftware'
if ($swLines.Count -eq 0) {
    WL "(cas_installedsoftware not available)"
} else {
    foreach ($line in $swLines) {
        $f = $line -split '\|'
        $component = $f[0].Trim()
        $version   = if ($f.Count -gt 5) { $f[5].Trim() } else { '' }
        $instDate  = if ($f.Count -gt 6) { $f[6].Trim() } else { '' }
        $desc      = if ($f.Count -gt 3) { $f[3].Trim() } else { '' }
        if ($component -ne '') {
            $line2 = "  $component"
            if ($desc -ne '') { $line2 += " ($desc)" }
            if ($version -ne '') { $line2 += "  v$version" }
            if ($instDate -ne '') { $line2 += "  [$instDate]" }
            W $line2
        }
    }
}

# ============================================================
# SECTION 2 - Authentication
# ============================================================
WH "AUTHENTICATION"

# Auth method list (XML blob - extract key settings)
$authCfg = EV 'cas||clientauthconfig'
if (-not [string]::IsNullOrEmpty($authCfg)) {
    WS "Auth Method List"
    foreach ($method in @('EquitracPINS','ExternalUserIDPassword','ExternalPassword','ExternalUserID','CardSwipe','CardAndPIN')) {
        $v = XV $authCfg $method
        if ($v -ne '') { WL "  $method : $v" }
    }
}

WS "Card / Badge"
WF "Card Swipe (enableswipe)"         (FV (EV 'dce||enableswipe'))
WF "Register PIN"                      (FV (EV 'dce||registerpin'))
WF "PIN as Alternate to Card"          (FV (EV 'dce||registerpinasalternate'))
WF "Register Two Cards"                (FV (EV 'dce||registertwocards'))
WF "No Secondary ID with Swipe"        (FV (EV 'dce||nosecondaryidwithswipe'))
WF "Admin PIN"                         (FV (EV 'dce||adminpin'))
WF "Max PIN Length"                    (FV (EV 'dce||maxpinlength'))
WF "Deny Prompt Empty Password"        (FV (EV 'dce||denypromptemptypassword'))
WF "Encrypt Secondary PIN"             (FV (EV 'cas||encryptsecondarypin'))

WS "Equitrac Auth (Equitrac card reg)"
WF "Equitrac Card Reg"                 (FV (EV 'dce||authequitraccardreg'))
WF "Identity Provider Card Reg"        (FV (EV 'dce||authidentityprovidercardreg'))

WS "Other Auth Settings"
WF "Default Function at Device"        (FV (EV 'dce||defaultfunction'))
WF "Login Expiry (seconds)"            (FV (EV 'cas||loginexpiry'))
WF "AzureAD / LDAP Sync"              (FV (EV 'ads||settingsdoc') 200)

# ============================================================
# SECTION 3 - SMTP / Email
# ============================================================
WH "SMTP / EMAIL"

# Primary SMTP config lives in cas||smtpauthenticationsec as XML
# Format: <smtpdefault><address>host:port</address><user>u</user><password>encrypted</password>...
$smtpXml = EV 'cas||smtpauthenticationsec'
if (-not [string]::IsNullOrEmpty($smtpXml)) {
    $smtpAddr = XV $smtpXml 'address'
    $smtpUser = XV $smtpXml 'user'
    $smtpTls  = XV $smtpXml 'tls'
    $smtpTitle= XV $smtpXml 'title'

    # address is usually "host:port"
    if ($smtpAddr -match '^([^:]+):(\d+)$') {
        WF "SMTP Server"      $Matches[1]
        WF "SMTP Port"        $Matches[2]
    } else {
        WF "SMTP Server/Port" (FV $smtpAddr)
    }
    WF "SMTP Username"        (FV $smtpUser)
    WF "SMTP TLS"             (FV $smtpTls)
    WF "SMTP Profile Name"    (FV $smtpTitle)
} else {
    WF "SMTP Server"  (FV (EV 'cas||emailserver'))
    WL "  (smtpauthenticationsec not configured)"
}

WF "Send Email Notifications"  (FV (EV 'cas||sendemailnotif'))
WF "Default From Address"      (FV (EV 'cas||defaultfromaddress'))

# ============================================================
# SECTION 4 - Job Management
# ============================================================
WH "JOB MANAGEMENT"

WF "Job Expiry Time (minutes)"              (FV (EV 'cas||jobexpirytime'))
WF "Distribution List Job Expiry (minutes)" (FV (EV 'cas||distributionlistjobexpirytime'))
WF "Accounting Precision (decimal places)"  (FV (EV 'cas||precision'))
WF "Login Expiry (seconds)"                 (FV (EV 'cas||loginexpiry'))
WF "Offline Lifetime"                       (FV (EV 'dce||offlinelifetime'))
WF "Requeue Released Jobs on Logout"        (FV (EV 'dce||requeuereleasedjobsonlogout'))
WF "Release Behaviour"                      (FV (EV 'dce||releasebehaviour'))
WF "Max Pull Print Timeout"                 (FV (EV 'dce||pulltimeout'))

# Escrow config
$escrow = EV 'cas||escrowcfg'
if (-not [string]::IsNullOrEmpty($escrow)) {
    WF "Escrow Enabled"         (XV $escrow 'escrow_enabled')
    WF "Escrow Expiry (minutes)"(XV $escrow 'expiration_mins')
    WF "Escrow Offline Mode"    (XV $escrow 'offline_mode')
}

# ============================================================
# SECTION 5 - Quotas and Messages
# ============================================================
WH "QUOTAS AND MESSAGES"

WF "Colour Quota Type"             (FV (EV 'cas||colourquota'))
WF "Auto User Color Quota Limit"   (FV (EV 'cas||autousercolorquotalimit'))
WF "Auto User Hard Limit"          (FV (EV 'cas||autouserhardlimit'))
WF "Quota Enforcement"             (FV (EV 'cas||accenforcelimit'))
WF "Insufficient Funds Message"    (FV (EV 'cas||insufficientfundsmsg'))
WF "Color Quota Message"           (FV (EV 'cas||colorquotamessage') 200)

# ============================================================
# SECTION 6 - Currency and Accounting
# ============================================================
WH "CURRENCY AND ACCOUNTING"

WF "Currency (ISO 4217)"           (FV (EV 'cas||currencyiso4217'))
WF "Cost Preview"                  (FV (EV 'dce||costpreview'))
WF "Colour Multiplier"             (FV (EV 'dce||colourmultiplier'))
WF "Oversize Multiplier"           (FV (EV 'dce||oversizemultiplier'))
WF "Display Balance Info"          (FV (EV 'dce||displaybalanceinfo'))
WF "Display Cost Info"             (FV (EV 'dce||displaycostinfo'))
WF "Charge Before Copying"         (FV (EV 'dce||chargebeforecopying'))

# ============================================================
# SECTION 7 - Pricing
# ============================================================
WH "PRICING (cat_pricelist)"

# cat_pricelist BCP format:
#   ID | Name | Description | Type | XMLBlob  (XMLBlob may span multiple BCP lines)
$plLines = Get-BcpLines -Table 'cat_pricelist'
if ($plLines.Count -eq 0) {
    WL "(no price lists found or BCP unavailable)"
} else {
    $currentId   = ''
    $currentName = ''
    $currentType = ''
    $currentXml  = [System.Text.StringBuilder]::new()
    $inRecord    = $false

    foreach ($line in $plLines) {
        if ($line -match '^(\d+)\|') {
            # Flush previous record
            if ($inRecord -and $currentName -ne '') {
                WS "Price List: $currentName  [ID=$currentId  Type=$currentType]"
                $xml    = $currentXml.ToString()
                $ranges = [regex]::Matches($xml, 'rate="([^"]+)"')
                if ($ranges.Count -eq 0) {
                    WL "  (no rate data found)"
                } else {
                    $rateList = $ranges | ForEach-Object { $_.Groups[1].Value }
                    $rateList = $rateList | Select-Object -Unique | Sort-Object { [double]$_ }
                    WL "  Rates: $($rateList -join ' | ')"
                }
            }
            $fields = $line -split '\|', 5
            $currentId   = $fields[0].Trim()
            $currentName = if ($fields.Count -gt 1) { $fields[1].Trim() } else { '' }
            $currentType = if ($fields.Count -gt 3) { $fields[3].Trim() } else { '' }
            $currentXml  = [System.Text.StringBuilder]::new()
            if ($fields.Count -gt 4) { [void]$currentXml.Append($fields[4]) }
            $inRecord    = $true
        } elseif ($inRecord) {
            [void]$currentXml.AppendLine($line)
        }
    }
    # Flush last
    if ($inRecord -and $currentName -ne '') {
        WS "Price List: $currentName  [ID=$currentId  Type=$currentType]"
        $xml    = $currentXml.ToString()
        $ranges = [regex]::Matches($xml, 'rate="([^"]+)"')
        if ($ranges.Count -eq 0) {
            WL "  (no rate data found)"
        } else {
            $rateList = $ranges | ForEach-Object { $_.Groups[1].Value }
            $rateList = $rateList | Select-Object -Unique | Sort-Object { [double]$_ }
            WL "  Rates: $($rateList -join ' | ')"
        }
    }
}

# ============================================================
# SECTION 8 - Workflows / Scan Destinations
# ============================================================
WH "WORKFLOWS (cas_scan_alias)"

# cas_scan_alias BCP format (discovered via testing):
#   ID | Name | LastModDate | ScopeType | XMLBlob
$wfLines = Get-BcpLines -Table 'cas_scan_alias'
if ($wfLines.Count -eq 0) {
    WL "(no scan aliases found)"
} else {
    foreach ($line in $wfLines) {
        if (-not ($line -match '^\d+\|')) { continue }
        $f    = $line -split '\|', 5
        $id   = $f[0].Trim()
        $name = if ($f.Count -gt 1) { $f[1].Trim() } else { '' }
        $mod  = if ($f.Count -gt 2) { $f[2].Trim() } else { '' }
        # Extract scan type from XML blob
        $xml  = if ($f.Count -gt 4) { $f[4] } else { '' }
        $stype = XV $xml 'scan_alias_type'
        $stypeName = switch ($stype) {
            '5' { 'Copy' }
            '6' { 'Scan to Email' }
            '7' { 'Fax' }
            '8' { 'Print to Me' }
            '9' { 'Release All' }
            '10'{ 'Scan to Folder' }
            '11'{ 'Scan to FTP' }
            '12'{ 'Scan to USB' }
            default { if ($stype -ne '') { "Type $stype" } else { '' } }
        }
        $typeStr = if ($stypeName -ne '') { "  [$stypeName]" } else { '' }
        WL "  [$id] $name$typeStr   (modified: $mod)"
    }
}

WS "Workflow Folders (cas_workflow_folders)"
$wfFolderLines = Get-BcpLines -Table 'cas_workflow_folders'
if ($wfFolderLines.Count -eq 0) {
    WL "(no workflow folders found)"
} else {
    foreach ($line in $wfFolderLines) {
        $f    = $line -split '\|'
        $id   = $f[0].Trim()
        $name = if ($f.Count -gt 1) { $f[1].Trim() } else { '' }
        WL "  [$id] $name"
    }
}

# ============================================================
# SECTION 9 - Pull Print Groups
# ============================================================
WH "PULL PRINT GROUPS (cas_pullgroups)"

$pgLines = Get-BcpLines -Table 'cas_pullgroups'
if ($pgLines.Count -eq 0) {
    WL "(no pull groups found)"
} else {
    foreach ($line in $pgLines) {
        $f    = $line -split '\|'
        $id   = $f[0].Trim()
        $name = if ($f.Count -gt 1) { $f[1].Trim() } else { '' }
        WL "  [$id] $name"
    }
}

# ============================================================
# SECTION 10 - Users
# ============================================================
WH "USER ACCOUNTS (cas_user_ext)"

$userLines = Get-BcpLines -Table 'cas_user_ext'
if ($userLines.Count -eq 0) {
    WL "(no users found)"
} else {
    $count = 0
    foreach ($line in $userLines) {
        $f      = $line -split '\|'
        $id     = $f[0].Trim()
        $name   = if ($f.Count -gt 1) { $f[1].Trim() } else { '' }
        $domain = if ($f.Count -gt 2) { $f[2].Trim() } else { '' }
        $email  = if ($f.Count -gt 3) { $f[3].Trim() } else { '' }
        WL "  [$id] $domain\$name  <$email>"
        $count++
        if ($count -ge 50) { WL "  ... ($($userLines.Count - 50) more users - see full BCP dump)"; break }
    }
}

# ============================================================
# SECTION 11 - Device Settings
# ============================================================
WH "DEVICE SETTINGS (cas_prq_device_ext)"

WF "Default Page Size"             (FV (EV 'dce||defaultpagesize'))
WF "Copier Timeout (ms)"           (FV (EV 'dce||copiertimeout'))
WF "Message Pause Time (ms)"       (FV (EV 'dce||messagepausetime'))
WF "Enable Keypad"                 (FV (EV 'dce||enablekeypad'))
WF "Device Connect Timeout (ms)"   (FV (EV 'dce||deviceconnecttimeout'))
WF "Billable Feature Enabled"      (FV (EV 'dce||enablebillablefeature'))
WF "Display Account Info"          (FV (EV 'dce||displayaccountinfo'))
WF "Prompt for Billing Code"       (FV (EV 'dce||promptforbillingcode'))

# ============================================================
# SECTION 12 - License Server
# ============================================================
WH "LICENSE SERVER"

WF "FNE Server Host"               (FV (EV 'cas||fneserverhost'))
WF "FNE Server Port"               (FV (EV 'cas||fneserverport'))
WF "FNE Server Protocol"           (FV (EV 'cas||fneserverprotocol'))

# Registry FlexNet info
$flexReg = 'HKLM:\SOFTWARE\FLEXlm License Manager'
if (Test-Path $flexReg) {
    WS "FlexNet License Manager (registry)"
    try {
        $item = Get-Item $flexReg -ErrorAction Stop
        foreach ($vn in $item.GetValueNames()) {
            $vd = $item.GetValue($vn); $vs = [string]$vd
            if ($vs.Length -gt 100) { $vs = $vs.Substring(0,100)+'...' }
            WL "  $vn = $vs"
        }
    } catch {}
}

# ============================================================
# SECTION 13 - Full EQVar Dump (grouped by SubSystem)
# ============================================================
WH "FULL EQVAR CONFIGURATION (DCE_config.db3)"

if ($dce.Count -eq 0) {
    WL "(DCE_config.db3 not found or empty)"
} else {
    $bySub = @{}
    $dce.GetEnumerator() | ForEach-Object {
        $parts = $_.Key -split '\|', 3
        $sub   = if ($parts[0] -ne '') { $parts[0] } else { '(blank)' }
        if (-not $bySub.ContainsKey($sub)) { $bySub[$sub] = [System.Collections.Generic.List[string]]::new() }
        $val = $_.Value
        if ([string]::IsNullOrWhiteSpace($val)) { $val = '(empty)' }
        elseif ($val.Length -gt 150) { $val = $val.Substring(0,150) + '...' }
        $bySub[$sub].Add("  $($_.Key) = $val")
    }
    foreach ($sub in ($bySub.Keys | Sort-Object)) {
        WS "SubSystem: $sub"
        $bySub[$sub] | Sort-Object | ForEach-Object { W $_ }
    }
}

WH "FULL EQVAR CONFIGURATION (DREEQVar.db3) - DRE-exclusive keys only"

if ($dre.Count -eq 0) {
    WL "(DREEQVar.db3 not found or empty)"
} else {
    $dreOnly = $dre.GetEnumerator() | Where-Object { -not $dce.ContainsKey($_.Key) }
    $bySub2 = @{}
    $dreOnly | ForEach-Object {
        $parts = $_.Key -split '\|', 3
        $sub   = if ($parts[0] -ne '') { $parts[0] } else { '(blank)' }
        if (-not $bySub2.ContainsKey($sub)) { $bySub2[$sub] = [System.Collections.Generic.List[string]]::new() }
        $val = $_.Value
        if ([string]::IsNullOrWhiteSpace($val)) { $val = '(empty)' }
        elseif ($val.Length -gt 150) { $val = $val.Substring(0,150) + '...' }
        $bySub2[$sub].Add("  $($_.Key) = $val")
    }
    if ($bySub2.Count -eq 0) {
        WL "(all keys mirrored in DCE_config - no DRE-exclusive keys)"
    } else {
        foreach ($sub in ($bySub2.Keys | Sort-Object)) {
            WS "SubSystem: $sub"
            $bySub2[$sub] | Sort-Object | ForEach-Object { W $_ }
        }
    }
}

# ============================================================
# Write output file
# ============================================================
[System.IO.File]::WriteAllLines($outFile, $out, [System.Text.Encoding]::ASCII)

Write-Host ""
Write-Host "EquitracConfig export complete."
Write-Host "Output: $outFile"
Write-Host "Lines : $($out.Count)"

# Cleanup temp
Remove-Item "$tmpDir\*.txt" -Force -ErrorAction SilentlyContinue
Remove-Item "$tmpDir\*.db3*" -Force -ErrorAction SilentlyContinue
