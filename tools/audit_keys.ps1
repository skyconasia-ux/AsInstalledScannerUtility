$dcePath = 'C:\Windows\System32\config\systemprofile\AppData\Local\Equitrac\Equitrac Platform Component\EQDCESrv\Cache\DCE_config.db3'
$drePath = 'C:\Windows\System32\config\systemprofile\AppData\Local\Equitrac\Equitrac Platform Component\EQDRESrv\EQSpool\DREEQVar.db3'
$sq = 'C:\Windows\System32\sqlite3.exe'

function Get-Map($db) {
    $m = @{}
    $r = & $sq $db 'SELECT Key,Value FROM EQVar;' 2>$null
    foreach ($line in $r) {
        $i = $line.IndexOf('|')
        if ($i -gt 0) { $m[$line.Substring(0,$i)] = $line.Substring($i+1) }
    }
    return $m
}

$dce = Get-Map $dcePath
$dre = Get-Map $drePath

Write-Host "DCE keys: $($dce.Count)   DRE keys: $($dre.Count)"
Write-Host ""

$keys = @(
    'cas||clientauthconfig','dce||enableswipe','dce||registerpin','dce||registerpinasalternate',
    'dce||registertwocards','dce||nosecondaryidwithswipe','dce||adminpin','dce||maxpinlength',
    'cas||encryptsecondarypin','dce||authequitraccardreg','dce||authidentityprovidercardreg',
    'dce||defaultfunction','cas||loginexpiry','ads||settingsdoc',
    'cas||smtpauthenticationsec','cas||emailserver','cas||sendemailnotif','cas||defaultfromaddress',
    'cas||jobexpirytime','cas||distributionlistjobexpirytime','cas||precision',
    'dce||offlinelifetime','dce||requeuereleasedjobsonlogout','dce||releasebehaviour','cas||escrowcfg',
    'cas||colourquota','cas||autousercolorquotalimit','dce||facaccesscolorcopy',
    'cas||bcexcludeifinsufficientfunds','cas||accenforcelimit','cas||insufficientfundsmsg','dre||colorquotamessage',
    'cas||currencyiso4217','dce||costpreview','dce||colourmultiplier','dce||oversizemultiplier',
    'dce||displaybalanceinfo','dce||displaycostinfo','dce||chargebeforecopying',
    'cas||fneserverhost','cas||fneserverport','cas||fneserverprotocol',
    'dce||defaultpagesize','dce||copiertimeout','dce||enablekeypad','dce||deviceconnecttimeout',
    'dce||enablebillablefeature','dce||displayaccountinfo','dce||promptforbillingcode'
)

foreach ($k in $keys) {
    $inDce = $dce.ContainsKey($k)
    $inDre = $dre.ContainsKey($k)
    if (-not $inDce -and -not $inDre) {
        Write-Host "MISSING : $k"
    } else {
        $val = if ($inDce) { $dce[$k] } else { $dre[$k] }
        $short = if ($val.Length -gt 70) { $val.Substring(0,70) + '...' } else { $val }
        $src = if ($inDce) { 'DCE' } else { 'DRE' }
        Write-Host "OK($src)  : $k = $short"
    }
}
