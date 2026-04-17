# ControlSuite / Equitrac -- SQL Server database inspection
# Credentials: sa / FujiFilm_11111  (local instance, non-production)
# Output: C:\Temp\ControlSuite_DB.txt
# NOTE: ASCII-only script (PS 5.1 compatible)

$out = [System.Collections.Generic.List[string]]::new()
$ts  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

function W { param([string]$line='') $out.Add($line) }
function WH { param([string]$h) W ''; W ('=' * 72); W "  $h"; W ('=' * 72) }
function WS { param([string]$s) W ''; W "  --- $s ---" }

W "ControlSuite / Equitrac -- SQL Server Inspection"
W "Generated : $ts"
W "Host      : $env:COMPUTERNAME"

# Invoke-SQL: returns DataTable; uses comma-wrap to prevent PS pipeline enumeration
function Invoke-SQL {
    param(
        [string]$Server,
        [string]$Database = 'master',
        [string]$Query,
        [string]$User = 'sa',
        [string]$Pass = 'FujiFilm_11111',
        [int]$TimeoutSec = 30
    )
    try {
        $cs  = "Server=$Server;Database=$Database;User Id=$User;Password=$Pass;Connection Timeout=8;TrustServerCertificate=True;"
        $con = New-Object System.Data.SqlClient.SqlConnection $cs
        $con.Open()
        $cmd = $con.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSec
        $da  = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $dt  = New-Object System.Data.DataTable
        $da.Fill($dt) | Out-Null
        $con.Close()
        # Comma-wrap prevents PowerShell from enumerating the DataTable as rows
        return ,$dt
    } catch {
        return $null
    }
}

# Print a DataTable as aligned columns
function Show-Table {
    param($dt, [int]$Indent = 2, [int]$MaxRows = 500, [int]$MaxColWidth = 100)
    if ($dt -is [array]) { $dt = $dt[0] }   # unwrap comma-wrap if caller didn't
    if ($null -eq $dt -or -not ($dt -is [System.Data.DataTable]) -or $dt.Rows.Count -eq 0) {
        W (' ' * $Indent + '(no rows)')
        return
    }
    $pad = ' ' * $Indent
    $widths = @{}
    foreach ($col in $dt.Columns) { $widths[$col.ColumnName] = $col.ColumnName.Length }
    $rows = if ($dt.Rows.Count -gt $MaxRows) { $dt.Rows | Select-Object -First $MaxRows } else { $dt.Rows }
    foreach ($row in $rows) {
        foreach ($col in $dt.Columns) {
            $len = [Math]::Min("$($row[$col.ColumnName])".Length, $MaxColWidth)
            if ($len -gt $widths[$col.ColumnName]) { $widths[$col.ColumnName] = $len }
        }
    }
    $hdr = ($dt.Columns | ForEach-Object { $_.ColumnName.PadRight($widths[$_.ColumnName]) }) -join '  '
    $sep = ($dt.Columns | ForEach-Object { '-' * $widths[$_.ColumnName] }) -join '  '
    W ($pad + $hdr)
    W ($pad + $sep)
    foreach ($row in $rows) {
        $line = ($dt.Columns | ForEach-Object {
            $v = "$($row[$_.ColumnName])"
            if ($v.Length -gt $MaxColWidth) { $v = $v.Substring(0, $MaxColWidth - 3) + '...' }
            $v.PadRight($widths[$_.ColumnName])
        }) -join '  '
        W ($pad + $line)
    }
    if ($dt.Rows.Count -gt $MaxRows) {
        W ($pad + "... ($($dt.Rows.Count - $MaxRows) more rows truncated)")
    }
}

# ==============================================================================
WH 'SECTION 0: Locate SQL Server instances'
# ==============================================================================
$instances = @()
@('HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server',
  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server') | ForEach-Object {
    if (Test-Path $_) {
        $p = Get-ItemProperty $_ -Name InstalledInstances -EA SilentlyContinue
        if ($p -and $p.InstalledInstances) {
            foreach ($inst in $p.InstalledInstances) { $instances += $inst }
        }
    }
}
$instances = @($instances | Select-Object -Unique)
W "Registry instances: $($instances -join ', ')"

# Build connection string candidates
$candidates = @()
foreach ($i in $instances) {
    if ($i -eq 'MSSQLSERVER') { $candidates += 'localhost' }
    else { $candidates += "localhost\$i" }
}
if (-not $candidates) { $candidates = @('localhost','localhost\SQLEXPRESS','localhost\MSSQLSERVER') }

# Find first working connection
$srv = $null
foreach ($c in $candidates) {
    $r = Invoke-SQL -Server $c -Database 'master' -Query 'SELECT 1'
    if ($r -ne $null) { $srv = $c; break }
}
if (-not $srv) {
    W 'ERROR: Could not connect to any SQL instance with sa/FujiFilm_11111'
    $out | Set-Content 'C:\Temp\ControlSuite_DB.txt' -Encoding UTF8
    exit 1
}

WH "CONNECTED: $srv"
$verDt = Invoke-SQL -Server $srv -Database 'master' -Query 'SELECT @@SERVERNAME AS ServerName, @@VERSION AS Version, SERVERPROPERTY(''Edition'') AS Edition, SERVERPROPERTY(''ProductVersion'') AS ProductVersion'
Show-Table $verDt

# ==============================================================================
WH 'SECTION 1: All Databases'
# ==============================================================================
$dbQ = @"
SELECT d.name,
       d.state_desc                                             AS State,
       d.recovery_model_desc                                    AS RecoveryModel,
       d.compatibility_level                                    AS CompatLevel,
       CAST(ROUND(SUM(f.size) * 8.0 / 1024, 1) AS VARCHAR(20)) AS SizeMB,
       CONVERT(VARCHAR(20), d.create_date, 120)                 AS CreateDate
FROM   sys.databases d
LEFT   JOIN sys.master_files f ON d.database_id = f.database_id
GROUP  BY d.name, d.state_desc, d.recovery_model_desc, d.compatibility_level, d.create_date
ORDER  BY d.name
"@
$allDbsDt = Invoke-SQL -Server $srv -Database 'master' -Query $dbQ
Show-Table $allDbsDt

# Determine which databases to inspect
$systemDbs = @('master','model','msdb','tempdb')
$targetDbs = @()
if ($allDbsDt -and $allDbsDt[0].Rows.Count -gt 0) {
    foreach ($row in $allDbsDt[0].Rows) {
        if ($row.name -notin $systemDbs -and $row.State -eq 'ONLINE') {
            $targetDbs += $row.name
        }
    }
}
W ''
W "Target databases for inspection: $($targetDbs -join ', ')"

# ==============================================================================
WH 'SECTION 2: Server-level Logins'
# ==============================================================================
$loginQ = "SELECT name, type_desc, is_disabled, CONVERT(VARCHAR(20),create_date,120) AS Created, default_database_name FROM sys.server_principals WHERE type IN ('S','U','G') ORDER BY type_desc, name"
Show-Table (Invoke-SQL -Server $srv -Database 'master' -Query $loginQ)

# ==============================================================================
WH 'SECTION 3: SQL Agent Jobs'
# ==============================================================================
$jobQ = @"
SELECT j.name, j.enabled, j.description,
       ISNULL(CAST(jh.last_run_date AS VARCHAR),'never') AS LastRunDate,
       CASE jh.last_run_outcome WHEN 1 THEN 'Success' WHEN 0 THEN 'Fail' ELSE '?' END AS LastResult
FROM   msdb.dbo.sysjobs j
LEFT   JOIN msdb.dbo.sysjobservers jh ON j.job_id = jh.job_id
ORDER  BY j.name
"@
Show-Table (Invoke-SQL -Server $srv -Database 'master' -Query $jobQ)

# ==============================================================================
# Per-database inspection
# ==============================================================================
# Config-like table names to look for (Equitrac / ControlSuite naming conventions)
$configTables = @(
    'tblConfiguration','tblConfig','Configuration','Settings',
    'tblSettings','tblSystemSettings','SystemSettings',
    'tblServer','tblServers','Servers',
    'tblParameter','Parameters','tblParameters',
    'tblOption','Options','tblOptions',
    'tblSite','Sites','tblDomain','Domains',
    'tblLicense','License','tblLicenses',
    'tblUsers','tblUser','Users',
    'tblUserGroup','UserGroups','tblAccount','Accounts',
    'tblPrinter','Printers','tblQueue','Queues',
    'tblDevice','Devices','tblPort','tblRule','Rules',
    'tblPolicy','Policies','tblAudit','tblTracking','Tracking',
    'tblJob','Jobs','tblCost','Costs','tblRate','Rates',
    'tblBillingCode','BillingCodes','tblDepartment','Departments',
    'tblCasServer','CasServer','tblWebServer','WebServer',
    'tblAuthProvider','tblDirectory','tblLdap','tblActiveDirectory',
    'tblEmail','tblSmtp','tblNotification','tblAlert'
)

foreach ($db in $targetDbs) {
    WH "DATABASE: $db"

    # Table list with row counts
    WS 'Tables (by row count desc)'
    $tblQ = @"
SELECT   t.name                                                          AS TableName,
         SUM(p.rows)                                                     AS Rows,
         CAST(ROUND(SUM(a.total_pages)*8.0/1024,2) AS VARCHAR(20))+'MB' AS Size
FROM     sys.tables t
JOIN     sys.indexes     i ON t.object_id = i.object_id AND i.index_id IN (0,1)
JOIN     sys.partitions  p ON i.object_id = p.object_id AND i.index_id = p.index_id
JOIN     sys.allocation_units a ON p.partition_id = a.container_id
GROUP BY t.name
ORDER BY SUM(p.rows) DESC
"@
    Show-Table (Invoke-SQL -Server $srv -Database $db -Query $tblQ)

    # Stored procedures
    WS 'Stored Procedures'
    $spDt = Invoke-SQL -Server $srv -Database $db -Query "SELECT name, CONVERT(VARCHAR(20),create_date,120) AS Created FROM sys.procedures WHERE is_ms_shipped=0 ORDER BY name"
    if ($spDt -and $spDt[0].Rows.Count -gt 0) {
        $spNames = ($spDt[0].Rows | ForEach-Object { $_.name }) -join ', '
        W "  $($spDt[0].Rows.Count) procs: $spNames"
    } else { W '  (none)' }

    # Check which config-like tables exist
    $inList = "'" + ($configTables -join "','") + "'"
    $existDt = Invoke-SQL -Server $srv -Database $db -Query "SELECT name FROM sys.tables WHERE name IN ($inList) ORDER BY name"
    $existing = if ($existDt -and $existDt[0].Rows.Count -gt 0) { $existDt[0].Rows | ForEach-Object { $_.name } } else { @() }

    foreach ($tbl in $existing) {
        WS "Config table: $tbl"
        $colDt = Invoke-SQL -Server $srv -Database $db -Query "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='$tbl' ORDER BY ORDINAL_POSITION"
        if ($colDt -and $colDt[0].Rows.Count -gt 0) {
            $colSummary = ($colDt[0].Rows | ForEach-Object { "$($_.COLUMN_NAME)($($_.DATA_TYPE))" }) -join ', '
            W "  Columns: $colSummary"
            # Build safe column list (exclude binary/blob types)
            $blobTypes = @('image','varbinary','binary','timestamp','hierarchyid','geometry','geography','xml')
            $safeCols = ($colDt[0].Rows | Where-Object { $_.DATA_TYPE -notin $blobTypes } | ForEach-Object { "[$($_.COLUMN_NAME)]" }) -join ','
            if ($safeCols) {
                Show-Table (Invoke-SQL -Server $srv -Database $db -Query "SELECT TOP 200 $safeCols FROM [$tbl] WITH (NOLOCK)" -TimeoutSec 20) -Indent 4
            }
        }
    }

    # Find any name/value style tables not already covered
    WS 'Additional name-value config tables (auto-detected)'
    $nvQ = @"
SELECT DISTINCT t.name AS tbl
FROM   sys.tables t
JOIN   sys.columns cn ON t.object_id=cn.object_id AND cn.name IN ('Name','Key','Setting','Parameter','ConfigName','SettingName','PropertyName')
JOIN   sys.columns cv ON t.object_id=cv.object_id AND cv.name IN ('Value','StringValue','IntValue','Data','ConfigValue','SettingValue','PropertyValue')
WHERE  t.name NOT IN ($inList)
ORDER  BY t.name
"@
    $nvDt = Invoke-SQL -Server $srv -Database $db -Query $nvQ
    if ($nvDt -and $nvDt[0].Rows.Count -gt 0) {
        foreach ($row in $nvDt[0].Rows) {
            $tbl = $row.tbl
            WS "  name-value table: $tbl"
            $colDt2 = Invoke-SQL -Server $srv -Database $db -Query "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='$tbl' ORDER BY ORDINAL_POSITION"
            if ($colDt2 -and $colDt2[0].Rows.Count -gt 0) {
                $blobTypes = @('image','varbinary','binary','timestamp','hierarchyid','geometry','geography','xml')
                $sc = ($colDt2[0].Rows | Where-Object { $_.DATA_TYPE -notin $blobTypes } | ForEach-Object { "[$($_.COLUMN_NAME)]" }) -join ','
                if ($sc) {
                    Show-Table (Invoke-SQL -Server $srv -Database $db -Query "SELECT TOP 200 $sc FROM [$tbl] WITH (NOLOCK)" -TimeoutSec 20) -Indent 4
                }
            }
        }
    } else { W '  (none auto-detected)' }

    # Database-level users
    WS 'Database Users'
    $dbuQ = "SELECT name, type_desc, CONVERT(VARCHAR(20),create_date,120) AS Created, default_schema_name FROM sys.database_principals WHERE type NOT IN ('R') AND name NOT LIKE '##%' ORDER BY name"
    Show-Table (Invoke-SQL -Server $srv -Database $db -Query $dbuQ)
}

# ==============================================================================
WH 'SECTION 5: msdb -- SQL Agent history for brand-related jobs'
# ==============================================================================
$histQ = @"
SELECT TOP 100
    j.name                                                          AS JobName,
    CONVERT(VARCHAR(20), CONVERT(DATETIME,
        CONVERT(VARCHAR, jh.run_date) + ' ' +
        STUFF(STUFF(RIGHT('000000'+CONVERT(VARCHAR,jh.run_time),6),5,0,':'),3,0,':')),120) AS RunTime,
    CASE jh.run_status WHEN 1 THEN 'Success' WHEN 0 THEN 'Fail' WHEN 3 THEN 'Cancel' ELSE 'Other' END AS Status,
    jh.message
FROM  msdb.dbo.sysjobhistory jh
JOIN  msdb.dbo.sysjobs       j  ON jh.job_id = j.job_id
WHERE jh.step_id = 0
ORDER BY jh.run_date DESC, jh.run_time DESC
"@
Show-Table (Invoke-SQL -Server $srv -Database 'master' -Query $histQ)

# ==============================================================================
$outPath = 'C:\Temp\ControlSuite_DB.txt'
$out | Set-Content -Path $outPath -Encoding UTF8
Write-Host "Done. $($out.Count) lines -> $outPath"
