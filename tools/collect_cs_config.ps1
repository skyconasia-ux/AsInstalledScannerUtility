# ControlSuite / Equitrac -- cas_* table deep-dive
# Credentials: sa / FujiFilm_11111  (local instance, non-production)
# Output: C:\Temp\ControlSuite_Config.txt
# NOTE: ASCII-only script (PS 5.1 compatible)

$out = [System.Collections.Generic.List[string]]::new()
$ts  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

function W  { param([string]$line='') $out.Add($line) }
function WH { param([string]$h) W ''; W ('=' * 72); W "  $h"; W ('=' * 72) }
function WS { param([string]$s) W ''; W "  --- $s ---" }

W "ControlSuite / Equitrac -- cas_* Config Deep-Dive"
W "Generated : $ts"
W "Host      : $env:COMPUTERNAME"

function Invoke-SQL {
    param(
        [string]$Server,
        [string]$Database = 'master',
        [string]$Query,
        [string]$User = 'sa',
        [string]$Pass = 'FujiFilm_11111',
        [int]$TimeoutSec = 60
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
        return ,$dt
    } catch {
        return $null
    }
}

function Show-Table {
    param($dt, [int]$Indent = 2, [int]$MaxRows = 1000, [int]$MaxColWidth = 120)
    if ($dt -is [array]) { $dt = $dt[0] }
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

# Connect
$srv = 'localhost\SQLEXPRESS'
$db  = 'eqcas'

$verDt = Invoke-SQL -Server $srv -Database 'master' -Query 'SELECT @@SERVERNAME AS ServerName, @@VERSION AS Version'
if (-not $verDt) {
    W "ERROR: Cannot connect to $srv"
    $out | Set-Content 'C:\Temp\ControlSuite_Config.txt' -Encoding UTF8
    exit 1
}
WH "CONNECTED: $srv / $db"
Show-Table $verDt

# ==============================================================================
WH 'SECTION 1: All cas_* and cat_* tables with row counts'
# ==============================================================================
$allTablesQ = @"
SELECT   t.name AS TableName,
         SUM(p.rows) AS Rows,
         CAST(ROUND(SUM(a.total_pages)*8.0/1024,2) AS VARCHAR(20))+'MB' AS Size
FROM     sys.tables t
JOIN     sys.indexes i ON t.object_id=i.object_id AND i.index_id IN (0,1)
JOIN     sys.partitions p ON i.object_id=p.object_id AND i.index_id=p.index_id
JOIN     sys.allocation_units a ON p.partition_id=a.container_id
WHERE    t.name LIKE 'cas[_]%' OR t.name LIKE 'cat[_]%'
GROUP BY t.name
ORDER BY SUM(p.rows) DESC
"@
$allTblDt = Invoke-SQL -Server $srv -Database $db -Query $allTablesQ
Show-Table $allTblDt

# ==============================================================================
WH 'SECTION 2: cas_config -- full dump (all rows, all non-blob columns)'
# ==============================================================================
$colDt = Invoke-SQL -Server $srv -Database $db -Query "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='cas_config' ORDER BY ORDINAL_POSITION"
if ($colDt -and $colDt[0].Rows.Count -gt 0) {
    $blobTypes = @('image','varbinary','binary','timestamp','hierarchyid','geometry','geography')
    $safeCols = ($colDt[0].Rows | Where-Object { $_.DATA_TYPE -notin $blobTypes } | ForEach-Object { "[$($_.COLUMN_NAME)]" }) -join ','
    $colSummary = ($colDt[0].Rows | ForEach-Object { "$($_.COLUMN_NAME)($($_.DATA_TYPE))" }) -join ', '
    W "Columns: $colSummary"
    W ''
    Show-Table (Invoke-SQL -Server $srv -Database $db -Query "SELECT $safeCols FROM [cas_config] WITH (NOLOCK) ORDER BY 1" -TimeoutSec 60) -MaxRows 2000
} else {
    W '  cas_config not found or no columns'
}

# ==============================================================================
WH 'SECTION 3: Key operational tables -- sample 200 rows each'
# ==============================================================================
$keyTables = @(
    'cas_server',
    'cas_site',
    'cas_domain',
    'cas_device',
    'cas_printer',
    'cas_queue',
    'cas_user',
    'cas_account',
    'cas_department',
    'cas_usergroup',
    'cas_group',
    'cas_license',
    'cas_pricelist',
    'cas_pricelist_attributes',
    'cas_dms_item',
    'cas_dms_config',
    'cas_rule',
    'cas_policy',
    'cas_authprovider',
    'cas_auth_provider',
    'cas_ldap',
    'cas_directory',
    'cas_smtp',
    'cas_email',
    'cas_notification',
    'cas_alert',
    'cas_webserver',
    'cas_casserver',
    'cas_billing',
    'cas_billingcode',
    'cas_cost',
    'cas_rate',
    'cas_tracking',
    'cas_audit',
    'cas_job',
    'cat_type',
    'cat_category',
    'cat_status'
)

foreach ($tbl in $keyTables) {
    $chkDt = Invoke-SQL -Server $srv -Database $db -Query "SELECT COUNT(*) AS cnt FROM sys.tables WHERE name='$tbl'"
    $exists = ($chkDt -and $chkDt[0].Rows.Count -gt 0 -and $chkDt[0].Rows[0]['cnt'] -gt 0)
    if (-not $exists) { continue }

    WS "Table: $tbl"
    $colDt2 = Invoke-SQL -Server $srv -Database $db -Query "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='$tbl' ORDER BY ORDINAL_POSITION"
    if ($colDt2 -and $colDt2[0].Rows.Count -gt 0) {
        $blobTypes = @('image','varbinary','binary','timestamp','hierarchyid','geometry','geography')
        $sc = ($colDt2[0].Rows | Where-Object { $_.DATA_TYPE -notin $blobTypes } | ForEach-Object { "[$($_.COLUMN_NAME)]" }) -join ','
        if ($sc) {
            Show-Table (Invoke-SQL -Server $srv -Database $db -Query "SELECT TOP 200 $sc FROM [$tbl] WITH (NOLOCK)" -TimeoutSec 30) -MaxRows 200
        }
    }
}

# ==============================================================================
WH 'SECTION 4: Any remaining cas_* tables with rows not yet covered'
# ==============================================================================
$coveredTables = $keyTables + @('cas_config')
$inList = "'" + ($coveredTables -join "','") + "'"

$remainQ = @"
SELECT   t.name AS TableName, SUM(p.rows) AS Rows
FROM     sys.tables t
JOIN     sys.indexes i ON t.object_id=i.object_id AND i.index_id IN (0,1)
JOIN     sys.partitions p ON i.object_id=p.object_id AND i.index_id=p.index_id
WHERE    (t.name LIKE 'cas[_]%' OR t.name LIKE 'cat[_]%')
  AND    t.name NOT IN ($inList)
GROUP BY t.name
HAVING   SUM(p.rows) > 0
ORDER BY SUM(p.rows) DESC
"@
$remainDt = Invoke-SQL -Server $srv -Database $db -Query $remainQ
if ($remainDt -and $remainDt[0].Rows.Count -gt 0) {
    W "Uncovered tables with rows:"
    Show-Table $remainDt
    foreach ($row in $remainDt[0].Rows) {
        $tbl = $row.TableName
        WS "Table: $tbl"
        $colDt3 = Invoke-SQL -Server $srv -Database $db -Query "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='$tbl' ORDER BY ORDINAL_POSITION"
        if ($colDt3 -and $colDt3[0].Rows.Count -gt 0) {
            $blobTypes = @('image','varbinary','binary','timestamp','hierarchyid','geometry','geography')
            $sc = ($colDt3[0].Rows | Where-Object { $_.DATA_TYPE -notin $blobTypes } | ForEach-Object { "[$($_.COLUMN_NAME)]" }) -join ','
            if ($sc) {
                Show-Table (Invoke-SQL -Server $srv -Database $db -Query "SELECT TOP 200 $sc FROM [$tbl] WITH (NOLOCK)" -TimeoutSec 30) -MaxRows 200
            }
        }
    }
} else {
    W '  (all cas_* / cat_* tables covered)'
}

# ==============================================================================
$outPath = 'C:\Temp\ControlSuite_Config.txt'
$out | Set-Content -Path $outPath -Encoding UTF8
Write-Host "Done. $($out.Count) lines -> $outPath"
