# EQ-Snapshot.ps1
# Captures a full snapshot of Equitrac/ControlSuite config state:
#   - All eqcas SQL tables (row counts)
#   - Registry: Equitrac, Kofax, Nuance, Tungsten, FLEXlm/Flexera, Windows Services
#   - File tree timestamps: Program Files\Kofax, ProgramData, FlexNet,
#     user/system profile AppData paths, Temp directories
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File EQ-Snapshot.ps1 -Label BEFORE
#   powershell -NoProfile -ExecutionPolicy Bypass -File EQ-Snapshot.ps1 -Label AFTER
#
# Output: C:\Temp\EQ_Snapshots\<Label>\
#
# NOTE: ASCII-only script (PS 5.1 compatible)

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('BEFORE','AFTER')]
    [string]$Label
)

$ErrorActionPreference = 'SilentlyContinue'

$SqlServer = '.\SQLExpress'
$SqlDb     = 'eqcas'
$SqlUser   = 'sa'
$SqlPass   = 'FujiFilm_11111'

$OutDir = "C:\Temp\EQ_Snapshots\$Label"
if (Test-Path $OutDir) { Remove-Item $OutDir -Recurse -Force }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
New-Item -ItemType Directory -Path "$OutDir\sql" -Force | Out-Null

$log = [System.Collections.Generic.List[string]]::new()

function Log { param([string]$m) $log.Add($m); Write-Host $m }

# Single persistent connection - avoids pool exhaustion across 110+ table queries
$script:SqlCon = $null
function Get-SqlCon {
    if ($script:SqlCon -eq $null -or $script:SqlCon.State -ne 'Open') {
        $cs = "Server=$SqlServer;Database=$SqlDb;User Id=$SqlUser;Password=$SqlPass;Connection Timeout=10;TrustServerCertificate=True;"
        $script:SqlCon = New-Object System.Data.SqlClient.SqlConnection $cs
        $script:SqlCon.Open()
    }
    return $script:SqlCon
}

# Returns a DataTable. Uses DataTable.Load() which tolerates empty column names.
# Columns with empty names get renamed col_0, col_1, etc. after loading.
function Invoke-SQL {
    param([string]$Query, [int]$Timeout = 60)
    try {
        $con             = Get-SqlCon
        $cmd             = $con.CreateCommand()
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = $Timeout
        $reader = $cmd.ExecuteReader()
        $dt     = New-Object System.Data.DataTable
        $dt.Load($reader)

        # Rename any empty column names so TSV headers are usable
        for ($i = 0; $i -lt $dt.Columns.Count; $i++) {
            if ([string]::IsNullOrWhiteSpace($dt.Columns[$i].ColumnName)) {
                $dt.Columns[$i].ColumnName = "col_$i"
            }
        }
        return $dt
    } catch {
        Log "  SQL ERROR: $_"
        try { if ($script:SqlCon) { $script:SqlCon.Close(); $script:SqlCon = $null } } catch {}
        return $null
    }
}

# --- Registry recursive dump ---
function Dump-RegistryKey {
    param([string]$Path, [System.Collections.Generic.List[string]]$Lines)
    if (-not (Test-Path $Path)) { return }
    $Lines.Add("[[$Path]]")
    try {
        $item = Get-Item $Path -ErrorAction Stop
        foreach ($vn in $item.GetValueNames()) {
            $vd = $item.GetValue($vn)
            $vt = $item.GetValueKind($vn)
            $Lines.Add("  $vn = $vd  ($vt)")
        }
    } catch {}
    try {
        Get-ChildItem $Path -ErrorAction Stop | ForEach-Object {
            Dump-RegistryKey -Path $_.PSPath -Lines $Lines
        }
    } catch {}
}

# Tables to skip full content export (transaction/log/cache data)
$SkipFullExport = @(
    'cas_trx_acc_ext','cas_trx_doc_ext','cas_trx_docdetail_ext','cas_trx_escrow_ext',
    'cas_trx_exception_rule','cas_trx_exported','cas_trx_fax_ext','cas_trx_orphan_ext',
    'cas_trx_pcb_ext','cas_trx_print_ext','cas_trx_scan_ext',
    'cas_uplink_trx_acc_ext','cas_uplink_trx_doc_ext','cas_uplink_trx_docdetail_ext',
    'cas_uplink_trx_fax_ext','cas_uplink_trx_print_ext','cas_uplink_trx_scan_ext','cas_uplink_trx_sum',
    'cas_fas_trx_cache','cas_fas_trx_cache_id','cas_fas_acc_cache',
    'cas_audit_log','cas_sdr_history','cas_printer_status_history',
    'cas_dashboard_data','cas_notif_clients_data','cas_mru_bc_by_user',
    'cas_spe_jobs','cat_transaction','cat_trxid','cat_trxval','cat_trxxml',
    'cas_report_bitmap','cas_update_sequence_number'
)

# ============================================================
# 1. SQL - Row counts for all tables
# ============================================================
Log "[$Label] SQL: reading row counts..."

$countQuery = @"
SELECT
    t.TABLE_NAME,
    p.row_count
FROM INFORMATION_SCHEMA.TABLES t
LEFT JOIN (
    SELECT OBJECT_NAME(object_id) AS tbl, SUM(row_count) AS row_count
    FROM sys.dm_db_partition_stats
    WHERE index_id IN (0,1)
    GROUP BY object_id
) p ON p.tbl = t.TABLE_NAME
WHERE t.TABLE_TYPE = 'BASE TABLE'
ORDER BY t.TABLE_NAME
"@

$counts = Invoke-SQL -Query $countQuery
$countLines = [System.Collections.Generic.List[string]]::new()
$countLines.Add("TABLE_NAME`tROW_COUNT")

if ($counts) {
    foreach ($row in $counts.Rows) {
        $countLines.Add("$($row[0])`t$($row[1])")
    }
}
[System.IO.File]::WriteAllLines("$OutDir\sql_counts.txt", $countLines, [System.Text.Encoding]::ASCII)
Log "  -> $($counts.Rows.Count) tables recorded"

# Row counts captured above via sys.dm_db_partition_stats (one query, no per-table loop).

# ============================================================
# 2. SQL - BCP content dump of config tables
# ============================================================
# Row counts miss UPDATE operations (pricing edits, workflow renames, etc.).
# BCP dumps give us full row content so EQ-Diff can detect any changed value.
# These tables are config-only (not transactions/logs) and small enough to dump.
# ============================================================
Log "[$Label] SQL: BCP content dump of config tables..."

$BcpTables = @(
    'cas_config',           # master key-value config store
    'cas_scan_alias',       # workflow scan destinations (name, scope, destination, active)
    'cat_pricelist',        # pricing lists (B&W, colour page prices)
    'cas_pullgroups',       # pull-print / FollowYou groups
    'cas_workflow_folders', # workflow folder assignments
    'cas_user_ext',         # user accounts and properties
    'cat_validation',       # billing/PIN validation rules
    'cas_prq_device_ext',   # per-device pull-print settings
    'cas_installedsoftware' # installed component versions
)

$BcpDir = "$OutDir\sql_bcp"
New-Item -ItemType Directory -Path $BcpDir -Force | Out-Null

foreach ($tbl in $BcpTables) {
    $outFile = "$BcpDir\$tbl.txt"
    $bcpArgs = @("$SqlDb..$tbl", 'out', $outFile,
                 '-S', $SqlServer, '-U', $SqlUser, '-P', $SqlPass,
                 '-c', '-t', '|')
    & bcp @bcpArgs 2>$null | Out-Null
    $rows = if (Test-Path $outFile) { (Get-Content $outFile -ErrorAction SilentlyContinue).Count } else { 0 }
    Log "  $tbl : $rows rows"
}

# ============================================================
# 3. Registry snapshot
# ============================================================
Log "[$Label] Registry: dumping Equitrac/Kofax/Nuance paths..."

$regRoots = @(
    # Software keys - vendor namespaces
    'HKLM:\SOFTWARE\Equitrac',
    'HKLM:\SOFTWARE\Kofax',
    'HKLM:\SOFTWARE\Nuance',
    'HKLM:\SOFTWARE\Tungsten Automation',
    'HKLM:\SOFTWARE\WOW6432Node\Equitrac',
    'HKLM:\SOFTWARE\WOW6432Node\Kofax',
    'HKLM:\SOFTWARE\WOW6432Node\Nuance',
    'HKLM:\SOFTWARE\WOW6432Node\Tungsten Automation',

    # FlexNet / FLEXlm licensing
    'HKLM:\SOFTWARE\FLEXlm License Manager',
    'HKLM:\SOFTWARE\WOW6432Node\FLEXlm License Manager',
    'HKLM:\SOFTWARE\Flexera Software',
    'HKLM:\SOFTWARE\WOW6432Node\Flexera Software',

    # Windows Services - Equitrac / Kofax / Nuance
    'HKLM:\SYSTEM\CurrentControlSet\Services\Equitrac',
    'HKLM:\SYSTEM\CurrentControlSet\Services\EQCASSrv',
    'HKLM:\SYSTEM\CurrentControlSet\Services\EQDCESrv',
    'HKLM:\SYSTEM\CurrentControlSet\Services\EQDMESrv',
    'HKLM:\SYSTEM\CurrentControlSet\Services\EQDRESrv',
    'HKLM:\SYSTEM\CurrentControlSet\Services\EQDWSSrv',
    'HKLM:\SYSTEM\CurrentControlSet\Services\Kofax',
    'HKLM:\SYSTEM\CurrentControlSet\Services\Nuance',
    'HKLM:\SYSTEM\CurrentControlSet\Services\flexnetls',
    'HKLM:\SYSTEM\CurrentControlSet\Services\NDISecurityService'
)

$regLines = [System.Collections.Generic.List[string]]::new()
$regLines.Add("# Registry snapshot - $Label - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")

foreach ($root in $regRoots) {
    if (Test-Path $root) {
        $regLines.Add("")
        $regLines.Add("# ROOT: $root")
        Dump-RegistryKey -Path $root -Lines $regLines
    }
}
[System.IO.File]::WriteAllLines("$OutDir\registry.txt", $regLines, [System.Text.Encoding]::ASCII)
Log "  -> $($regLines.Count) registry lines recorded"

# ============================================================
# 3b. SQLite snapshot - EQVar config databases
# ============================================================
# All Equitrac/ControlSuite configuration lives in EQVar tables inside
# SQLite .db3 files under the Local System profile. We copy + query each
# database so EQ-Diff can do row-level diffs, not just timestamp diffs.
# ============================================================
Log "[$Label] SQLite: dumping EQVar config databases..."

$DbBase = 'C:\Windows\System32\config\systemprofile\AppData\Local\Equitrac\Equitrac Platform Component'
$SqliteDbs = @(
    @{ Name='DCE_config'; Path="$DbBase\EQDCESrv\Cache\DCE_config.db3" },
    @{ Name='DREEQVar';   Path="$DbBase\EQDRESrv\EQSpool\DREEQVar.db3"  },
    @{ Name='EQDMECache'; Path="$DbBase\EQDMESrv\EQDMECache.db3"         }
)

$SqliteTmp = "$OutDir\sqlite_tmp"
New-Item -ItemType Directory -Path $SqliteTmp -Force | Out-Null

foreach ($db in $SqliteDbs) {
    if (-not (Test-Path $db.Path)) {
        Log "  [skip] $($db.Name) - not found"
        continue
    }
    # Copy db + WAL + SHM so sqlite3 reads the full current state
    $tmp = "$SqliteTmp\$($db.Name).db3"
    Copy-Item $db.Path $tmp -Force
    if (Test-Path "$($db.Path)-wal") { Copy-Item "$($db.Path)-wal" "$tmp-wal" -Force }
    if (Test-Path "$($db.Path)-shm") { Copy-Item "$($db.Path)-shm" "$tmp-shm" -Force }

    $outTsv = "$OutDir\sqlite_$($db.Name).tsv"
    # Dump EQVar table as TSV. sqlite3 exits non-zero on error.
    $q = "SELECT System,SubSystem,Class,Name,Value FROM EQVar ORDER BY SubSystem,Class,Name;"
    & sqlite3 -separator "`t" $tmp $q 2>$null | Out-File -FilePath $outTsv -Encoding ASCII

    $lineCount = if (Test-Path $outTsv) { (Get-Content $outTsv).Count } else { 0 }
    Log "  $($db.Name): $lineCount rows -> sqlite_$($db.Name).tsv"
}

# Clean up temp copies
Remove-Item $SqliteTmp -Recurse -Force -ErrorAction SilentlyContinue

# ============================================================
# 4. File system snapshot (timestamps + sizes)
# ============================================================
Log "[$Label] Files: scanning config directories..."

$WatchDirs = @(
    # Main installation tree (binaries + config files under Program Files)
    'C:\Program Files\Kofax',

    # ProgramData - primary runtime config, logs, data dirs
    'C:\ProgramData\Equitrac',
    'C:\ProgramData\flexnetsas',
    'C:\ProgramData\Kofax',
    'C:\ProgramData\Nuance',

    # FlexNet license server data (runs as NetworkService)
    'C:\Windows\ServiceProfiles\NetworkService\flexnetls',

    # User profile paths - if CAS service runs as a named Windows account
    'C:\Users\Administrator\AppData\Local\Equitrac',
    'C:\Users\Administrator\AppData\Local\Nuance',

    # System profile paths - if CAS service runs as Local System
    'C:\Windows\System32\config\systemprofile\AppData\Local\Equitrac',
    'C:\Windows\System32\config\systemprofile\AppData\Local\Nuance',

    # Temp directories - catch installer/upgrade artifacts
    'C:\Temp',
    'C:\Windows\Temp',
    'C:\Users\Administrator\AppData\Local\Temp'
)

$fileLines = [System.Collections.Generic.List[string]]::new()
$fileLines.Add("FullName`tLastWriteTime`tLength")

foreach ($dir in $WatchDirs) {
    if (-not (Test-Path $dir)) { continue }
    Get-ChildItem -Path $dir -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        $ts  = $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
        $fileLines.Add("$($_.FullName)`t$ts`t$($_.Length)")
    }
}
[System.IO.File]::WriteAllLines("$OutDir\files.txt", $fileLines, [System.Text.Encoding]::ASCII)
Log "  -> $($fileLines.Count - 1) files recorded"

# ============================================================
# 5. Write metadata
# ============================================================
$meta = @(
    "Label     : $Label",
    "Timestamp : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    "Host      : $env:COMPUTERNAME",
    "SQL Server: $SqlServer",
    "Database  : $SqlDb"
)
[System.IO.File]::WriteAllLines("$OutDir\info.txt", $meta, [System.Text.Encoding]::ASCII)
[System.IO.File]::WriteAllLines("$OutDir\snapshot.log", $log, [System.Text.Encoding]::ASCII)

if ($script:SqlCon -and $script:SqlCon.State -eq 'Open') { $script:SqlCon.Close() }

Log ""
Log "[$Label] Snapshot complete -> $OutDir"
