# EQ-Diff.ps1
# Compares BEFORE and AFTER snapshots produced by EQ-Snapshot.ps1.
# Outputs a diff report showing exactly what changed in SQL, registry, and files.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File EQ-Diff.ps1
#
# Reads:  C:\Temp\EQ_Snapshots\BEFORE\
#         C:\Temp\EQ_Snapshots\AFTER\
# Output: C:\Temp\EQ_Snapshots\EQ_Diff_Report.txt  (and printed to console)
#
# Behaviour:
#   - Empty tables and unchanged tables are silently skipped - no noise.
#   - Changed SQL tables are BCP-dumped so you see actual new/current content.
#   - Known cache/session file paths are filtered out of the file diff.
#
# NOTE: ASCII-only script (PS 5.1 compatible)

$ErrorActionPreference = 'SilentlyContinue'

$BeforeDir  = 'C:\Temp\EQ_Snapshots\BEFORE'
$AfterDir   = 'C:\Temp\EQ_Snapshots\AFTER'
$ReportFile = 'C:\Temp\EQ_Snapshots\EQ_Diff_Report.txt'

$SqlServer = '.\SQLExpress'
$SqlDb     = 'eqcas'
$SqlUser   = 'sa'
$SqlPass   = 'FujiFilm_11111'

$out = [System.Collections.Generic.List[string]]::new()

function W  { param([string]$l='') $out.Add($l); Write-Host $l }
function WH { param([string]$h) W ''; W ('=' * 72); W "  $h"; W ('=' * 72) }
function WS { param([string]$s) W ''; W "  --- $s ---" }

if (-not (Test-Path $BeforeDir)) { Write-Host "ERROR: BEFORE snapshot not found at $BeforeDir"; exit 1 }
if (-not (Test-Path $AfterDir))  { Write-Host "ERROR: AFTER snapshot not found at $AfterDir";  exit 1 }

$beforeInfo = Get-Content "$BeforeDir\info.txt" -ErrorAction SilentlyContinue
$afterInfo  = Get-Content "$AfterDir\info.txt"  -ErrorAction SilentlyContinue

W "EQ Diff Report"
W "Generated  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
W "BEFORE     : $($beforeInfo | Select-String 'Timestamp' | ForEach-Object { $_ -replace '.*: ',''})"
W "AFTER      : $($afterInfo  | Select-String 'Timestamp' | ForEach-Object { $_ -replace '.*: ',''})"

# ============================================================
# File noise filter - paths to suppress from file diff output.
# These are known runtime/session/cache paths that change on
# every run and carry no configuration signal.
# ============================================================
$FileNoisePatterns = @(
    'EQ_Snapshots',           # snapshot tool's own output files
    'EQ_BcpDump',             # bcp temp exports
    'EQ_DbQuery',             # sqlite temp copies
    'ClientCommunicationsCache',
    'aa_token',
    'finder_cache',
    '[.]log$',
    '[.]ldf$',
    '[.]mdf$',
    'NDISecurity\\CS\\cache',
    'ssds',
    'ssds_seed_nodes'
)

function Is-NoisyFile {
    param([string]$path)
    foreach ($pat in $FileNoisePatterns) {
        if ($path -match $pat) { return $true }
    }
    return $false
}

# ============================================================
# 1. SQL - Row count changes + BCP content dump for changed tables
# ============================================================
WH "SQL TABLE CHANGES"

$beforeCnt = @{}
$afterCnt  = @{}

if (Test-Path "$BeforeDir\sql_counts.txt") {
    Import-Csv "$BeforeDir\sql_counts.txt" -Delimiter "`t" | ForEach-Object {
        if ($_.TABLE_NAME -ne '') { $beforeCnt[$_.TABLE_NAME] = $_.ROW_COUNT }
    }
}
if (Test-Path "$AfterDir\sql_counts.txt") {
    Import-Csv "$AfterDir\sql_counts.txt" -Delimiter "`t" | ForEach-Object {
        if ($_.TABLE_NAME -ne '') { $afterCnt[$_.TABLE_NAME] = $_.ROW_COUNT }
    }
}

$allTables   = ($beforeCnt.Keys + $afterCnt.Keys) | Sort-Object -Unique
$changedTbls = [System.Collections.Generic.List[string]]::new()

foreach ($tbl in $allTables) {
    $b = if ($beforeCnt.ContainsKey($tbl)) { [long]$beforeCnt[$tbl] } else { 0 }
    $a = if ($afterCnt.ContainsKey($tbl))  { [long]$afterCnt[$tbl]  } else { 0 }
    # Skip if both sides are zero - table was empty before and after
    if ($b -eq 0 -and $a -eq 0) { continue }
    $delta = $a - $b
    if ($delta -ne 0) {
        $sign = if ($delta -gt 0) { "+$delta" } else { "$delta" }
        W "  CHANGED  $tbl  ($b -> $a rows, $sign)"
        $changedTbls.Add($tbl)
    }
}

if ($changedTbls.Count -eq 0) {
    W "  No SQL table row-count changes detected."
} else {
    W ""
    W "  $($changedTbls.Count) table(s) changed."

    # ---- BCP dump current content of each changed table ----
    WS "Current content of changed tables (via BCP)"

    $bcpTmpDir = 'C:\Temp\EQ_BcpDump'
    if (-not (Test-Path $bcpTmpDir)) { New-Item -ItemType Directory -Path $bcpTmpDir | Out-Null }

    foreach ($tbl in $changedTbls) {
        $bcpFile = "$bcpTmpDir\$tbl.txt"
        # Remove stale file if present
        if (Test-Path $bcpFile) { Remove-Item $bcpFile -Force }

        $bcpArgs = @("$SqlDb..$tbl", 'out', $bcpFile,
                     '-S', $SqlServer, '-U', $SqlUser, '-P', $SqlPass,
                     '-c', '-t', '|')
        & bcp @bcpArgs 2>$null | Out-Null

        W ""
        W "  [TABLE: $tbl]"

        if (-not (Test-Path $bcpFile)) {
            W "    (BCP export failed or table unreadable)"
            continue
        }

        $rows = [System.IO.File]::ReadAllLines($bcpFile)
        if ($rows.Count -eq 0) {
            W "    (no rows exported)"
            continue
        }

        W "    Rows exported: $($rows.Count)"
        # Print all rows (cap at 200 to avoid runaway output)
        $limit = [Math]::Min($rows.Count, 200)
        for ($i = 0; $i -lt $limit; $i++) {
            # Truncate very long lines (e.g. XML blobs) at 300 chars
            $line = $rows[$i]
            if ($line.Length -gt 300) { $line = $line.Substring(0, 300) + '...[truncated]' }
            W "    $line"
        }
        if ($rows.Count -gt $limit) {
            W "    ...$($rows.Count - $limit) more rows (see $bcpFile for full export)"
        }
    }
}

# ============================================================
# 2. SQL - BCP content diff (catches UPDATE on existing rows)
# ============================================================
WH "SQL CONFIG TABLE CONTENT CHANGES"

$bcpTables = @(
    'cas_config','cas_scan_alias','cat_pricelist',
    'cas_pullgroups','cas_workflow_folders','cas_user_ext',
    'cat_validation','cas_prq_device_ext','cas_installedsoftware'
)

$anyBcpChange = $false

foreach ($tbl in $bcpTables) {
    $bFile = "$BeforeDir\sql_bcp\$tbl.txt"
    $aFile = "$AfterDir\sql_bcp\$tbl.txt"
    if (-not (Test-Path $bFile) -and -not (Test-Path $aFile)) { continue }

    $bLines = if (Test-Path $bFile) { [System.IO.File]::ReadAllLines($bFile) } else { @() }
    $aLines = if (Test-Path $aFile) { [System.IO.File]::ReadAllLines($aFile) } else { @() }

    # Index by first field (primary key / first column)
    $bMap = @{}; foreach ($l in $bLines) { $k = ($l -split '\|')[0]; $bMap[$k] = $l }
    $aMap = @{}; foreach ($l in $aLines) { $k = ($l -split '\|')[0]; $aMap[$k] = $l }

    $added   = $aMap.Keys | Where-Object { -not $bMap.ContainsKey($_) } | Sort-Object
    $removed = $bMap.Keys | Where-Object { -not $aMap.ContainsKey($_) } | Sort-Object
    $updated = $aMap.Keys | Where-Object { $bMap.ContainsKey($_) -and $bMap[$_] -ne $aMap[$_] } | Sort-Object

    if ($added.Count -eq 0 -and $removed.Count -eq 0 -and $updated.Count -eq 0) { continue }

    $anyBcpChange = $true
    WS "$tbl"

    foreach ($k in $updated) {
        W "  ~ row[$k]"
        $bv = $bMap[$k]; if ($bv.Length -gt 300) { $bv = $bv.Substring(0,300)+'...' }
        $av = $aMap[$k]; if ($av.Length -gt 300) { $av = $av.Substring(0,300)+'...' }
        W "      WAS: $bv"
        W "      NOW: $av"
    }
    foreach ($k in $added)   { $v = $aMap[$k]; if ($v.Length -gt 300){$v=$v.Substring(0,300)+'...'}; W "  + row[$k]: $v" }
    foreach ($k in $removed) { $v = $bMap[$k]; if ($v.Length -gt 300){$v=$v.Substring(0,300)+'...'}; W "  - row[$k]: $v" }
}

if (-not $anyBcpChange) { W "  No SQL config table content changes detected." }

# ============================================================
# 3. SQLite EQVar diff (primary config store)
# ============================================================
WH "SQLITE EQVAR CHANGES (primary config store)"

# Keys that change on every service cycle - suppress from diff output (noise)
$EQVarNoise = @(
    'cas||workflowfolderslastupdatetime',
    'cas||lastupdatetime',
    'dce||lastupdatetime',
    'cas||servicestarttimestamp',
    'dce||servicestarttimestamp'
)

$dbNames = @('DCE_config','DREEQVar','EQDMECache')
$anyDbChange = $false

foreach ($dbName in $dbNames) {
    $bFile = "$BeforeDir\sqlite_$dbName.tsv"
    $aFile = "$AfterDir\sqlite_$dbName.tsv"

    if (-not (Test-Path $bFile) -and -not (Test-Path $aFile)) { continue }

    # Load both snapshots into hashtables keyed by SubSystem|Class|Name
    function Load-EQVar {
        param([string]$Path)
        $map = @{}
        if (-not (Test-Path $Path)) { return $map }
        foreach ($line in [System.IO.File]::ReadAllLines($Path)) {
            $parts = $line -split "`t", 5
            if ($parts.Count -lt 4) { continue }
            $key = "$($parts[1])|$($parts[2])|$($parts[3])"
            $map[$key] = if ($parts.Count -ge 5) { $parts[4] } else { '' }
        }
        return $map
    }

    $bMap = Load-EQVar -Path $bFile
    $aMap = Load-EQVar -Path $aFile

    $allKeys = ($bMap.Keys + $aMap.Keys) | Sort-Object -Unique
    $dbChanged = $false

    $added   = [System.Collections.Generic.List[string]]::new()
    $removed = [System.Collections.Generic.List[string]]::new()
    $changed = [System.Collections.Generic.List[string]]::new()

    foreach ($k in $allKeys) {
        # Skip known background-noise keys
        if ($EQVarNoise -contains $k) { continue }
        $inB = $bMap.ContainsKey($k)
        $inA = $aMap.ContainsKey($k)
        if ($inB -and $inA) {
            if ($bMap[$k] -ne $aMap[$k]) {
                $bVal = $bMap[$k]; if ($bVal.Length -gt 200) { $bVal = $bVal.Substring(0,200)+'...' }
                $aVal = $aMap[$k]; if ($aVal.Length -gt 200) { $aVal = $aVal.Substring(0,200)+'...' }
                $changed.Add("  ~ $k")
                $changed.Add("      WAS: $bVal")
                $changed.Add("      NOW: $aVal")
            }
        } elseif ($inA) {
            $aVal = $aMap[$k]; if ($aVal.Length -gt 200) { $aVal = $aVal.Substring(0,200)+'...' }
            $added.Add("  + $k = $aVal")
        } else {
            $bVal = $bMap[$k]; if ($bVal.Length -gt 200) { $bVal = $bVal.Substring(0,200)+'...' }
            $removed.Add("  - $k = $bVal")
        }
    }

    if ($added.Count -gt 0 -or $removed.Count -gt 0 -or $changed.Count -gt 0) {
        $anyDbChange = $true
        $dbChanged = $true
        WS "$dbName"
        $changed | ForEach-Object { W $_ }
        $added   | ForEach-Object { W $_ }
        $removed | ForEach-Object { W $_ }
    }
}

if (-not $anyDbChange) {
    W "  No SQLite EQVar changes detected."
}

# ============================================================
# 3. Registry diff
# ============================================================
WH "REGISTRY CHANGES"

function Parse-RegDump {
    param([string]$Path)
    $map = @{}
    if (-not (Test-Path $Path)) { return $map }
    $currentKey = ''
    foreach ($line in [System.IO.File]::ReadAllLines($Path)) {
        if ($line.StartsWith('[[')) {
            $currentKey = $line.Trim('[',']')
        } elseif ($line.StartsWith('  ') -and $currentKey -ne '') {
            $map["$currentKey :: $($line.Trim())"] = $true
        }
    }
    return $map
}

$bReg = Parse-RegDump -Path "$BeforeDir\registry.txt"
$aReg = Parse-RegDump -Path "$AfterDir\registry.txt"

$regAdded   = $aReg.Keys | Where-Object { -not $bReg.ContainsKey($_) } | Sort-Object
$regRemoved = $bReg.Keys | Where-Object { -not $aReg.ContainsKey($_) } | Sort-Object

if ($regAdded.Count -gt 0) {
    WS "Added registry values ($($regAdded.Count))"
    $regAdded | ForEach-Object { W "  + $_" }
}
if ($regRemoved.Count -gt 0) {
    WS "Removed registry values ($($regRemoved.Count))"
    $regRemoved | ForEach-Object { W "  - $_" }
}
if ($regAdded.Count -eq 0 -and $regRemoved.Count -eq 0) {
    W "  No registry changes detected."
}

# ============================================================
# 3. File system diff (noise-filtered)
# ============================================================
WH "FILE SYSTEM CHANGES"

function Parse-FileList {
    param([string]$Path)
    $map = @{}
    if (-not (Test-Path $Path)) { return $map }
    Import-Csv $Path -Delimiter "`t" | ForEach-Object {
        $map[$_.FullName] = @{ Ts = $_.LastWriteTime; Len = $_.Length }
    }
    return $map
}

$bFiles = Parse-FileList -Path "$BeforeDir\files.txt"
$aFiles = Parse-FileList -Path "$AfterDir\files.txt"

$newFiles = $aFiles.Keys | Where-Object {
    -not $bFiles.ContainsKey($_) -and -not (Is-NoisyFile $_)
} | Sort-Object

$deletedFiles = $bFiles.Keys | Where-Object {
    -not $aFiles.ContainsKey($_) -and -not (Is-NoisyFile $_)
} | Sort-Object

$modifiedFiles = $aFiles.Keys | Where-Object {
    $bFiles.ContainsKey($_) -and
    -not (Is-NoisyFile $_) -and (
        $aFiles[$_].Ts  -ne $bFiles[$_].Ts -or
        $aFiles[$_].Len -ne $bFiles[$_].Len
    )
} | Sort-Object

if ($newFiles.Count -gt 0) {
    WS "New files ($($newFiles.Count))"
    $newFiles | ForEach-Object { W "  + $_ ($($aFiles[$_].Ts), $($aFiles[$_].Len) bytes)" }
}
if ($modifiedFiles.Count -gt 0) {
    WS "Modified files ($($modifiedFiles.Count))"
    $modifiedFiles | ForEach-Object {
        W "  M $_ (was: $($bFiles[$_].Ts) $($bFiles[$_].Len)b -> now: $($aFiles[$_].Ts) $($aFiles[$_].Len)b)"
    }
}
if ($deletedFiles.Count -gt 0) {
    WS "Deleted files ($($deletedFiles.Count))"
    $deletedFiles | ForEach-Object { W "  - $_" }
}
if ($newFiles.Count -eq 0 -and $modifiedFiles.Count -eq 0 -and $deletedFiles.Count -eq 0) {
    W "  No file system changes detected (excluding known cache/session paths)."
}

# ============================================================
# Write report
# ============================================================
[System.IO.File]::WriteAllLines($ReportFile, $out, [System.Text.Encoding]::ASCII)
W ""
W "Report written to: $ReportFile"
