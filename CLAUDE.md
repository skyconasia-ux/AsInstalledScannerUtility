# CLAUDE.md — AsInstalledScanner

## Purpose
Standalone PowerShell utility that collects comprehensive server information from a customer's Windows Server. Output is a plain-text data file that can be dropped into ConsultantApp (`data/Deployment/PMS/<CustomerName>/`) for AI-assisted extraction.

## Files
| File | Role |
|------|------|
| `Export-ServerInfo.bat` | Thin launcher — double-click on customer server. Calls the PS1. No logic here. |
| `Export-ServerInfo.ps1` | All logic (~1100 lines). Produces two output files in its own directory. |

## Output Files (written to same folder as script)
| File | Contents |
|------|----------|
| `ServerInfo_<HOST>_<TIMESTAMP>.txt` | Structured data (drop into ConsultantApp) |
| `ExportLog_<HOST>_<TIMESTAMP>.txt` | Audit trail — every query/action logged for cybersec review |

---

## Common Commands (copy-paste ready)

### Syntax check (run this BEFORE every test run)
```powershell
powershell -NoProfile -Command "
$e = $null; $t = $null
[System.Management.Automation.Language.Parser]::ParseFile('C:\Users\Administrator\AsInstalledScanner\Export-ServerInfo.ps1',[ref]$t,[ref]$e) | Out-Null
if ($e.Count -eq 0) { 'Syntax OK' } else { $e | ForEach-Object { 'Line ' + $_.Extent.StartLineNumber + ': ' + $_.Message } }
"
```

### Run the script
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\Administrator\AsInstalledScanner\Export-ServerInfo.ps1"
```

### Syntax check + run in one line
```powershell
powershell -NoProfile -Command "
$e = $null; $t = $null
[System.Management.Automation.Language.Parser]::ParseFile('C:\Users\Administrator\AsInstalledScanner\Export-ServerInfo.ps1',[ref]$t,[ref]$e) | Out-Null
if ($e.Count -eq 0) { 'Syntax OK' } else { $e | ForEach-Object { 'Line ' + $_.Extent.StartLineNumber + ': ' + $_.Message } }
" && powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\Administrator\AsInstalledScanner\Export-ServerInfo.ps1"
```

### Read the latest output file
```powershell
powershell -NoProfile -Command "
$f = Get-ChildItem 'C:\Users\Administrator\AsInstalledScanner\ServerInfo_*.txt' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Get-Content $f.FullName
"
```

### Read a specific section from the latest output
```powershell
# Replace [Print Server] with the section header you want
powershell -NoProfile -Command "
$f = Get-ChildItem 'C:\Users\Administrator\AsInstalledScanner\ServerInfo_*.txt' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Get-Content $f.FullName | Select-String -Pattern '\[Print Server\]' -Context 0,60
"
```

---

## Script Structure (Export-ServerInfo.ps1)

| Lines | Section | Key data collected |
|-------|---------|-------------------|
| 1-86 | Setup / helpers | `Write-Out`, `Write-Log`, `Find-PrintApp`, output file init |
| 87-155 | Section 1 - System Identity | VM/physical, hypervisor, UUID, BIOS, manufacturer/model |
| 156-169 | Section 2 - Operating System | OS name, version, architecture, install date |
| 170-203 | Section 3 - Hardware | RAM, CPU, cores/threads |
| 189-204 | Section 3b - Storage | All drives: total/used/free GB |
| 205-299 | Section 4 - Network | Per-NIC: MAC, IP, subnet, gateway, DNS1-N, DHCP/static, LMHOSTS, NetBIOS, link speed |
| 300-317 | Section 5 - Domain | Domain name, AD membership |
| 318-520 | Section 6 - Installed Software | PaperCut MF/NG, Equitrac/ControlSuite, YSoft SafeQ, AWMS2; checks registry, services, known dirs |
| 521-642 | Section 7 - Roles & Features | .NET 3.5/4.8, Print Services, LPR, Telnet; all installed roles + features |
| 643-890 | Section 8 - Database | SQL Server (registry, TCP/IP, Named Pipes, instances, ports), PostgreSQL, ODBC DSNs, remote DB connections |
| 891-1130 | Section 8b - Print Server | Role status, spooler, per-queue loop (see below), drivers, ports |
| 1131-1150 | Section 8c - Device Counts | FujiFilm/Fuji Xerox/FX device counts |

### Per-Queue Fields (lines ~960-1100)
Each queue outputs:
- **Name, Driver, Port, Status**
- **Advanced Print Features** (RAW_ONLY bit 0x1000 inverted)
- **Render on Client** (bit 0x40000)
- **Paper Size** (`Get-PrintConfiguration.PaperSize`)
- **2-sided/Duplex** (enum: OneSided / TwoSidedLongEdge / TwoSidedShortEdge)
- **Output Color** (`Get-PrintConfiguration.Color` boolean)
- **Staple** (PrintTicket XML XPath `//psf:Feature[contains(@name,'Staple') or contains(@name,'Finishing')]`)
- **Offset Stacking** (PrintTicket XML XPath `//psf:Feature[contains(@name,'OutputBin')]`, excludes InputBin)
- **Use Application Color** -- reads `psk:PageColorManagement` from the default PrintTicket XML.
  This is the "Use the dmColor specified by the application" toggle in the FujiFilm driver UI.
  - `psk:None` = **On** (driver passes application's dmColor through unchanged -- no driver colour override)
  - `psk:System` = **Off** (Windows ICM manages colour conversion)
  - `psk:Driver` = **Off** (driver handles ICM internally)
  - `psk:Device` = **Off** (device hardware handles ICM)
- **ICM Method** -- reads `dmICMMethod` from the Default DevMode binary at byte offset 188 (DWORD).
  Values per `$icmMap`: 1=ICM Disabled, 2=ICM by OS, 3=ICM by device, 4=ICM by host.
  **Note:** `Win32_Printer.ICMMethod` WMI property is unreliable for third-party drivers (always
  returns 0/null); the Default DevMode registry binary is the authoritative source.
- **Installable Options** (via `Get-PrinterProperty`) -- reads hardware configuration stored in the driver:
  - `Finisher Installed` -- derived from `Config:OP_FinisherA/B/C/D` (non-`No` values shown)
  - `Staple Capability` -- from `Config:DC_FIN_Staple`, `Config:DC_FIN_FreeStaple`, `Config:DC_FIN_4Staple`
  - `Punch Capability` -- from `Config:DC_FIN_Punch`, `Config:OP_Punch_2_3`, `Config:OP_Punch_2_4`
  - `Booklet Maker` -- from `Config:OP_Booklet`, `Config:DC_FIN_BiFold`, `Config:DC_FIN_CZFold`
  - `Offset Stacking` -- from `Config:DC_OffsetStacking`
  - Virtual/software drivers (e.g. FF Multi-model Print Driver 2) return no properties -- outputs `N/A`
- **Shared + Share Name, Listed in AD Directory** (Win32_Printer.Attributes bits)
- **x64 driver name, x86 (32-bit) additional driver presence**

---

## Known Gotchas (hard-won -- do not repeat these mistakes)

### 1. Unicode characters break the PS1 parser silently
**Never** use box-drawing or non-ASCII characters in comments or strings.
Use plain ASCII dashes only: `# ===` or `# ---`.

### 2. PowerShell enum value 0 is falsy
`DuplexingMode = OneSided` is enum value `0`. The guard `if ($cfg.DuplexingMode)` evaluates false.
Always use `if ($cfg.DuplexingMode -ne $null)` for enum properties.

### 3. Get-PrintConfiguration.Color is a boolean, not an int
`$cfg.Color` is `$true`/`$false`. Do not cast to `[int]` for color mapping.

### 4. OutputBin XPath must exclude InputBin
`contains(@name,'Bin')` also matches `psk:JobInputBin`. Use:
`contains(@name,'OutputBin') or (contains(@name,'Bin') and not(contains(@name,'Input')))`

### 5. $pg.'Base Directory' -- property names with spaces need quotes
`$pg.'Base Directory'` not `$pg.Base Directory`.

### 6. BAT cannot embed complex PowerShell inline
Any non-trivial PS1 logic must live in a separate `.ps1` file. The BAT is a thin launcher only.

### 7. No global mutable state in output helpers
`Write-Out` appends to `$script:outLines` (array); `Write-Log` appends to `$script:logLines`.
Both are flushed to disk at the end. Do not write to files mid-script.

### 8. Win32_Printer.ICMMethod is unreliable for third-party drivers
For FujiFilm, Kofax, and most OEM drivers, `$wmi.ICMMethod` always returns `0` (null/not set).
**Always read `dmICMMethod` directly from the Default DevMode registry binary:**
```
HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\<PrinterName>\Default DevMode
```
Byte offset 188, DWORD (little-endian). Standard values: 1=disabled, 2=OS, 3=device, 4=host.

### 9. FujiFilm drivers are named "FF *", not "FujiFilm *"
Driver names like `FF Apeos C7071 PCL 6` and `FF Multi-model Print Driver 2` do **not** match
`*FujiFilm*` or `*Fuji Xerox*`. Device count patterns must also include `FF *` and `*Apeos*`:
```powershell
$_.DriverName -like '*FujiFilm*' -or $_.DriverName -like '*Fuji Xerox*' -or
$_.DriverName -like '*FUJIFILM*' -or $_.DriverName -like 'FF *' -or $_.DriverName -like '*Apeos*'
```

### 10. FujiFilm "Use the dmColor specified by the application" = PageColorManagement in PrintTicket
The FujiFilm driver UI setting (Advanced > Printing Defaults > Advanced Settings > Items >
"Use the dmColor specified by the application") maps to `psk:PageColorManagement` in the default
PrintTicket XML retrieved by `Get-PrintConfiguration`:
- `psk:None` = **On** (application's dmColor passes through unchanged)
- `psk:System` / `psk:Driver` / `psk:Device` = **Off** (some form of ICM is active)

Read it with XPath: `//psf:Feature[contains(@name,'PageColorManagement')]//psf:Option`

This is **not** the same as `dmICMMethod` (gotcha #8 above). They are separate concepts:
- `dmICMMethod` = which system handles colour profile conversion
- `PageColorManagement` = whether the driver overrides the application's colour mode choice

### 11. Use Get-PrinterProperty for FujiFilm finisher/staple/punch/booklet installable options
`Get-PrinterProperty -PrinterName <name>` is the correct cmdlet for reading hardware installable
options (finisher unit type, staple, punch, booklet maker). It returns `Config:OP_*` and
`Config:DC_FIN_*` properties with human-readable string values:

```powershell
# Example output for FF Apeos C7071 PCL 6:
Config:OP_FinisherB      = GB_GB4
Config:OP_FinisherC      = GC4_GC5
Config:DC_FIN_Staple     = Yes
Config:DC_FIN_FreeStaple = Yes
Config:DC_FIN_Punch      = Yes
Config:OP_Punch_2_3      = 2_3Holes
Config:OP_Punch_2_4      = 2_4Holes
Config:OP_Booklet        = Booklet
Config:DC_FIN_BiFold     = Yes
Config:DC_FIN_CZFold     = Yes
Config:DC_OffsetStacking = Yes
```

**Virtual/software drivers** (e.g. FF Multi-model Print Driver 2 used for follow-you queues)
return **nothing** from `Get-PrinterProperty` -- this is expected, they have no physical hardware.

The following APIs do **not** work for FujiFilm finisher options and should not be used:
- `System.Printing.PrintQueue.GetPrintCapabilities()` -- returns empty StaplingCapability
- PrintCapabilities XML -- no staple/finisher Feature nodes at all for FF drivers
- `Win32_Printer` WMI -- no finishing properties
- Default PrintTicket XML -- no staple/finisher Feature nodes

### 12. PrintCapabilities XML namespace for FujiFilm Apeos PCL6
The Apeos PCL6 driver uses a custom XML namespace prefix:
`xmlns:ns0000="http://www.fujifilm.com/fb/2021/04/printing/printticket"`
FujiFilm-specific options (paper types, tray names, locale, resolution values) appear as `ns0000:*`.
Standard job features (duplex, color, collate, orientation) still use the `psk:` namespace.
Neither namespace exposes finishing/stapling options -- use `Get-PrinterProperty` instead (gotcha #11).

---

## Adding New Sections -- Checklist
1. Add `# === SECTION N - Name ===` header comment
2. Call `Write-Section "Section Name"` to write the `[Section Name]` header to output
3. Write-Log each WMI/CIM/registry query before executing it
4. Guard every external call with `-ErrorAction SilentlyContinue`
5. Syntax-check before running (see commands above)
6. Run and check the latest output file for the new section

---

## Deployment on Customer Server
1. Copy `Export-ServerInfo.bat` + `Export-ServerInfo.ps1` to the same folder on the customer's server
2. Double-click `Export-ServerInfo.bat` (no install required, no internet needed)
3. Collect `ServerInfo_<HOST>_<TIMESTAMP>.txt`
4. Drop that file into `ConsultantApp\data\Deployment\PMS\<CustomerName>\`
5. Trigger index refresh in ConsultantApp and search by customer name

---

## Session State (last updated 2026-04-17)

**Active VM:** CSTEMP (`192.168.60.150`) — Windows Server 2022, Kofax ControlSuite (Equitrac 6.5.2.191 / ControlSuite 1.5.0.2)

**Script state:** `Export-ServerInfo.ps1` — all 10 fixes applied, ~1278 lines, syntax OK. Live at `C:\Temp\` on CSTEMP. Last pushed: commit `979b3b5`.

**What was just done:**
- Registry dump collected: `ControlSuite_Registry.txt` (123 KB) — key config, all 7 EQ* services Running, CAS DB on `.\SQLExpress`
- DB structure dump collected: `ControlSuite_DB.txt` (11 KB) — 96 tables in `eqcas`, SQL Server 2022 Express
- Both files at `C:\Users\quick\` locally

**Immediate next step:**
Query `cas_config` (800 rows) — the main ControlSuite config table. The DB collection script used `tbl*` naming and missed all `cas_*` tables entirely.
```sql
SELECT TOP 200 * FROM cas_config WITH (NOLOCK) ORDER BY 1
```
Run via SSH on CSTEMP or extend `collect_cs_db.ps1`.

**See also:** `checkpoints.md` for full session checkpoint.
