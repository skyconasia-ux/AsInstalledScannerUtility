# CLAUDE.md — AsInstalledScanner

## Purpose
Standalone PowerShell utility that collects comprehensive server information from a customer's Windows Server. Output is a plain-text data file that can be dropped into ConsultantApp (`data/Deployment/PMS/<CustomerName>/`) for AI-assisted extraction.

## Files
| File | Role |
|------|------|
| `Export-ServerInfo.bat` | Thin launcher — double-click on customer server. Calls the PS1. No logic here. |
| `Export-ServerInfo.ps1` | All logic (~1010 lines). Produces two output files in its own directory. |

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
| 1–86 | Setup / helpers | `Write-Out`, `Write-Log`, `Find-PrintApp`, output file init |
| 87–155 | Section 1 — System Identity | VM/physical, hypervisor, UUID, BIOS, manufacturer/model |
| 156–169 | Section 2 — Operating System | OS name, version, architecture, install date |
| 170–203 | Section 3 — Hardware | RAM, CPU, cores/threads |
| 189–204 | Section 3b — Storage | All drives: total/used/free GB |
| 205–299 | Section 4 — Network | Per-NIC: MAC, IP, subnet, gateway, DNS1-N, DHCP/static, LMHOSTS, NetBIOS, link speed |
| 300–317 | Section 5 — Domain | Domain name, AD membership |
| 318–520 | Section 6 — Installed Software | PaperCut MF/NG, Equitrac/ControlSuite, YSoft SafeQ, AWMS2; checks registry, services, known dirs |
| 521–642 | Section 7 — Roles & Features | .NET 3.5/4.8, Print Services, LPR, Telnet; all installed roles + features |
| 643–890 | Section 8 — Database | SQL Server (registry, TCP/IP, Named Pipes, instances, ports), PostgreSQL, ODBC DSNs, remote DB connections |
| 891–1113 | Section 8b — Print Server | Role status, spooler, per-queue loop (see below), drivers, ports |
| 1114–1133 | Section 8c — Device Counts | FujiFilm/Fuji Xerox/FX device counts |

### Per-Queue Fields (lines ~960–1090)
Each queue outputs:
- Name, Driver, Port, Status
- Advanced Print Features (RAW_ONLY bit 0x1000 inverted)
- Render on Client (bit 0x40000)
- Paper Size (Get-PrintConfiguration.PaperSize)
- 2-sided/Duplex (enum: OneSided / TwoSidedLongEdge / TwoSidedShortEdge)
- Output Color (Get-PrintConfiguration.Color boolean)
- Staple (PrintTicket XML XPath `//psf:Feature[contains(@name,'Staple')]`)
- Offset Stacking (PrintTicket XML XPath `//psf:Feature[contains(@name,'OutputBin')]`, excludes InputBin)
- dmColor / ICM Method (Win32_Printer.ICMMethod: 1=app, 2=OS, 3=device, 4=host)
- Shared + Share Name, Listed in AD Directory (Win32_Printer.Attributes bits)
- x64 driver name, x86 (32-bit) additional driver presence

---

## Known Gotchas (hard-won — do not repeat these mistakes)

### 1. Unicode characters break the PS1 parser silently
**Never** use box-drawing characters (`─`, `—`, `═`) or any non-ASCII in comments or strings.
Use plain ASCII dashes only: `# ===` or `# ---`.

### 2. PowerShell enum value 0 is falsy
`DuplexingMode = OneSided` is enum value `0`. The guard `if ($cfg.DuplexingMode)` evaluates false.
Always use `if ($cfg.DuplexingMode -ne $null)` for enum properties.

### 3. Get-PrintConfiguration.Color is a boolean, not an int
`$cfg.Color` is `$true`/`$false`. Do not cast to `[int]` for color mapping.

### 4. OutputBin XPath must exclude InputBin
`contains(@name,'Bin')` also matches `psk:JobInputBin`. Use:
`contains(@name,'OutputBin') or (contains(@name,'Bin') and not(contains(@name,'Input')))`

### 5. $pg.'Base Directory' — property names with spaces need quotes
`$pg.'Base Directory'` not `$pg.Base Directory`.

### 6. BAT cannot embed complex PowerShell inline
Any non-trivial PS1 logic must live in a separate `.ps1` file. The BAT is a thin launcher only.

### 7. No global mutable state in output helpers
`Write-Out` appends to `$script:outLines` (array); `Write-Log` appends to `$script:logLines`.
Both are flushed to disk at the end. Do not write to files mid-script.

---

## Adding New Sections — Checklist
1. Add `# === SECTION N — Name ===` header comment
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
