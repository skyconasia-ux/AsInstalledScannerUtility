# AsInstalledScanner

Agentless PowerShell utility that exports a comprehensive snapshot of a Windows print server + Kofax ControlSuite / Equitrac configuration into a self-contained **HTML report** and structured plain-text files. Designed for pre-sales assessments, upgrade documentation, and change-tracking.

---

## Quick Start

### 1. Download

Download the latest release ZIP from the [Releases page](https://github.com/skyconasia-ux/AsInstalledScannerUtility/releases/latest) and extract it to any folder on your laptop.

### 2. Copy to the customer server

Copy the `tools\` folder and `AsInstalledScanner.bat` to a temporary location on the customer's Windows print server — for example `C:\Temp\`.

```
C:\Temp\
  AsInstalledScanner.bat
  tools\
    AsInstalledScanner.ps1
    sqlite3.exe
```

### 3. Run

**Double-click `AsInstalledScanner.bat`** for the interactive menu, or run silently via SSH / remote session:

```powershell
# Capture full snapshot + HTML report in one pass
powershell -NoProfile -ExecutionPolicy Bypass -File C:\Temp\tools\AsInstalledScanner.ps1 -Mode Full
```

Output is written to a timestamped subfolder:

```
C:\Temp\tools\Output\HOSTNAME_YYYYMMDD_HHMMSS_After\
  After_REPORT.html     <- open this in any browser
  After_FULL.txt        <- structured plain-text (drop into ConsultantApp)
  After_SUMMARY.txt     <- quick counts + key settings
  metadata.json         <- machine-readable
  eqvar_dce.tsv         <- raw EQVar snapshot (used by Compare mode)
  eqvar_dre.tsv
  bcp_*.txt             <- SQL table dumps (used by Compare mode)
```

### 4. View the report

Copy `After_REPORT.html` back to your laptop and open it in any browser. No server required — the file is fully self-contained.

---

## Modes

| Mode | What it does |
|------|-------------|
| `Before` | Capture a baseline snapshot before making changes (raw data only, no report) |
| `After` | Capture current state and generate FULL.txt + SUMMARY.txt + REPORT.html |
| `Compare` | Diff the latest Before vs After snapshot — shows every changed EQVar key and SQL row |
| `Full` | After + Compare in a single pass |

---

## What It Captures

### Windows / Print Server

| Section | Detail |
|---------|--------|
| System Identity | Platform (VM/physical), hypervisor, UUID, BIOS, manufacturer/model |
| Operating System | Name, build, architecture, install date |
| Hardware | RAM, CPU, all drives with size and free space |
| Network | Per-NIC: MAC, IP, subnet, gateway, DNS, DHCP/static, speed |
| Domain | Domain name, AD forest, membership |
| Installed Software | ControlSuite, Equitrac, PaperCut, YSoft SafeQ, AWMS2 |
| Roles & Features | .NET versions, Print Services, all Windows features |
| SQL Server | Per-instance: edition, version, auth mode, service account, sysadmin members, install/data/log/backup paths, error log, Named Pipes, TCP/IP (per-IP address table), FILESTREAM, max/min memory, User Instances, TempDB files |
| Print Queues | Driver, port, duplex, paper size, colour, staple, ICM method, installable options, x64/x86 driver, sharing, AD publication |

### Kofax ControlSuite / Equitrac

| Section | Detail |
|---------|--------|
| Authentication | Auth method, card swipe, PIN settings, AD/LDAP sync |
| SMTP / Email | Server, auth, from address |
| Job Management | Expiry, escrow, release behaviour, offline lifetime |
| Quotas & Messages | Color quota mode, enforcement, account limit, custom messages |
| Currency & Accounting | Currency, cost preview, colour multiplier |
| License Server | FlexNet host, port, protocol |
| Device Settings | Page size, timeouts, keypad, billable feature, billing code prompt |
| Workflows | Per-workflow: type, name, output folder, enabled status |
| Pull Groups | Name, members |
| Pricing | Price lists with rates |
| EQVar Dump | Full SQLite EQVar key-value export (DCE + DRE databases) |

---

## Change Tracking (Before / After / Compare)

Run `Before` before making changes in the ControlSuite Web UI, then `After` when done. Run `Compare` (or `Full`) to get a diff report showing exactly which EQVar keys changed (with WAS / NOW values) and which SQL rows were added, removed, or updated.

```powershell
# Before making changes
powershell -NoProfile -ExecutionPolicy Bypass -File AsInstalledScanner.ps1 -Mode Before

# ... make changes in ControlSuite Web UI ...

# After making changes — generates diff + full report
powershell -NoProfile -ExecutionPolicy Bypass -File AsInstalledScanner.ps1 -Mode Full
```

---

## Requirements

| Requirement | Detail |
|------------|--------|
| OS | Windows Server 2012 R2 or later (tested on 2019, 2022) |
| PowerShell | 5.1+ (ships with Server 2012 R2+) |
| Permissions | Local Administrator |
| Network | None — runs entirely locally |
| Install | None — xcopy deploy, delete when done |
| SQL access | `sa` credentials required for ControlSuite SQL Server queries |
| SQLite | `sqlite3.exe` included in the `tools\` folder |

---

## Files

```
AsInstalledScannerUtility/
├── AsInstalledScanner.bat         # Launcher — double-click on customer server
├── tools/
│   ├── AsInstalledScanner.ps1     # Primary tool (~1600 lines)
│   └── sqlite3.exe                # SQLite CLI (bundled)
├── README.md
├── ARCHITECTURE.md
├── TASKS.md
└── docs/
    ├── equitrac-storage-map.md    # Where every ControlSuite setting lives in EQVar
    └── patterns.md                # ConsultantApp extraction patterns
```

---

## Supported Print Driver Families

| Driver Pattern | Detected As |
|---------------|-------------|
| `FF *` (e.g., FF Apeos C7071 PCL 6) | FujiFilm FF |
| `FF Multi-model Print Driver 2` | FujiFilm virtual/follow-you |
| `*FujiFilm*`, `*FUJIFILM*`, `*Apeos*` | FujiFilm |
| `*Fuji Xerox*`, `*FX*` | Fuji Xerox |
| `Kofax Universal Print Driver` | Kofax ControlSuite |

---

## Development

See [`CLAUDE.md`](CLAUDE.md) for coding conventions and AI assistant instructions.  
See [`ARCHITECTURE.md`](ARCHITECTURE.md) for design rationale.  
See [`TASKS.md`](TASKS.md) for the backlog.
