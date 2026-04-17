# AsInstalledScanner

Agentless PowerShell utility that exports a comprehensive snapshot of a Windows print server configuration into a structured plain-text file. The output is designed to be dropped directly into **ConsultantApp** for AI-assisted analysis, gap reporting, and quoting.

---

## What It Does

Runs on any Windows Server (no install, no internet, no admin-portal login) and collects:

| Section | What is captured |
|---------|-----------------|
| System Identity | VM/physical, hypervisor type, UUID, BIOS, manufacturer/model |
| Operating System | Name, build, architecture, install date |
| Hardware | RAM, CPU model, cores/threads, all drives (GB) |
| Network | Per-NIC: MAC, IP, subnet, gateway, DNS, DHCP/static, link speed |
| Domain | Domain name, AD forest, membership |
| Installed Software | PaperCut MF/NG, Equitrac/ControlSuite, YSoft SafeQ, AWMS2 |
| Roles & Features | .NET, Print Services, LPR, Telnet, all Windows features |
| Database | SQL Server (instances, ports, protocols), PostgreSQL, ODBC DSNs |
| Print Server | Role, spooler, spool folder, per-queue full config (see below) |
| Device Counts | FujiFilm / Fuji Xerox / FF driver queue totals |

### Per-Print-Queue Detail
For every queue the export captures:
- Driver name, port(s), status
- Duplex, paper size, output colour
- Staple / offset stacking from default PrintTicket XML
- **Use Application Color** (`psk:PageColorManagement` — the "pass dmColor through" toggle)
- **ICM Method** (read from DEVMODE binary offset 188, not unreliable WMI)
- **Installable Options** — finisher units, staple, punch, booklet maker from `Get-PrinterProperty`
- x64 and x86 (32-bit) driver availability
- Sharing and Active Directory publication status

---

## Quick Start

### On the customer's print server

1. Copy `Export-ServerInfo.bat` and `Export-ServerInfo.ps1` into the **same folder**.
2. Double-click `Export-ServerInfo.bat`.  
   No install, no internet, runs as whatever Windows user you're logged in as (needs local Admin).
3. Two files are written to the same folder:
   - `ServerInfo_<HOSTNAME>_<TIMESTAMP>.txt` — the data file
   - `ExportLog_<HOSTNAME>_<TIMESTAMP>.txt` — audit log of every query

### Loading into ConsultantApp

1. Copy `ServerInfo_<HOST>_<TIMESTAMP>.txt` to:  
   `ConsultantApp\data\Deployment\PMS\<CustomerName>\`
2. Trigger index refresh in ConsultantApp.
3. Search by customer name — the AI will answer questions from the structured data.

---

## Requirements

| Requirement | Detail |
|------------|--------|
| OS | Windows Server 2012 R2 or later (tested on 2019, 2022) |
| Shell | PowerShell 5.1+ (ships with Server 2012 R2+) |
| Permissions | Local Administrator (for registry, WMI, printer enumeration) |
| Network | None — runs entirely locally |
| Install | None — xcopy deploy, delete when done |

---

## Files

```
AsInstalledScanner/
├── Export-ServerInfo.bat      # Thin launcher — double-click this
├── Export-ServerInfo.ps1      # All logic (~1280 lines)
├── README.md                  # This file
├── ARCHITECTURE.md            # Design decisions, data sources, limitations
├── SYSTEM_PROMPT.md           # AI system prompt for ConsultantApp integration
├── TASKS.md                   # Backlog and known improvements
├── CLAUDE.md                  # AI coding assistant instructions
└── docs/
    └── patterns.md            # Field extraction patterns for ConsultantApp
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

> **Note:** FujiFilm FF drivers use the `FF *` prefix — they do **not** contain "FujiFilm" in the driver name.

---

## Development & Contributing

See [`CLAUDE.md`](CLAUDE.md) for coding conventions, known gotchas, and the AI assistant instruction set used when modifying this project.

See [`ARCHITECTURE.md`](ARCHITECTURE.md) for design rationale.

See [`TASKS.md`](TASKS.md) for the improvement backlog.
