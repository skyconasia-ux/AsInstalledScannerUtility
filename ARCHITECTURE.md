# Architecture

## Overview

`tools/AsInstalledScanner.ps1` is the current main entry point for the project. It is a PowerShell 5.1 scanner that runs against Windows print server environments and collects both server inventory data and Kofax ControlSuite / Equitrac configuration data.

The root `AsInstalledScanner.bat` file is a thin launcher. It starts the PowerShell script in interactive mode when no argument is supplied, or passes a mode such as `Before`, `After`, `Compare`, or `Full` to the script for non-interactive use.

```
Customer Windows print server / scanner host
|
+-- AsInstalledScanner.bat
    |
    +-- tools/AsInstalledScanner.ps1
        |
        +-- Windows inventory
        |   +-- CIM / WMI
        |   +-- Registry
        |   +-- PrintManagement cmdlets
        |   +-- Network, storage, software, SQL Server, roles/features
        |
        +-- ControlSuite / Equitrac inventory
        |   +-- SQL Server eqcas data
        |   +-- SQLite EQVar databases
        |   +-- Kofax / Equitrac registry and service data
        |   +-- Workflow, pricing, pull group, authentication, SMTP, quota,
        |       device, license, and directory sync settings
        |
        +-- Output folder
            +-- *_FULL.txt
            +-- *_SUMMARY.txt
            +-- *_REPORT.html
            +-- metadata.json
            +-- raw snapshot files used by Compare mode
```

The project also keeps older and lower-level helper scripts under `tools/` for Equitrac-specific discovery:

- `EQ-Snapshot.ps1` captures BEFORE or AFTER Equitrac / ControlSuite state.
- `EQ-Diff.ps1` compares snapshot folders and reports changed SQL rows, EQVar keys, registry values, and file timestamps.
- `Export-EquitracConfig.ps1` exports ControlSuite / Equitrac configuration details.

Those helpers are useful for targeted investigation, but `tools/AsInstalledScanner.ps1` is the consolidated scanner/report generator.

---

## Runtime Modes

| Mode | Purpose |
|------|---------|
| `Before` | Capture a baseline snapshot before changes. |
| `After` | Capture current state and write text, JSON, and HTML report outputs. |
| `Compare` | Compare the latest Before and After snapshot folders. |
| `Full` | Run After and then Compare. |
| `Settings` | Configure scanner settings used by remote, SQL, or collector workflows. |
| `LocalCollector` | Collect local Windows/server data for later combined reporting. |
| `BuildCombined` | Build a combined report from collected local/remote data. |

---

## Design Decisions

### 1. Main scanner plus thin launcher

**Decision:** The current scanner logic lives in `tools/AsInstalledScanner.ps1`; `AsInstalledScanner.bat` only locates and invokes it.

**Reason:** Deployment stays simple for customer sites while keeping the main implementation in one reviewable PowerShell script.

### 2. Local-first collection with optional remote support

**Decision:** The scanner can collect from the local host and also includes settings and code paths for WinRM, SQL-only, and collector-style workflows.

**Reason:** Some customer environments allow direct local scanning, while others require collecting Windows data separately or limiting access to SQL/configuration sources.

### 3. Human-readable and machine-readable outputs

**Decision:** The scanner writes structured text, summary text, metadata JSON, and a self-contained HTML report.

**Reason:** Consultants need a readable report for review and handoff, while downstream tooling can consume the text and JSON outputs.

### 4. Snapshot and diff workflow

**Decision:** Before/After/Compare modes preserve raw snapshot files and generate comparison reports.

**Reason:** ControlSuite and Equitrac configuration changes can be spread across SQL tables, SQLite EQVar keys, registry values, and files. A diff workflow makes those changes visible after a UI or configuration update.

### 5. Registry and driver-specific print data

**Decision:** Print queue details use Windows print cmdlets where possible, with registry fallbacks for driver data such as DEVMODE fields.

**Reason:** Third-party print drivers do not always expose complete data through WMI or standard PrintTicket XML.

---

## Data Sources

| Area | Primary Sources |
|------|-----------------|
| System identity | `Win32_ComputerSystem`, `Win32_BIOS`, `Win32_BaseBoard`, registry fallback |
| Operating system | `Win32_OperatingSystem` |
| Hardware and storage | CIM/WMI, `Get-PSDrive`, disk and memory classes |
| Network | `Get-NetAdapter`, `Get-NetIPConfiguration`, DNS client cmdlets |
| Domain | Active Directory cmdlets/APIs where available, environment fallback |
| Installed software | Registry uninstall hives, known services, known paths |
| Windows roles/features | `Get-WindowsFeature` when available |
| SQL Server | SQL Server registry keys, services, and SQL connection data |
| Print queues | `Get-Printer`, `Get-PrintConfiguration`, PrintTicket XML, registry |
| Print drivers | `Get-PrinterDriver` by printer environment |
| ControlSuite / Equitrac SQL | `eqcas` SQL tables and row/content dumps |
| ControlSuite / Equitrac EQVar | SQLite EQVar databases such as DCE and DRE config stores |
| ControlSuite / Equitrac system data | Kofax, Equitrac, Nuance, Tungsten, FLEXlm/Flexera registry/service/file locations |

---

## Outputs

A normal `After` or `Full` run writes to a timestamped folder under `Output/`. The main report artifacts are:

| File | Purpose |
|------|---------|
| `After_FULL.txt` | Full structured text export for review and downstream ingestion. |
| `After_SUMMARY.txt` | Short count and key-setting summary. |
| `After_REPORT.html` | Self-contained HTML report for browser review. |
| `metadata.json` | Machine-readable metadata and high-level scan details. |
| `eqvar_*.tsv`, `bcp_*.txt` | Raw snapshot inputs used by comparison workflows. |

Compare mode produces diff-oriented text and HTML outputs that show changed configuration values between snapshots.

---

## Known Limitations

### Virtual driver installable options

FF Multi-model Print Driver 2 and similar virtual/follow-you queues may store hardware option data in proprietary driver blobs. The scanner can identify that the data is not exposed through standard APIs, but it may not be able to decode every setting.

### PrintCapabilities XML coverage

Some FujiFilm and third-party print drivers omit finishing features from standard PrintCapabilities XML. The scanner uses `Get-PrinterProperty` and registry data where those sources are more reliable.

### WMI print driver fields

Some WMI printer fields are incomplete for third-party drivers. For example, ICM method is read from the DEVMODE registry binary rather than relying on `Win32_Printer.ICMMethod`.

### Windows feature availability

`Get-WindowsFeature` is only available on Windows Server systems with the relevant module installed. The scanner handles missing feature cmdlets gracefully.

---

## Related Documentation

- `README.md` contains run instructions and user-facing usage notes.
- `TASKS.md` tracks completed work, backlog items, and known scanner limitations.
- `docs/equitrac-storage-map.md` documents where confirmed ControlSuite / Equitrac settings are stored.
- `docs/patterns.md` documents extraction patterns for structured text output.
