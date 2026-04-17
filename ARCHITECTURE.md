# Architecture

## Overview

`Export-ServerInfo.ps1` is a single-file PowerShell script (~1280 lines). It runs locally on the target Windows Server, collects data from multiple Windows subsystems, and writes a single structured text file. No network calls, no external dependencies, no persistent state.

```
┌─────────────────────────────────────────────────────────┐
│  Customer Windows Server (print server)                 │
│                                                         │
│  Export-ServerInfo.bat                                  │
│       │                                                 │
│       └─► Export-ServerInfo.ps1                         │
│               │                                         │
│               ├─► WMI / CIM  (hardware, OS, printers)  │
│               ├─► Registry   (DEVMODE, drivers, spool)  │
│               ├─► PowerShell cmdlets (printers, roles)  │
│               ├─► .NET APIs  (printing, XML parsing)    │
│               └─► Filesystem (software detection)       │
│                                                         │
│       ServerInfo_<HOST>_<TS>.txt  ◄── structured output │
│       ExportLog_<HOST>_<TS>.txt   ◄── audit log         │
└─────────────────────────────────────────────────────────┘
         │
         │  (copy manually)
         ▼
┌─────────────────────────────────────────────────────────┐
│  ConsultantApp                                          │
│  data\Deployment\PMS\<CustomerName>\ServerInfo_*.txt   │
│                                                         │
│  AI reads structured sections → answers questions       │
└─────────────────────────────────────────────────────────┘
```

---

## Design Decisions

### 1. Single PS1 file + thin BAT launcher
**Decision:** All logic in one PowerShell file. The BAT does nothing except call it.

**Reason:** Customers can review one file. No module installation, no `Import-Module`. Xcopy deploy, xcopy remove.

**Trade-off:** The script is long (~1280 lines). Accepted; sections are clearly delimited with `# === SECTION N ===` comments.

### 2. Plain-text structured output (not JSON/XML/CSV)
**Decision:** Output is human-readable indented text, not machine-native format.

**Reason:** ConsultantApp's AI ingestion works on freeform text; JSON would be harder for humans to visually review and spot gaps. The structure is consistent enough for pattern extraction.

**Trade-off:** Parsing requires regex/pattern matching rather than schema validation. See [`docs/patterns.md`](docs/patterns.md).

### 3. Two output files (data + log)
**Decision:** Data goes to `ServerInfo_*.txt`, every query/action goes to `ExportLog_*.txt`.

**Reason:** Cybersecurity review — the customer or their IT security team can audit exactly what was queried, in what order, with what result. The log file is also useful for debugging failures.

### 4. In-memory accumulation, flush at end
**Decision:** `Write-Out` appends to `$script:outLines`; `Write-Log` appends to `$script:logLines`. Both are written to disk at the very end.

**Reason:** Avoids partial/corrupt files if the script is interrupted mid-run. Avoids repeated file I/O.

**Gotcha:** No mid-script file writes. Any code that calls `Add-Content` or `Out-File` during the run breaks this contract.

### 5. Registry over WMI for driver data
**Decision:** `dmICMMethod` is read from the Default DevMode registry binary (offset 188), not `Win32_Printer.ICMMethod`.

**Reason:** Third-party drivers (FujiFilm FF, Kofax) do not populate `Win32_Printer.ICMMethod` reliably. It always returns 0 for these drivers. The registry binary is authoritative.

**Location:** `HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\<name>\Default DevMode`  
**Offset:** 188 (DWORD, little-endian) = `dmICMMethod` field per DEVMODE spec.

### 6. Get-PrinterProperty for installable options (physical queues)
**Decision:** `Get-PrinterProperty -PrinterName <name>` is the correct API for finisher/staple/punch/booklet installable options on physical printer queues.

**Reason:** Neither PrintCapabilities XML, System.Printing, nor WMI expose these for FujiFilm FF drivers. `Get-PrinterProperty` reads the driver's `FeatureKeyword` binary blob in `PrinterDriverData`, which contains `Config:OP_*` and `Config:DC_FIN_*` entries in human-readable ASCII.

**Limitation:** Virtual/follow-you queues (FF Multi-model Print Driver 2) do not populate `FeatureKeyword`. `Get-PrinterProperty` returns nothing for them. Their installable options are stored in proprietary binary blobs (`PrinterData1..N`) in `PrinterDriverData` with an undocumented format.

### 7. Get-PrinterDriver -PrinterEnvironment for x86/x64 split
**Decision:** Driver architecture is determined by querying `Get-PrinterDriver` with `-PrinterEnvironment 'Windows x64'` and `-PrinterEnvironment 'Windows NT x86'` separately.

**Reason:** `Get-PrinterDriver` without `-PrinterEnvironment` returns all drivers from all environments but leaves the `Environment` property blank, making it impossible to distinguish architectures from the object itself.

---

## Data Sources by Section

| Section | Primary Source | Fallback |
|---------|---------------|---------|
| System Identity | `Get-CimInstance Win32_ComputerSystem`, `Win32_BIOS`, `Win32_BaseBoard` | Registry `HKLM:\HARDWARE\DESCRIPTION` |
| OS | `Get-CimInstance Win32_OperatingSystem` | — |
| Hardware | `Win32_Processor`, `Win32_PhysicalMemory` | — |
| Storage | `Get-PSDrive` (logical), `Win32_DiskDrive` | — |
| Network | `Get-NetAdapter`, `Get-NetIPConfiguration`, `Get-DnsClientServerAddress` | — |
| Domain | `Get-ADDomain`, `[System.DirectoryServices.ActiveDirectory.Domain]` | `$env:USERDOMAIN` |
| Software | Registry uninstall hives + known service/path checks | — |
| Roles | `Get-WindowsFeature` | — |
| Database | Registry `HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server` | — |
| Print queues | `Get-Printer`, `Get-PrintConfiguration`, PrintTicket XML | `Win32_Printer` (WMI) |
| Printer drivers | `Get-PrinterDriver -PrinterEnvironment` (x64 and x86) | — |
| ICM Method | `Default DevMode` registry binary byte 188 | — |
| Color mgmt | `psk:PageColorManagement` from PrintTicket XML | — |
| Installable opts | `Get-PrinterProperty` (`Config:OP_*`, `Config:DC_FIN_*`) | — |
| Spool folder | `HKLM:\...\Print\Printers` `DefaultSpoolDirectory` | Default path |

---

## Known Limitations

### Virtual driver installable options
FF Multi-model Print Driver 2 (follow-you/ControlSuite virtual queue) stores its device option settings (Punch/Staple/Booklet installed, passcode length, user prompts) in proprietary binary blobs (`PrinterData1..N`) under `PrinterDriverData`. There is no documented API to decode these. The export notes this as a limitation and directs the reviewer to Printer Properties > Device Settings tab.

### PrintCapabilities XML vs Get-PrinterProperty
FujiFilm Apeos PCL6 PrintCapabilities XML does not expose stapling/finishing Feature nodes at all. `Get-PrinterProperty` is the correct API — it reads the `FeatureKeyword` blob the driver writes to `PrinterDriverData`. Do not attempt to use PrintCapabilities XML for installable options on FF drivers.

### WMI ICMMethod
`Win32_Printer.ICMMethod` always returns 0 for third-party drivers. Use DEVMODE binary offset 188.

### FujiFilm Apeos PCL6 PrintCapabilities namespace
Uses `xmlns:ns0000="http://www.fujifilm.com/fb/2021/04/printing/printticket"` for device-specific options. Standard `psk:` namespace still applies to job settings (duplex, color, etc.).

### Get-WindowsFeature availability
Only available on Windows Server with RSAT. Not available on Windows 10/11 workstations. Script handles this gracefully.

---

## Output File Format

```
============================================================
AsInstalled Server Export - HOSTNAME
Generated: YYYY-MM-DD HH:MM:SS
============================================================

[Section Name]
FieldName: Value
FieldName: Value

[Next Section]
...
```

Sections are delimited by `[Section Name]` headers. Fields follow a `Name: Value` pattern. Multi-line items use indented continuation. Print queue entries use `===` separator lines.

See [`docs/patterns.md`](docs/patterns.md) for extraction patterns.
