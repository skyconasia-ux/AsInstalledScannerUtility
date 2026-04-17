# Checkpoints

---

## CURRENT CHECKPOINT — 2026-04-17 (Session 2)

### Objective
Produce a single self-contained `Export-ServerInfo.ps1` that collects all server, print, and solution config (Equitrac/ControlSuite/Nuance/Kofax/Tungsten) into one results file for ConsultantApp AI extraction.

### Completed
- Organized all scripts into `AsInstalledScanner/` folder structure (root: bat+ps1; `tools/`: collection scripts; `results/`: output; `logs/`: trace).
- Fixed critical bug: SSH warning lines baked into top of PS1 caused `**` command error on VM. Removed.
- Added `results\` output folder (auto-created by script).
- Added `logs\TraceLog_*.txt` — real-time section-by-section progress log, also echoes to console.
- Added **Section 6b — Print Management Solution** (~300 lines): 10 subsections covering installed products, all brand services (EQ*/Nuance/Kofax/Tungsten), full registry dump of brand keys + EQ* service params, install directory scan + config file connection string extraction, external DB detection (registry + active TCP), SQL inspection via Windows auth (lists DBs, table row counts, full `cas_config` dump + 10 key tables), scheduled tasks, firewall rules, IIS bindings, event log entries.
- Script pushed to VM: `C:\Temp\AsInstalledScanner\Export-ServerInfo.ps1`. Syntax OK, ~1570 lines.

### Current
Script on VM ready to run. Not yet executed this session.

### Next Step
Run the script on VM and review output:
```powershell
ssh Administrator@192.168.60.150 "powershell -NoProfile -ExecutionPolicy Bypass -File C:\Temp\AsInstalledScanner\Export-ServerInfo.ps1"
scp Administrator@192.168.60.150:"C:/Temp/AsInstalledScanner/results/ServerInfo_CSTEMP_*.txt" C:\Users\quick\
```

### Pending
- Verify Section 6b output quality on CSTEMP (SQL Windows auth may need credential fallback if integrated auth fails)
- `collect_cs_config.ps1` in `tools/` still not run — superseded by Section 6b SQL inspection, but available as standalone if needed
- Review `cas_config` dump in results for completeness

### Files
| Location | File | Description |
|---|---|---|
| Local + VM | `Export-ServerInfo.ps1` | Main script ~1570 lines, syntax OK |
| Local + VM | `Export-ServerInfo.bat` | Launcher |
| Local | `tools/collect_cs_config.ps1` | Standalone cas_* deep-dive (not yet run) |
| Local | `tools/collect_cs_db.ps1` | DB structure (already run) |
| Local | `tools/collect_cs_registry.ps1` | Registry dump (already run) |
| Local | `ControlSuite_Registry.txt` | 123 KB registry dump at `C:\Users\quick\` |
| Local | `ControlSuite_DB.txt` | 11 KB DB structure at `C:\Users\quick\` |

---

## HISTORY

### 2026-04-17 Session 1
- All 10 PS1 fixes applied and pushed (`979b3b5`). Full docs pushed.
- `collect_cs_registry.ps1` run → 123 KB registry dump. Key: Equitrac 6.5.2.191 / ControlSuite 1.5.0.2, 7 EQ* services Running, CAS DB = `eqcas` on `.\SQLExpress`.
- `collect_cs_db.ps1` run → 11 KB DB structure dump. Key: SQL Server 2022 Express, `eqcas` (144 MB), 96 `cas_*`/`cat_*` tables.
- Discovered original DB script missed all `cas_*` tables (used `tbl*` naming). `collect_cs_config.ps1` written as fix.
