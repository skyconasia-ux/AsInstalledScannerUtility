# Checkpoints

---

## Checkpoint 2026-04-17 — Session end (context limit)

### Objective
Collect and document all ControlSuite / Equitrac configuration from a Windows Server 2022 VM (`CSTEMP`, `192.168.60.150`) to support an AI consultant system. Primary deliverable is `Export-ServerInfo.ps1` — a ~1280-line PowerShell script that exports print server config into a structured text file for ConsultantApp.

### Completed This Session
- **Fix 7** — Spool folder location added to `[Print Server]` section (reads `DefaultSpoolDirectory` from registry).
- **Fix 8** — x64 vs x86 driver distinction: per-queue `x86 (32-bit) Driver` field + installed driver list split into `--- x64 ---` / `--- x86 ---` subsections. Root cause: `Get-PrinterDriver` always returns blank `Environment`; fix uses `-PrinterEnvironment` parameter.
- **Fix 9/10** — Virtual queue installable options (Punch/Staple/FreeStaple/Fold/Booklet/Passcode/UserPrompts) confirmed **unresolvable** via any standard API. Proprietary binary blobs (`PrinterData1..N`) with magic `30 52 D5 21`. Documented as known limitation in TASKS.md.
- All 10 fixes applied to `C:\Temp\Export-ServerInfo.ps1` on CSTEMP. Pushed to GitHub commit `979b3b5`.
- Full project docs pushed: README.md, ARCHITECTURE.md, SYSTEM_PROMPT.md, TASKS.md, docs/patterns.md.
- **Registry collection** — `collect_cs_registry.ps1` written and run on CSTEMP. Output: `ControlSuite_Registry.txt` (123 KB). Key findings: Equitrac 6.5.2.191 / ControlSuite 1.5.0.2, 7 EQ* services all Running/Automatic, CAS DB = `eqcas` on `.\SQLExpress`, CAS server = `cstemp.skyconasia.com`, 11 installed roles.
- **DB collection** — `collect_cs_db.ps1` written and run on CSTEMP. Output: `ControlSuite_DB.txt` (11 KB). Key findings: SQL Server 2022 Express, single non-system DB `eqcas` (144 MB), 96 tables using `cas_` / `cat_` prefix, no stored procedures.
- Both output files SCP'd to `C:\Users\quick\` for local review.

### Last Thing Before Context Limit
Reading and reviewing `ControlSuite_DB.txt`. Identified that the DB script's `$configTables` list used `tbl*` naming convention and missed all `cas_*` tables — meaning `cas_config` (800 rows, the main ControlSuite config table) was **not sampled**.

### Next Step
**Query `cas_config` and other populated `cas_*` tables directly.** Run on CSTEMP:
```sql
SELECT TOP 200 * FROM cas_config WITH (NOLOCK) ORDER BY 1
SELECT TOP 50  * FROM cas_pricelist_attributes WITH (NOLOCK)
SELECT TOP 50  * FROM cas_dms_item WITH (NOLOCK)
```
Either extend `collect_cs_db.ps1` to explicitly list `cas_*` tables, or run a quick targeted query via SSH. This is the most valuable remaining data — `cas_config` likely contains all ControlSuite operational settings.

### Local Files of Note
| File | Description |
|------|-------------|
| `C:\Users\quick\ControlSuite_Registry.txt` | 123 KB registry dump from CSTEMP |
| `C:\Users\quick\ControlSuite_DB.txt` | 11 KB DB structure dump from CSTEMP |
| `C:\Users\quick\collect_cs_registry.ps1` | Registry collection script (ASCII-safe) |
| `C:\Users\quick\collect_cs_db.ps1` | DB collection script (DataTable comma-wrap fix) |
| `C:\Temp\Export-ServerInfo.ps1` on CSTEMP | Live script, all 10 fixes applied |
