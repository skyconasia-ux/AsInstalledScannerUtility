# Equitrac / ControlSuite 6 — Configuration Storage Map

Discovered via BEFORE/AFTER snapshot diff on **ControlSuite 1.5.0.2 (Equitrac 6.5.2.191)**.
Server: Windows Server 2022, SQL Server Express, `eqcas` database.

---

## Where Configuration Lives

### 1. SQLite EQVar databases (PRIMARY config store)

Almost all system-level settings live here as key-value rows in an `EQVar` table.

| Database | Path | Service |
|----------|------|---------|
| `DCE_config.db3` | `C:\Windows\System32\config\systemprofile\AppData\Local\Equitrac\Equitrac Platform Component\EQDCESrv\Cache\DCE_config.db3` | EQDCESrv |
| `DREEQVar.db3` | `...EQDRESrv\EQSpool\DREEQVar.db3` | EQDRESrv |
| `EQDMECache.db3` | `...EQDMESrv\EQDMECache.db3` | EQDMESrv |

**Schema:** `System | SubSystem | Class | Name | Value`

**Key:** `SubSystem|Class|Name` (SubSystem is usually `cas`, `dce`, or `dre`; Class and Name are usually empty strings)

**Mirroring:** `cas||` and `dce||` keys appear in **both** `DCE_config.db3` and `DREEQVar.db3`. `dre||` keys are exclusive to `DREEQVar.db3`. The mirror is by design — DRE has a full config copy for offline operation.

**Access:** `sqlite3.exe -separator TAB <db> "SELECT SubSystem,Class,Name,Value FROM EQVar ORDER BY SubSystem,Class,Name;"`

#### Known EQVar keys (discovered)

| SubSystem | Key (Class||Name) | What it controls |
|-----------|-------------------|-----------------|
| `dce` | `\|\|clientauthconfig` | External authentication server URL/config |
| `cas` | `\|\|cardswipeenabled` | Card/badge swipe authentication on/off |
| `cas` | `\|\|cardstorage` | Card number storage type (primary/secondary) |
| `cas` | `\|\|registerpinaslternate` | Allow PIN as alternate to card swipe (0/1) |
| `cas` | `\|\|upgrademode` | Card upgrade mode setting |
| `cas` | `\|\|cardauththreshold` | Card auth threshold |
| `cas` | `\|\|smtpserver` | SMTP server hostname |
| `cas` | `\|\|smtpport` | SMTP port |
| `cas` | `\|\|smtpfromaddress` | From address for email notifications |
| `cas` | `\|\|smtpusername` | SMTP auth username |
| `cas` | `\|\|smtpssl` | SMTP SSL enabled flag |
| `cas` | `\|\|smtpauthentication` | SMTP authentication method |
| `cas` | `\|\|jobexpirytime` | Default job expiry time (minutes) |
| `cas` | `\|\|distributionlistjobexpirytime` | Distribution list job expiry (minutes) |
| `cas` | `\|\|precision` | Accounting decimal precision (e.g. 2 or 3) |
| `cas` | `\|\|colorquotamessage` | Message shown when color quota exceeded |
| `cas` | `\|\|quotaenabled` | Quota enforcement enabled |
| `cas` | `\|\|maxjobsperpullgroup` | Max queued jobs per pull group |
| `cas` | `\|\|workflowfolderslastupdatetime` | Internal sync timestamp — NOISE, ignore in diffs |
| `cas` | `\|\|accesspermissions` | EQReports / Access Permissions config |

---

### 2. eqcas SQL Server database (user/workflow/pricing data)

**Server:** `.\SQLExpress`  **DB:** `eqcas`  **Credentials:** `sa` / `FujiFilm_11111`

> **Critical quirk:** All columns in every eqcas table have **empty string names** (`''`) in `sys.columns`.
> `SELECT *`, named column references, and `BINARY_CHECKSUM(*)` all fail with SQL Error 1037.
> **Only BCP can read table data** (reads by ordinal position, bypasses column name issue).
> Row counts use `sys.dm_db_partition_stats` (no column names needed).

#### Config tables (BCP-readable)

| Table | What it stores | Notes |
|-------|---------------|-------|
| `cas_config` | Key-value config store (1700+ rows) | Subset of settings not in SQLite |
| `cas_scan_alias` | Workflow scan destinations (name, scope, destination, active) | Workflow renames detected here |
| `cat_pricelist` | Pricing lists with XML price data per page type | UPDATE and INSERT both detected via BCP diff |
| `cas_pullgroups` | Pull-print / FollowYou groups | |
| `cas_workflow_folders` | Workflow folder assignments | |
| `cas_user_ext` | User accounts and properties | |
| `cat_validation` | Billing/PIN validation rules | |
| `cas_prq_device_ext` | Per-device pull-print settings | |
| `cas_installedsoftware` | Installed component versions | |

#### Transaction/log tables (row count only — skip full export)

`cas_trx_*`, `cas_uplink_trx_*`, `cas_fas_trx_*`, `cas_audit_log`, `cas_sdr_history`,
`cas_printer_status_history`, `cas_dashboard_data`, `cas_mru_bc_by_user`, `cas_spe_jobs`,
`cat_transaction`, `cat_trxid`, `cat_trxval`, `cat_trxxml`, `cas_report_bitmap`,
`cas_update_sequence_number`

#### cat_pricelist BCP format

```
ID | Name | Description | Type | XMLBlob
```

The XML blob contains `<range min="N" rate="0.0NNNNN" />` elements per page type.
Example: `511|Another Price List| |basic|<catcost><catcostranges>...`

---

### 3. Registry

Covered by `EQ-Snapshot.ps1`. Key roots:

| Root | Contents |
|------|----------|
| `HKLM:\SOFTWARE\Kofax` | Install paths, version info |
| `HKLM:\SOFTWARE\Equitrac` | Legacy/compatibility keys |
| `HKLM:\SOFTWARE\FLEXlm License Manager` | License server config |
| `HKLM:\SOFTWARE\Flexera Software` | FlexNet license data |
| `HKLM:\SYSTEM\CurrentControlSet\Services\EQCASSrv` | CAS service config |
| `HKLM:\SYSTEM\CurrentControlSet\Services\EQDCESrv` | DCE service config |
| `HKLM:\SYSTEM\CurrentControlSet\Services\EQDRESrv` | DRE service config |
| `HKLM:\SYSTEM\CurrentControlSet\Services\EQDMESrv` | DME service config |
| `HKLM:\SYSTEM\CurrentControlSet\Services\flexnetls` | FlexNet license server |
| `HKLM:\SYSTEM\CurrentControlSet\Services\NDISecurityService` | NDI security service |

---

### 4. File system (watch directories)

| Directory | What changes here |
|-----------|------------------|
| `C:\Program Files\Kofax` | Binaries, version files |
| `C:\ProgramData\Equitrac` | Runtime config, logs |
| `C:\ProgramData\Kofax` | ControlSuite runtime data |
| `C:\ProgramData\Nuance` | Legacy Nuance data |
| `C:\ProgramData\flexnetsas` | FlexNet license data |
| `C:\Windows\ServiceProfiles\NetworkService\flexnetls` | FlexNet license server data |
| `C:\Windows\System32\config\systemprofile\AppData\Local\Equitrac` | SQLite .db3 files (primary config) |
| `C:\Users\Administrator\AppData\Local\Equitrac` | If CAS runs as Administrator |
| `C:\Temp`, `C:\Windows\Temp` | Installer/upgrade artifacts |

#### File noise filter (suppress from diffs)

These paths change on every run and carry no configuration signal:

| Pattern | Reason |
|---------|--------|
| `EQ_Snapshots` | Snapshot tool's own output |
| `EQ_BcpDump` | BCP temp exports |
| `EQ_DbQuery` | SQLite temp copies |
| `ClientCommunicationsCache` | Client session cache |
| `aa_token` | Auth token files |
| `finder_cache` | Device finder cache |
| `[.]log$` | All log files |
| `[.]ldf$` | SQL log database files |
| `[.]mdf$` | SQL data database files |
| `NDISecurity\\CS\\cache` | NDI HTTP cache |
| `ssds`, `ssds_seed_nodes` | Distributed service discovery |

---

## Discovery Method

1. Run `EQ-Snapshot.ps1 -Label BEFORE` on the server
2. Make changes via ControlSuite Web UI
3. Run `EQ-Snapshot.ps1 -Label AFTER`
4. Run `EQ-Diff.ps1` — shows exactly what changed where

### What each diff section covers

| Diff section | Detects |
|-------------|---------|
| SQL row count changes | Table INSERT/DELETE (e.g. new price list, new user) |
| SQL BCP content diff | Table UPDATE (e.g. rename workflow, edit price values) |
| SQLite EQVar diff | System setting changes (auth, SMTP, job expiry, card config) |
| Registry diff | Service config, license info changes |
| File system diff | New/modified/deleted files (upgrades, installs, config exports) |

---

## Confirmed Change Mappings (tested)

| UI Action | Storage location | Detected by |
|-----------|-----------------|-------------|
| External auth server change | `DCE_config.db3` → `dce\|\|clientauthconfig` | SQLite EQVar diff |
| Card swipe disable/enable | `DCE_config.db3` → `cas\|\|cardswipeenabled` | SQLite EQVar diff |
| Card storage type change | `DCE_config.db3` → `cas\|\|cardstorage` | SQLite EQVar diff |
| PIN as alternate (on/off) | `DCE_config.db3` → `cas\|\|registerpinaslternate` | SQLite EQVar diff |
| SMTP server change | `DCE_config.db3` → `cas\|\|smtpserver` | SQLite EQVar diff |
| Color quota message | `DCE_config.db3` → `cas\|\|colorquotamessage` | SQLite EQVar diff |
| Job expiry time | `DCE_config.db3` → `cas\|\|jobexpirytime` | SQLite EQVar diff |
| Distribution list expiry | `DCE_config.db3` → `cas\|\|distributionlistjobexpirytime` | SQLite EQVar diff |
| Precision (decimal places) | `DCE_config.db3` → `cas\|\|precision` | SQLite EQVar diff |
| Workflow rename (Mail→Email) | `eqcas..cas_scan_alias` | SQL BCP content diff |
| New price list created | `eqcas..cat_pricelist` (INSERT) | SQL BCP content diff |
| Existing price values edited | `eqcas..cat_pricelist` (UPDATE) | SQL BCP content diff |
| Access Permissions / EQReports | `DCE_config.db3` → `cas\|\|accesspermissions` | SQLite EQVar diff |
| Upgrade mode | `DCE_config.db3` → `cas\|\|upgrademode` | SQLite EQVar diff |

---

## Areas Not Yet Mapped

- Device/MFP registration and properties
- License seat count changes
- User card enrolment
- Scan-to-email destination setup beyond scan aliases
- Report configuration
- FlexNet license server config changes
