# Checkpoints

---

## CURRENT CHECKPOINT — 2026-05-07 (Sessions 14-15)

### Completed
- **Fixed Print Behavior stale value** - `cas||casdownaction` reads from DRE (authoritative), not DCE cache.
- **Redesigned Authentication section** - Device Clients (nested sub-rows), Workstation Clients and Web Client, Equitrac Authentication (checkbox flags), Card Registration.
- **Added Global Network Settings section** - 4 cards: Backward Compatibility, Domain Qualification, SMTP Mail Server (summary + expandable per-server), SNMP Configuration (summary + expandable per-config, v1/v2c vs v3, passwords masked).
- **Fixed SNMP polling interval** - key is `dre||dmepolltime` (seconds / 60 = minutes).
- **Fixed SNMP protocol labels** - AuthProtocol 1=MD5/2=SHA1, PrivProtocol 1=AES-128/2=DES.
- **Fixed Domain Qualification default domain** - key is `cas||domainqualif`, not `dce||defaultdomainqualif`.
- **Removed legacy SMTP/Email section** - superseded by Global Network Settings SMTP card.
- **Added Directory Services Synchronization section** - Active Directory card with per-server collapsible panel (DC, partition, masked auth, container-level flags, per-container Filtering/Field Mappings/Sync).
- **Added LDAP section** - per-server cards from `cas_config.LDAPSettingsDoc` via `Get-CasConfigAttr`; shows connection info (server, port, base DN, login ID, masked password, LDAP version, SSL), Filtering, Field Mappings (incl. AccountNameAlt/DisplayNameAlt/DisplayNameAlt2), Synchronization.
- **Added Microsoft Entra ID section** - from `cas_config.AzureADSettingsDoc`; blue header with Configured/Unknown status badge, last import time, enforce-limits flag, Differential Import (adds/changes/deletes, auto sync, interval, sync on save), Full Import, Field Mappings.
- **Promoted all HTML helpers to script scope** - AG, NG, ARow, ARowH, ARS1, ARS2, SecLbl, ATable, DCOn, DCEnI, YNo, CBx, AVMap, AVMapH, NRow, NTable, NRS1, NOn, NCB, NMask now available to all Build-* functions.
- **Fixed em-dash encoding bug** - replaced U+2014 with ASCII hyphen throughout to prevent Windows PowerShell 5.1 misreading UTF-8 bytes as CP1252 closing-quote.
- **Improved Users/Accounts section** - summary tiles for Total Accounts/Billing Codes/Departments; PIN usage counts (Primary, Alternate, Any); Secondary PIN marked "Unable to verify" (sentinel hash `3CE59CD2B1F5525CFB84E3B1C10F8942` on all DB rows makes it unreliable); 10-row sample user table with raw HTML to avoid double-encoding.
- **Fixed BCP null-byte corruption** - `Get-BcpLines` now reads file as ASCII and strips `\x00` bytes globally; affected `primarypin` NULL fields returning `"\0"` (non-empty) instead of empty string.
- **Improved Installed Components section** - added SystemName (f[1] from `cas_installedsoftware`, was ignored); renamed Desc->ExtraInfo, Date->LastUsed; HTML groups rows by system under dark blue headers; FULL.txt groups under `=== SystemName ===`; metadata JSON includes `systemName`/`lastUsed`. Tested: CSTEMP (8 comps) and DC02-MAIN (6 comps).
- Commits: `28c761f`, `620b728`, `feac9c9`, `4b46557`, `45aac65`, `825dac6`, `8368e85`, `d55026c`, `d54dc5f`, `479de8c`, `38d64ad`.

### Pending
- Add `cas||workflowfolderslastupdatetime` to `$EQVarNoise` (false-positive suppression in Compare mode)
- Continue UI change mapping: device registration, license seats, user card enrolment, report config
- Push to `skyconasia-ux/AsInstalledScannerUtility`
- (Future) Replace raw datetime strings in Last Used / sync timestamps with formatted dates

---

## HISTORY

### 2026-04-18 Session 4–8
- Deployed `AsInstalledScanner.ps1` (4 modes: Before/After/Compare/Full) + `.bat` launcher. All modes tested on CSTEMP.
- Merged `Export-ServerInfo.ps1` Windows data collection into `AsInstalledScanner.ps1`. HTML report now has Windows Server sections (10 sections, sidebar nav).
- Fixed: output stream leak in `Run-Full`, EQVar TSV newline bug breaking Compare mode.
- HTML summary cards: removed nested sgrid, rebuilt EQ cards (working). Windows cards still empty (bug open).

### 2026-04-17 Session 3
- Resolved git merge conflict, pushed full repo to GitHub (`skyconasia-ux/AsInstalledScannerUtility`, commit `1c95884`).
- Confirmed workflow: edit locally → scp → run on VM → scp results back.
- `tools\` folder synced to VM. Script ~1570 lines, syntax OK.

### 2026-04-17 Session 2
- Section 6b added (~300 lines): brand registry dump, services, config files, external DB detection, SQL inspection (cas_config full dump), scheduled tasks, firewall, IIS, event log.
- SSH warning lines bug fixed (removed from top of PS1).
- results\ and logs\ output folders added. TraceLog added.
- tools\ synced to VM. Folder structure organized.

### 2026-04-17 Session 1
- All 10 PS1 fixes applied and pushed (`979b3b5`). Full docs pushed.
- Registry + DB structure dumps collected from CSTEMP. Key: Equitrac 6.5.2.191 / ControlSuite 1.5.0.2, 7 EQ* services Running, `eqcas` DB on `.\SQLExpress` (144 MB, 96 tables).
- Discovered DB script missed all `cas_*` tables. `collect_cs_config.ps1` written as fix.
