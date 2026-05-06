# Checkpoints

---

## CURRENT CHECKPOINT — 2026-05-06 (Session 10)

### Completed
- **Fixed Windows summary cards** — root cause: `HE (if ($wd) {...})` pattern fails because `if` is
  treated as a command name when PowerShell evaluates the expression in certain contexts (e.g. SSH).
  Fix: replaced with inline `$(HE $data.WinData.Platform)` subexpression pattern matching EQ cards.
  Verified on CSTEMP: Platform=Virtual Machine, OS=Microsoft WS 2022 Standard Evaluation, RAM=8 GB, Queues=5.

### Pending
- Add `cas||workflowfolderslastupdatetime` to `$EQVarNoise` (false-positive suppression in Compare mode)
- Continue UI change mapping: device registration, license seats, user card enrolment, report config
- Push to `skyconasia-ux/AsInstalledScannerUtility`

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
