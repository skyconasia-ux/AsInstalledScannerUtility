# Checkpoints

---

## CURRENT CHECKPOINT — 2026-05-06 (Session 9)

### Objective
Fix HTML report summary cards showing blank values after merging `Export-ServerInfo.ps1` Windows data collection into `AsInstalledScanner.ps1`.

### Completed
- Merged `Collect-WindowsData` function (~230 lines) into `AsInstalledScanner.ps1` — all 4 modes now collect Windows server info (system, OS, hardware, network, roles, SQL, print queues)
- Fixed `$null = Write-HtmlReport` (output stream leak corrupting `$afterDir` in Run-Full)
- Fixed EQVar TSV newline bug (multi-line XML broke Compare mode false positives)
- HTML report: Windows sections added (sidebar nav, 10 new sections), formatted content confirmed correct
- Summary cards: removed outer `<div class='sgrid'>` wrapper (was nesting grids, causing blank display)
- Summary cards: replaced `Where-Object` hashtable pipeline with direct string-concat card build
- EQ cards confirmed working (show correct counts); Windows cards still showing empty values

### Current
Windows card values (Platform, OS, RAM, Print Queues) are empty in the HTML summary even though:
- `Write-FullTxt` correctly outputs Platform, OS, RAM from `$data.WinData`
- HTML body sections (System Identity etc.) correctly display all Windows data
- EQ cards display correct counts
- Root cause not yet confirmed — suspected `Collect-WindowsData` output stream leak making `$winData` an array, so `$wd.Platform` via member enumeration returns wrong type to `HE`

### Next Step
SSH back into CSTEMP (was offline — connection timed out) and run diagnostic:
```powershell
ssh Administrator@192.168.60.150 "powershell -NoProfile -ExecutionPolicy Bypass -Command \". 'C:\Temp\AsInstalledScanner\tools\AsInstalledScanner.ps1'; \$d = Collect-Data; Write-Host ('WinData type: ' + \$d.WinData.GetType().FullName); Write-Host ('Count: ' + @(\$d.WinData).Count); Write-Host ('Platform: ' + \$d.WinData.Platform)\""
```
If `$winData` is an array → suppress pipeline leaks in `Collect-WindowsData` with `$null = ...` or `| Out-Null`.
Then redeploy, regenerate, verify cards show values.

### Pending
- Fix Windows summary card empty values (root cause: confirm & fix)
- Commit + push fix to `skyconasia-ux/AsInstalledScannerUtility`
- Add `cas||workflowfolderslastupdatetime` to `$EQVarNoise` (false-positive suppression)
- Continue UI change mapping: device registration, license seats, user card enrolment, report config

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
