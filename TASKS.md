# Tasks & Backlog

Status legend: `[x]` done · `[-]` partial / documented limitation · `[ ]` not started

---

## EQ Tools — Equitrac/ControlSuite Config Discovery

### Completed

- [x] **EQ-Snapshot.ps1** — captures SQL row counts, BCP content dumps (9 config tables), registry, SQLite EQVar, file timestamps. Deployed at `C:\Temp\EQ-Snapshot.ps1` on CSTEMP.
- [x] **EQ-Diff.ps1** — compares BEFORE/AFTER snapshots. 4 sections: SQL row counts, SQL BCP content (catches UPDATE), SQLite EQVar, registry, files. Noise filter suppresses log/cache files.
- [x] **Export-EquitracConfig.ps1** — standalone config export. Reads live SQLite EQVar databases + BCP table dumps. Sections: versions, auth, SMTP, job management, quotas, pricing, workflows, pull groups, users, full EQVar dump, registry summary.
- [x] **Noise filter regex fix** — `'\\.log$'` → `'[.]log$'` etc. (double-backslash bug matched `\<anychar>log` not `.log`).
- [x] **BCP content diff** — added to catch UPDATE operations that row-count diff misses. Confirmed: workflow rename, price edits, new price list all detected.
- [x] **SQLite EQVar diff** — reads DCE_config.db3 and DREEQVar.db3 via sqlite3.exe. Shows WAS/NOW for each changed key.
- [x] **Storage map** — `docs/equitrac-storage-map.md` documents where every confirmed setting lives.
- [x] **Project structure** — CLAUDE.md, ARCHITECTURE.md, TASKS.md, docs/, tools/, launchers (.bat), .gitignore all updated.

### Pending

- [ ] **Deploy Export-EquitracConfig.ps1 to server** — scp and test run, verify output file is complete.
- [ ] **Add `cas||workflowfolderslastupdatetime` to noise filter** — this background sync timestamp shows up as a false positive in SQLite EQVar diff. Add to a `$EQVarNoiseKeys` set in EQ-Diff.ps1.
- [ ] **Continue UI change mapping** — areas not yet tested: device registration, license seat changes, user card enrolment, report config, FlexNet license changes. Run BEFORE/AFTER for each.
- [ ] **Export-EquitracConfig.ps1: improve pricing display** — cat_pricelist XML parsing works but rate labels (BW/color) aren't yet labeled by page type. Investigate XML structure further.
- [ ] **Export-EquitracConfig.ps1: known column names** — BCP dumps show pipe-delimited data but column names are unknown (eqcas quirk). Once confirmed, add column labels to the output.
- [ ] **Integrate into Export-ServerInfo.ps1** — add a ControlSuite section that calls the EQ export logic when ControlSuite is detected on the server.

---

---

## Completed

- [x] **Per-queue: Use Application Color** — reads `psk:PageColorManagement` from default PrintTicket XML. Maps to FujiFilm "Use the dmColor specified by the application" toggle. `psk:None` = On (pass-through), `psk:System/Driver/Device` = Off.
- [x] **Per-queue: ICM Method** — reads `dmICMMethod` from Default DevMode binary at offset 188 (DWORD). Replaces unreliable `Win32_Printer.ICMMethod` WMI property which always returns 0 for third-party drivers.
- [x] **Per-queue: Installable Options (physical)** — uses `Get-PrinterProperty -PrinterName` to read `Config:OP_FinisherA/B/C/D`, `Config:DC_FIN_Staple`, `Config:DC_FIN_FreeStaple`, `Config:DC_FIN_4Staple`, `Config:DC_FIN_Punch`, `Config:OP_Punch_2_3`, `Config:OP_Punch_2_4`, `Config:OP_Booklet`, `Config:DC_FIN_BiFold`, `Config:DC_FIN_CZFold`, `Config:DC_OffsetStacking`.
- [x] **Per-queue: Installable Options (virtual)** — detects virtual/software queues by absence of properties from `Get-PrinterProperty`. Outputs `N/A (software/virtual driver - no hardware options)`.
- [x] **FujiFilm device count** — count now matches on `FF *`, `*Apeos*`, `*FujiFilm*`, `*FUJIFILM*`, `*Fuji Xerox*` patterns. Previously missed `FF *` prefix.
- [x] **Per-queue Staple XPath** — extended to match `Finishing` feature name in addition to `Staple`/`staple`.
- [x] **Spool folder location** — reads `DefaultSpoolDirectory` from `HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers`. Shown in `[Print Server]` section header.
- [x] **x64 vs x86 driver distinction** — per-queue `x86 (32-bit) Driver` field now correctly reads from registry environment (`Get-PrinterDriver -PrinterEnvironment 'Windows NT x86'`) instead of unreliable `Get-PrinterDriver` `Environment` property (which is always blank). Installed Drivers list is now split into `--- x64 (64-bit) ---` and `--- x86 (32-bit) ---` subsections.

---

## Known Limitations (documented, not planned to fix)

- [-] **Virtual driver installable options not decodeable** — FF Multi-model Print Driver 2 stores Punch/Staple/Booklet/Passcode settings in proprietary binary blobs (`PrinterData1..N`) under `PrinterDriverData`. There is no public API or documentation to decode these. The driver UI (Printer Properties > Options tab) is the only reliable way to view them. Options listed in AddtotheScanner.txt: Punch, Staple, Staple-Free Staple, Hole Punch holes, C/Z Fold Tray, Booklet Tray, Minimum Passcode Length, Customize User Prompts.

---

## Backlog

### Medium priority

- [ ] **Virtual queue: user prompt / passcode settings** — investigate if `Minimum Passcode Length` and `Customize User Prompts` from the FF Multi-model Print Driver 2 Options tab are stored in a readable location (Equitrac/ControlSuite DB, registry, WMI). May require Equitrac COM API or database query.

- [ ] **Per-queue: Staple default value** — current XPath `//psf:Feature[contains(@name,'Staple') or contains(@name,'staple') or contains(@name,'Finishing')]//psf:Option` returns `Unknown` for most FujiFilm queues because the default PrintTicket doesn't include Staple as a Feature node (it's job-level, not device-level). Consider reading `dmFields` bit `DM_COLLATE` + `DM_COPIES` from DEVMODE to infer.

- [ ] **Per-queue: Offset Stacking default** — same issue as Staple above. Returns `Unknown` if not in PrintTicket default.

- [ ] **ControlSuite version detail** — expand to pull ControlSuite build number, patch level, and license key/seat count from registry or service binary.

- [ ] **PaperCut version detail** — pull version from application registry key or binary.

### Low priority

- [ ] **Kofax Universal Print Driver installable options** — verify if `Get-PrinterProperty` works for Kofax UPD in the same way as FujiFilm Apeos. If so, enable the installable options block for Kofax queues.

- [ ] **Port detail: LPR vs RAW vs WSD** — distinguish port types more precisely. Currently shows IP + TCP port number but not the protocol type (LPR/RAW/WSD/IPP).

- [ ] **AD-published printer attributes** — if a queue is published in AD, pull the additional DS attributes (`driverName`, `portName`, `url`, `location`) from `DsSpooler` registry key.

- [ ] **IPv6 NIC addresses** — currently only captures IPv4. Add IPv6 link-local and global addresses where present.

- [ ] **Certificate / HTTPS port** — for queues using IPPS or HTTPS ports, capture the certificate thumbprint and expiry.

- [ ] **Windows Event Log: spooler errors (last 7 days)** — pull count of Event ID 372 (driver crashes) and 23/24 (print job errors) from the System log to surface reliability issues.

- [ ] **Multi-server support** — currently one invocation = one server. Consider a wrapper that accepts a host list and runs SSH/WinRM collection, aggregating output per server.

---

## AddtotheScanner.txt Items (2026-04-17)

Original request items and their status:

| Item | Status | Notes |
|------|--------|-------|
| Virtual queue installable options (Punch, Staple, FreeStaple, Hole Punch, C/Z Fold, Booklet, Passcode, User Prompts) | Limitation documented | Binary blob in `PrinterDriverData`, no public decode API. See Known Limitations above. |
| Distinguish x64 driver from x86 in per-queue block | Done | Fixed via `Get-PrinterDriver -PrinterEnvironment` |
| Distinguish x64/x86 in installed drivers list | Done | List now has `--- x64 ---` and `--- x86 ---` subsections |
| Spool folder location | Done | Reads `DefaultSpoolDirectory` from registry |
