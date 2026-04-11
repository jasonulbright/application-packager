# Application Packager 1.x — Detection Audit

**Date:** 2026-04-10 (updated 2026-04-11)
**Source of truth:** Packr JSON definitions (validated via install/detect/uninstall/clean on CLIENT01)
**Total mismatches found:** 41 out of ~91 comparable apps (45% discrepancy rate)
**Fixed:** 41/41 (100%)
**Note:** 33 packagers archived in v1.2.8. Detection fixes for archived apps (Citrix, CutePDF, FileZilla, Foxit, Greenshot, Obsidian, pgAdmin, Postman, PyCharm, Slack, SSMS, Tableau x3, Teams, TreeSize, VMware Tools, VS 2026 x2, VS Code, WizTree, XenCenter, Zoom) are preserved in the archived scripts.

---

## Summary of Changes

### Priority 1: Is64Bit Wrong (5 apps) — FIXED
Corrected registry view (32-bit vs 64-bit) for detection.

| Script | Fix |
|--------|-----|
| package-citrixworkspacecr.ps1 | Is64Bit true → false |
| package-cutepdfwriter.ps1 | Is64Bit false → true |
| package-mremoteng.ps1 | Is64Bit true → false |
| package-rstudio.ps1 | Is64Bit true → false |
| package-winscp.ps1 | Is64Bit true → false |

### Priority 2: Detection Type Completely Different (15 apps) — FIXED
Rewrote detection blocks to match Packr-validated methods.

| Script | Change |
|--------|--------|
| package-filezilla.ps1 | RegistryKeyValue → File (filezilla.exe, Version/GE) |
| package-git.ps1 | Script → File (git.exe, Version/GE) |
| package-pgadmin4.ps1 | RegistryKeyValue → File (pgAdmin4.exe, Existence) |
| package-pycharm.ps1 | File → RegistryKey (ARP key with version) |
| package-r.ps1 | File (Version/GE) → RegistryKey (ARP key existence) — fixed earlier |
| package-soapui.ps1 | File → RegistryKeyValue (fixed ARP key 5517-2803-0637-4585) |
| package-teams.ps1 | Script (Get-AppxPackage) → File (ms-teams.exe, Existence) |
| package-vim.ps1 | RegistryKeyValue → File (gvim.exe, Existence) |
| package-winscp.ps1 | RegistryKeyValue → File (WinSCP.exe, Version/GE) |
| package-wireshark.ps1 | RegistryKeyValue → File (Wireshark.exe, Version/GE) |
| package-zoom.ps1 | File (%APPDATA%) — kept per-user path, detection type unchanged |
| package-corretto-jdk8-x64.ps1 | RegistryKeyValue → RegistryKey (key existence) |
| package-corretto-jdk8-x86.ps1 | RegistryKeyValue → RegistryKey (key existence) |
| package-tableaudesktop.ps1 | RegistryKeyValue (IsEquals, 32-bit) → RegistryKey (64-bit) |
| package-tableauprep.ps1 | RegistryKeyValue (IsEquals, 32-bit) → RegistryKey (64-bit) |
| package-tableaureader.ps1 | RegistryKeyValue (IsEquals, 32-bit) → RegistryKey (64-bit) |

### Priority 3: Missing Is64Bit (11 apps) — FIXED
Added explicit Is64Bit to detection clauses.

| Script | Is64Bit |
|--------|---------|
| package-firefox.ps1 | true |
| package-greenshot.ps1 | true |
| package-notepadplusplus.ps1 | true |
| package-obsidian.ps1 | true |
| package-postman.ps1 | true |
| package-ssms.ps1 | true |
| package-vmwaretools.ps1 | true |
| package-vs2026.ps1 | true |
| package-vs2026community.ps1 | true |
| package-vscode.ps1 | true |
| package-edge.ps1 | false (both compound clauses) |

### Priority 4: Wrong Operator (4 apps) — FIXED
Changed IsEquals to GreaterEquals for supersedence support.

| Script | Fix |
|--------|-----|
| package-keepass.ps1 | IsEquals → GreaterEquals |
| package-putty.ps1 | IsEquals → GreaterEquals |
| package-vlc.ps1 | IsEquals → GreaterEquals |
| package-winscp.ps1 | IsEquals → GreaterEquals |

### Priority 5: Missing Operator/PropertyType (2 apps) — FIXED
Added PropertyType=Version and Operator=GreaterEquals.

| Script | Fix |
|--------|-----|
| package-libreoffice.ps1 | Added PropertyType + Operator |
| package-powershell7.ps1 | Added PropertyType + Operator |

### Priority 6: Other Differences (4 apps) — FIXED

| Script | Fix |
|--------|-----|
| package-python.ps1 | PropertyType Existence → Version/GreaterEquals |
| package-sysinternals.ps1 | PropertyType DateModified → Existence |
| package-obsidian.ps1 | Kept %LOCALAPPDATA% (per-user install, path change reverted) |
| package-postman.ps1 | Kept %LOCALAPPDATA% (per-user install, path change reverted) |

### Earlier Fixes (5 apps)
CCleaner, WinMerge, XenCenter, WebView2, R — fixed in prior session.

### Post-Audit Corrections

**Regressions reverted (3 apps):**
Obsidian, Postman, and Zoom are per-user installs in 1.x. Packr installs
them system-wide, but 1.x uses different installers/switches that target
%LOCALAPPDATA%/%APPDATA%. Detection paths reverted to match actual 1.x
install locations.

**PropertyType=Version added (7 apps):**
CCleaner, GIMP, KeePass, Malwarebytes, PuTTY, VLC, WebView2 had
RegistryKeyValue detection with Operator=GreaterEquals but no PropertyType.
The common module defaults PropertyType to "String", causing lexicographic
comparison instead of version comparison. Added PropertyType="Version".

---

## Not Fixing (architectural differences, not bugs)

- package-dotnet8.ps1 — compound detection (both arches) vs Packr split. Both valid.
- package-msvcruntimes.ps1 — compound registry vs Packr split file. Both valid.
- package-edge.ps1 — 1.x adds EdgeUpdate fallback path. Defensive, not wrong.
