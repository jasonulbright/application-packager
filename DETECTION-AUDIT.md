# Application Packager 1.x — Detection Audit

**Date:** 2026-04-10
**Source of truth:** Packr JSON definitions (validated via install/detect/uninstall/clean on CLIENT01)
**Total mismatches:** 41 out of ~91 comparable apps (45% discrepancy rate)
**Fixed so far:** 5 (CCleaner, WinMerge, XenCenter, WebView2, R)
**Remaining:** 36

---

## Priority 1: Is64Bit Wrong (5 apps)

Detection will FAIL on 64-bit Windows — checking wrong registry view.

| # | Script | Packr Is64Bit | 1.x Is64Bit | Notes |
|---|--------|--------------|-------------|-------|
| 1 | package-citrixworkspacecr.ps1 | false | true | WOW6432Node key, must be false |
| 2 | package-cutepdfwriter.ps1 | true | false | Should be true per Packr |
| 3 | package-mremoteng.ps1 | false | true | 32-bit MSI, must be false |
| 4 | package-rstudio.ps1 | false | true | WOW6432Node key, must be false |
| 5 | package-winscp.ps1 | false | true | Installs to Program Files (x86) |

## Priority 2: Detection Type Completely Different (15 apps)

1.x uses a different detection method than Packr's validated approach.

| # | Script | Packr Detection | 1.x Detection |
|---|--------|----------------|---------------|
| 1 | package-filezilla.ps1 | File (filezilla.exe, Version/GE) | RegistryKeyValue (IsEquals) |
| 2 | package-git.ps1 | File (git.exe, Version/GE) | Script (PowerShell) |
| 3 | package-pgadmin4.ps1 | File (pgAdmin4.exe, Existence) | RegistryKeyValue (GE) |
| 4 | package-pycharm.ps1 | RegistryKey existence | File existence |
| 5 | package-r.ps1 | RegistryKey existence | File (Version/GE) — already improved |
| 6 | package-soapui.ps1 | RegistryKeyValue (GE) | File (version in path+filename) |
| 7 | package-teams.ps1 | File (ms-teams.exe) | Script (Get-AppxPackage) |
| 8 | package-vim.ps1 | File (gvim.exe, Existence) | RegistryKeyValue (GE) |
| 9 | package-winscp.ps1 | File (WinSCP.exe, Version/GE, x86) | RegistryKeyValue (IsEquals, x64) |
| 10 | package-wireshark.ps1 | File (Wireshark.exe, Version/GE) | RegistryKeyValue (IsEquals) |
| 11 | package-zoom.ps1 | RegistryKeyValue (ProductCode) | File (%APPDATA%, per-user) |
| 12 | package-corretto-jdk8-x64.ps1 | RegistryKey existence | RegistryKeyValue |
| 13 | package-corretto-jdk8-x86.ps1 | RegistryKey existence | RegistryKeyValue |
| 14 | package-tableaudesktop.ps1 | RegistryKey (GE, Is64Bit=true) | RegistryKeyValue (IsEquals, Is64Bit=false) |
| 15 | package-tableauprep.ps1 | RegistryKey (GE, Is64Bit=true) | RegistryKeyValue (IsEquals, Is64Bit=false) |
| 16 | package-tableaureader.ps1 | RegistryKey (GE, Is64Bit=true) | RegistryKeyValue (IsEquals, Is64Bit=false) |

## Priority 3: Missing Is64Bit (11 apps)

| # | Script | Packr Is64Bit |
|---|--------|--------------|
| 1 | package-firefox.ps1 | true |
| 2 | package-greenshot.ps1 | true |
| 3 | package-notepadplusplus.ps1 | true |
| 4 | package-obsidian.ps1 | true |
| 5 | package-postman.ps1 | true |
| 6 | package-ssms.ps1 | true |
| 7 | package-vmwaretools.ps1 | true |
| 8 | package-vs2026.ps1 | true |
| 9 | package-vs2026community.ps1 | true |
| 10 | package-vscode.ps1 | true |
| 11 | package-edge.ps1 | false |

## Priority 4: Wrong Operator (4 apps)

Using IsEquals instead of GreaterEquals — breaks supersedence detection.

| # | Script | Packr Operator | 1.x Operator |
|---|--------|---------------|-------------|
| 1 | package-keepass.ps1 | GreaterEquals | IsEquals |
| 2 | package-putty.ps1 | GreaterEquals | IsEquals |
| 3 | package-vlc.ps1 | GreaterEquals | IsEquals |
| 4 | package-winscp.ps1 | GreaterEquals | IsEquals |

## Priority 5: Missing Operator/PropertyType (2 apps)

| # | Script | Missing Fields |
|---|--------|---------------|
| 1 | package-libreoffice.ps1 | Operator, PropertyType |
| 2 | package-powershell7.ps1 | Operator, PropertyType |

## Priority 6: Other Differences (4 apps)

| # | Script | Issue |
|---|--------|-------|
| 1 | package-python.ps1 | PropertyType Existence → should be Version |
| 2 | package-sysinternals.ps1 | PropertyType DateModified → should be Existence |
| 3 | package-obsidian.ps1 | Path %LOCALAPPDATA% → should be Program Files |
| 4 | package-postman.ps1 | Path %LOCALAPPDATA% → should be Program Files |

## Not Fixing (architectural differences, not bugs)

- package-dotnet8.ps1 — compound detection (both arches) vs Packr split. Both valid.
- package-msvcruntimes.ps1 — compound registry vs Packr split file. Both valid.
- package-edge.ps1 — 1.x adds EdgeUpdate fallback path. Defensive, not wrong.
