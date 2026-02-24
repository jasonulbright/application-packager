# Changelog

All notable changes to AppPackager are documented in this file.

## [1.0] - 2026-02-24

### Added
- **Vendor URL debug column** — new "Vendor URL (Debug)" column in the DataGridView, visible when the Debug Columns checkbox is checked; shows the vendor's product page URL for each packager
- **Ctrl+Click to open vendor page** — Ctrl+Left-click any non-checkbox cell to open the vendor's product page in the default browser; URL sourced from the new `VendorUrl:` metadata tag in each packager script header
- **Row hover tooltips** — hovering over any grid row displays the packager's `.SYNOPSIS` description from the script header; uses DataGridView's built-in `CellToolTipTextNeeded` event (replaces the former static grid tooltip)
- `VendorUrl:` metadata tag added to all 62 packager script headers — parsed by `Get-PackagerMetadata` alongside existing `Vendor:`, `App:`, and `CMName:` tags
- `Description` field added to `Get-PackagerMetadata` — parses the first non-blank line after `.SYNOPSIS` in the script's comment-based help block

### Changed
- `Invoke-PackagerGetLatestVersion` rewritten with async I/O — replaced synchronous `ReadToEnd()` calls with `ReadToEndAsync()` tasks and bounded `WaitForExit(30000)` timeout; prevents GUI hang when grandchild processes (e.g., `curl.exe`, `expand.exe` in M365 scripts) inherit stdout/stderr pipe handles
- Debug Columns checkbox tooltip updated to reflect new VendorURL column: "Show or hide the CMName, Script, and Vendor URL debug columns"
- Removed static grid tooltip ("Select packagers using the checkboxes...") — conflicted with per-cell tooltip system
- GUI version bumped from 0.4.0 to 1.0.0

### Fixed
- GUI "Check Latest" hang — `Invoke-PackagerGetLatestVersion` used synchronous `ReadToEnd()` which blocked indefinitely when M365 ODT packagers spawned `expand.exe` / `curl.exe` as grandchild processes that inherited the stdout pipe handle; rewritten to use `ReadToEndAsync()` with a 30-second process timeout and 5-second task timeout

---

## [0.0.9] - 2026-02-23 6:45 PM

### Added
- Preferences dialog (`File > Preferences`) — modal dialog with 6 fields (Site Code, File Share Root, Download Root, Est. Runtime, Max Runtime, Company Name); persisted to `AppPackager.preferences.json` alongside the GUI script; validates numeric fields with `[int]::TryParse` fallback; Company Name syncs bidirectionally with `packager-preferences.json` (seeded from existing value on first launch, written back on save)
- MenuStrip with `File` menu — "Preferences..." (`Ctrl+,`) opens the Preferences dialog; "Exit" closes the application
- Tooltips on all interactive controls — Comment field, log pane, DataGridView, Debug Columns checkbox, Select Update Available button, and all 4 action buttons; `AutoPopDelay = 10000`, `InitialDelay = 400`

- 6 M365 ODT-based packager scripts — M365 Apps for Enterprise (x64/x86), M365 Visio (x64/x86), M365 Project (x64/x86); each downloads the Office Deployment Tool, queries the Semi-Annual Enterprise Channel CDN for the current version, downloads offline Office source files, generates install/uninstall XML configs and content wrappers, and writes a stage manifest with file-version detection (WINWORD.EXE, VISIO.EXE, WINPROJ.EXE respectively)
- `package-vs2026.ps1` — Visual Studio 2026 Enterprise offline layout packager; downloads the VS bootstrapper, creates an offline layout with configurable workloads (`$LayoutArgs`), generates interactive install wrapper (no `--quiet`, uses `--noWeb` for offline); MECM deployment type uses `RequireUserInteraction` and `OnlyWhenUserLoggedOn` so the VS Installer UI is visible to logged-on users
- `package-ssms.ps1` — SQL Server Management Studio packager; downloads the SSMS bootstrapper (VS Installer backend), creates an offline layout with `--layout`, generates silent install/uninstall wrappers; file-version detection via Ssms.exe
- Optional manifest override fields in `New-MECMApplicationFromManifest` — `LogonRequirementType` and `RequireUserInteraction` can now be set per-manifest to override the default `WhetherOrNotUserLoggedOn` behavior (used by VS2026 for interactive deployment)
- `Get-PackagerPreferences` in `AppPackagerCommon.psm1` — reads `packager-preferences.json` from the Packagers folder for universal settings (e.g., `CompanyName` used in ODT config XML)
- `New-OdtConfigXml` in `AppPackagerCommon.psm1` — generates full ODT configuration XML for download and install phases; supports parameterized `OfficeClientEdition`, `Version`, `ProductIds` array, optional `SourcePath` (download only), and `CompanyName` from preferences; produces complete XML matching production template with `ExcludeApp`, `SharedComputerLicensing`, `FORCEAPPSHUTDOWN`, `MigrateArch`, `RemoveMSI`, `AppSettings`, `Display`, and `Logging` elements
- `packager-preferences.json` in `Packagers/` — user-editable JSON file for universal packager settings; currently contains `CompanyName` for ODT `AppSettings/Company` value
- `AppPackagerCommon.Tests.ps1` — Pester 5.x test suite for the shared module (70 tests); covers `Write-Log`, `Initialize-Logging`, `New-MsiWrapperContent`, `New-ExeWrapperContent`, `Write-ContentWrappers`, `Write-StageManifest`/`Read-StageManifest` round-trip (including per-user deployment overrides and fixed ARP key detection), `New-OdtConfigXml` (structure, attributes, multi-product, CompanyName/AppSettings, SourcePath presence/absence), `Initialize-Folder`, and `Get-PackagerPreferences`; uses Pester `$TestDrive` for all file I/O — no MECM, network, or admin elevation required
- `package-vlc.ps1` — VLC Media Player (x64) MSI packager; scrapes the VideoLAN directory listing at `download.videolan.org/vlc/last/win64/` for the latest version, downloads the x64 MSI, derives detection from the fixed ARP key `VLC media player` (ProductCode is auto-generated per build); standard MSI install/uninstall wrappers via `New-MsiWrapperContent`
- `package-zoom.ps1` — Zoom Workplace (x64) per-user EXE packager; resolves version from the redirect URL at `zoom.us/client/latest/ZoomInstaller.exe`, downloads the per-user EXE installer and CleanZoom utility (extracted from ZIP); file-existence detection at `%APPDATA%\Zoom\bin\Zoom.exe`; first per-user packager in the project — uses `InstallationBehaviorType: InstallForUser` and `LogonRequirementType: OnlyWhenUserLoggedOn` manifest overrides
- `InstallationBehaviorType` manifest override in `New-MECMApplicationFromManifest` — allows per-manifest override of the default `InstallForSystem` behavior (used by Zoom Workplace for per-user deployment)
- `package-r.ps1` — R for Windows (x64) EXE packager; queries the r-hub versions API (`api.r-hub.io/rversions/r-release-win`) for the latest release, downloads the Inno Setup EXE installer from CRAN; file-existence detection at `C:\Program Files\R\R-{VERSION}\bin\R.exe` (version-specific install path)
- `package-rstudio.ps1` — RStudio Desktop (x64) EXE packager; queries the GitHub tags API for the latest `rstudio/rstudio` tag, constructs the download URL for the Posit CDN (`download1.rstudio.org/electron/windows/`); ARP registry detection via fixed key `RStudio` with `DisplayVersion`
- `package-positron.ps1` — Positron IDE (x64) EXE packager; queries the GitHub releases API for `posit-dev/positron`, downloads the system-level InnoSetup installer from the Posit CDN (`cdn.posit.co/positron/releases/`); file-existence detection at `C:\Program Files\Positron\Positron.exe`
- `package-python.ps1` — Python (x64) EXE packager; queries the endoflife.date API for the latest stable Python release, downloads the official installer from `python.org/ftp/python/`; file-existence detection at version-specific install path (`C:\Program Files\Python3XX\python.exe`); manual install/uninstall wrappers (same EXE with `/uninstall` flag)
- `package-anaconda.ps1` — Anaconda Distribution (x64) EXE packager; scrapes the Anaconda repository archive (`repo.anaconda.com/archive/`) for the latest Windows x64 installer (~1 GB); file-existence detection at `C:\ProgramData\anaconda3\python.exe`; manual install wrappers with NSIS `/D=` last requirement
- 14 Eclipse Temurin packager scripts (7 JRE + 7 JDK) — MSI-based Java packagers for LTS versions 8, 11, 17, 21, 25; versions 8 and 11 include both x64 and x86 variants; queries the Adoptium API (`api.adoptium.net/v3/assets/latest/`) for the latest MSI release; ARP registry detection derived from MSI ProductCode; `staged-version.txt` pattern for Package phase version resolution
- 7 Amazon Corretto packager scripts (JDK only) — MSI-based Java packagers for LTS versions 8, 11, 17, 21, 25; versions 8 and 11 include both x64 and x86 variants; queries the GitHub releases API for tag version, constructs download URL from known Corretto CDN pattern (`corretto.aws/downloads/resources/`); normalizes 5-part Corretto version to 4-part for GUI compatibility; ARP registry detection derived from MSI ProductCode
- `Get-LatestTemurinRelease` in `AppPackagerCommon.psm1` — queries the Adoptium API for the latest Eclipse Temurin MSI release; supports `-FeatureVersion` (8/11/17/21/25), `-ImageType` (jre/jdk), `-Architecture` (x64/x86); strips `-LTS` suffix from version; returns hashtable with Version, DownloadUrl, FileName
- `Get-LatestCorrettoRelease` in `AppPackagerCommon.psm1` — queries the GitHub releases API for the latest Amazon Corretto release tag; constructs MSI download URL from known CDN pattern (v8 has `-jdk` filename suffix, v11+ does not); normalizes 5-part version to 4-part; returns hashtable with Version, DownloadUrl, FileName
- `AppPackagerCommon.Tests.ps1` — 3 additional Pester tests for Java manifest round-trips: Temurin JRE 8 (version with `+build` suffix), Temurin JDK 21 (standard ARP detection), Corretto JDK 21 (4-part normalized version); test count 67 → 70

### Changed
- Settings that rarely change (Site Code, File Share Root, Download Root, Est/Max Runtime) moved from inline main form fields to the Preferences dialog — reduces main form clutter from 3 configuration rows to a single Comment row
- Removed the Load/Save named configuration system (`AppPackager.configurations.json`, `Get-ConfigStorePath`, `Read-ConfigStore`, `Write-ConfigStore`, `Set-ConfigurationInputs`, `Save-Configuration`, `Update-ConfigDropdown`) — replaced by the simpler single-preferences model
- Removed 18 controls from the main form (5 settings labels, 5 settings textboxes, 2 "mins" suffix labels, 2 config labels, 1 config combobox, 1 config textbox, 1 save button, 1 info disclaimer label)
- "WO / Comment" label renamed to "Comment"
- Log pane relocated from bottom of window (3 lines) to upper-right beside the logo (~12 visible lines); frees vertical space for the DataGridView
- Removed "No network or MECM actions occur..." info disclaimer label — redundant with tooltips on every action button
- All button and context menu handlers now read from `$script:Prefs` instead of textbox controls; Comment remains on the main form as `$txtComment`
- Simplified `Set-UILayout` — Comment field spans top row below the menu; logo and log share the second row; DataGridView fills remaining space down to the action buttons
- GUI version bumped from 0.3.0 to 0.4.0; `MinimumSize` height reduced from 640 to 560
- `-SiteCode` parameter preserved for backward compatibility — overrides the loaded preference for the session when specified on the command line
- VS2026 `$LayoutArgs` changed from 4 specific workloads to `--all` for a full offline layout (~35-50 GB)
- M365 ODT download URL changed from CDN channel-specific `setup.exe` (3.8 MB Click-to-Run client, not the ODT) to `https://officecdn.microsoft.com/pr/wsus/setup.exe` (7.1 MB, real Office Deployment Tool)
- M365 version detection changed from `Build` attribute in `VersionDescriptor.xml` to `I640Version` (x64) / `I320Version` (x86) — `Build` is a Windows 7 fallback version, not the current SAEC version
- SSMS bootstrapper URL changed from `SSMS-Setup-ENU.exe` to `vs_SSMS.exe` (`https://aka.ms/ssms/22/release/vs_SSMS.exe`) and version detection changed from `FileVersion` to `ProductVersion` — `FileVersion` reports the VS Installer engine version, not the SSMS version

### Fixed
- M365 ODT install.xml — removed `SourcePath="."` attribute from `<Add>` element; ODT silently exits with exit code 0 when given a relative source path; omitting `SourcePath` causes ODT to default to the XML file's directory, matching the production template
- GUI `Invoke-ProcessWithStreaming` — replaced unbounded `$p.WaitForExit()` with `$p.WaitForExit(15000)` timeout; installers like SSMS and VS2026 spawn grandchild processes that inherit stdout/stderr handles, causing `WaitForExit()` to block indefinitely even after the main process exits and all output is consumed; also guards `$errTask.Result` access to prevent blocking on inherited stderr handles
- GUI version string validation — widened regex to `^\d+(\.\d+){1,3}([+-]\d+)?$` to accept both `+build` and `-build` suffixes (e.g., RStudio `2026.01.1+403`, Positron `2026.02.1-5`, Anaconda `2025.12-2`); `Compare-SemVer` strips both `+` and `-` build suffixes before `[version]` parsing

---

## [0.0.8] - 2026-02-23

### Added
- `AppPackagerCommon.psm1` + `AppPackagerCommon.psd1` — promoted shared helpers from dot-sourced `.ps1` to a proper PowerShell module with manifest (version 0.0.8, explicit `FunctionsToExport`)
- `Write-ContentWrappers` in `AppPackagerCommon.psm1` — universal function that creates install/uninstall .bat and .ps1 wrapper files in a content folder; skips existing files; writes with `-Encoding ASCII` (no BOM)
- `New-MsiWrapperContent` in `AppPackagerCommon.psm1` — returns install and uninstall .ps1 content strings for MSI products using array-based `Start-Process -ArgumentList`
- `New-ExeWrapperContent` in `AppPackagerCommon.psm1` — returns install and uninstall .ps1 content strings for EXE products; accepts installer filename, install args, uninstall command, and optional uninstall args
- `Get-NetworkAppRoot` in `AppPackagerCommon.psm1` — creates and returns the `<FileServerPath>\Applications\<VendorFolder>\<AppFolder>` path, initializing each directory level; replaces 25 identical per-script `Get-<App>NetworkAppRoot` functions
- `New-MECMApplicationFromManifest` in `AppPackagerCommon.psm1` — creates a CM Application with Script deployment type from a stage manifest; handles site connection, duplicate check, `New-CMApplication` with `-AutoInstall $true`, `Add-CMScriptDeploymentType` with gold standard parameters (`-ContentFallback`, `-SlowNetworkDeploymentMode Download`), and revision history cleanup
- `New-SingleDetectionClause` (private helper) in `AppPackagerCommon.psm1` — builds a single CM detection clause object from a manifest detection block; supports RegistryKeyValue, RegistryKey, and File types via splatted parameters
- `-DownloadRoot` parameter on all 26 packager scripts — configurable local staging root (default `C:\temp\ap`) replaces the former hardcoded `$env:USERPROFILE\Downloads\_AutoPackager\<App>` path
- `-EstimatedRuntimeMins` parameter on all 26 packager scripts — MECM deployment type estimated runtime (default 15 minutes)
- `-MaximumRuntimeMins` parameter on all 26 packager scripts — MECM deployment type maximum runtime (default 30 minutes)
- `-StageOnly` and `-PackageOnly` switches on all 26 packager scripts — two-phase workflow separating download/metadata discovery from MECM application creation
- `Invoke-Stage<App>` and `Invoke-Package<App>` functions on all 26 packager scripts — encapsulate each phase's logic
- `Find-UninstallEntry` in `AppPackagerCommon.ps1` — ARP registry discovery function that searches both `Uninstall` and `WOW6432Node\Uninstall` paths by DisplayName pattern, returns registry key path, DisplayVersion, Publisher, and Is64Bit flag
- `Write-StageManifest` / `Read-StageManifest` in `AppPackagerCommon.ps1` — JSON manifest helpers for stage-manifest.json with automatic SchemaVersion and StagedAt timestamp
- GUI "Download Root" field — sets the local staging root passed to all packager invocations (Check Latest, Stage, and Package)
- GUI "Est. Runtime" and "Max Runtime" fields (with "mins" suffix labels) — override the MECM deployment type runtime values passed during Package
- GUI "Stage Packages" button (orange) — runs selected scripts with `-StageOnly` to download installers, discover ARP metadata, and write stage manifests
- GUI "Package Apps" button (purple) — runs selected scripts with `-PackageOnly` to read manifests, copy content to network, and create MECM applications
- GUI right-click context menu on grid rows — "Open Log Folder" (opens `Logs/` in Explorer), "Open Staged Folder" (resolves app-specific download subfolder from script, opens versioned folder if available), "Open Network Share" (resolves `VendorFolder/AppFolder` from script, opens UNC path), "Copy Latest Version" (clipboard); uses `Get-PackagerFolderInfo` helper to parse `$BaseDownloadRoot`, `$VendorFolder`, `$AppFolder` from packager script headers
- GUI "Select Update Available" button — flat-styled button next to the "Show Debug Columns" checkbox; calls existing `Select-OnlyUpdateAvailable` to auto-check only rows with status "Update available"
- GUI window size/position persistence — saves `Left`, `Top`, `Width`, `Height`, and `Maximized` state to `AppPackager.windowstate.json` on form close; restores on next launch with screen bounds validation (ensures title bar remains visible, respects minimum size); uses `RestoreBounds` when maximized to preserve normal-size dimensions
- GUI real-time log streaming — `Invoke-ProcessWithStreaming` helper replaces synchronous `ReadToEnd()` with `ReadLineAsync()` polling loop; streams packager script output line-by-line into the log pane during Stage and Package operations; strips `Write-Log` timestamp/severity prefix for display; uses `DoEvents()` for UI responsiveness with 50ms poll interval; stderr collected asynchronously via `ReadToEndAsync()`
- `.PARAMETER` documentation for `DownloadRoot`, `EstimatedRuntimeMins`, `MaximumRuntimeMins`, `StageOnly`, and `PackageOnly` on the gold standard (`package-7zip.ps1`)

### Changed
- All 26 packager scripts now use `Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force` instead of dot-sourcing `AppPackagerCommon.ps1` — `-Force` ensures fresh import each run
- Gold standard (`package-7zip.ps1`) Stage phase uses `New-MsiWrapperContent` + `Write-ContentWrappers` instead of ~40 lines of inline wrapper generation
- Gold standard (`package-7zip.ps1`) Package phase uses `New-MECMApplicationFromManifest` instead of ~65 lines of inline MECM creation code
- Gold standard (`package-7zip.ps1`) Stage manifest now includes `Detection.Type = "RegistryKeyValue"` for explicit type discrimination (backward compatible - missing Type defaults to RegistryKeyValue)
- `New-MECMApplicationFromManifest` extended to support all 5 detection types: RegistryKeyValue (single registry value comparison), RegistryKey (existence), File (existence or version), Script (PowerShell script text), Compound (multiple clauses with AND/OR connectors); uses splatted parameters for `Add-CMScriptDeploymentType`; supports optional `PostExecutionBehavior` from manifest (e.g., `ForceReboot` for MSVC Runtimes)
- `Write-ContentWrappers` extended with optional `-InstallBatExitCode` and `-UninstallBatExitCode` parameters (default `%ERRORLEVEL%`); enables VMware Tools to use `exit /b 3010` for forced reboot
- Consolidated 5 identical boilerplate functions (`Test-IsAdmin`, `Connect-CMSite`, `Initialize-Folder`, `Test-NetworkShareAccess`, `Remove-CMApplicationRevisionHistoryByCIId`) and TLS 1.2 initialization into `AppPackagerCommon.ps1` — removed 130 duplicate function definitions + 26 TLS lines (~83 KB) from packager scripts
- Consolidated `Get-MsiPropertyMap` into `AppPackagerCommon.ps1` — standardized on 4-property version (ProductName, ProductVersion, Manufacturer, ProductCode); removed 5 duplicate definitions (~7.7 KB) from MSI-based packager scripts (7-Zip, Chrome, MSODBCSQL18, MSOleDb, Webex)
- Gold standard (`package-7zip.ps1`) restructured into `Invoke-Stage7Zip` and `Invoke-Package7Zip` functions — Stage downloads MSI, derives ARP detection from MSI properties (ProductCode becomes registry key, ProductVersion becomes DisplayVersion), generates content wrappers, writes `stage-manifest.json`; Package reads manifest, copies to network, creates MECM app with registry-based detection (`New-CMDetectionClauseRegistryKeyValue` replaces `New-CMDetectionClauseWindowsInstaller`)
- Remaining 25 scripts restructured with `Invoke-Stage<App>` / `Invoke-Package<App>` functions; per-script `New-MECM<App>Application` and `Get-<App>NetworkAppRoot` functions removed from all 25 scripts
- **Batch 1 (MSI products):** Chrome, MSODBCSQL18, MSOleDb, Webex — ARP detection derived from MSI properties via `Get-MsiPropertyMap` (no temp install); registry key = `SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{ProductCode}`
- **Batch 2 (Simple EXE, known paths):** Dotnet9 x64, Dotnet10 x64, VSCode, WebView2 — file-based detection with version or existence checks at known install paths
- **Batch 3 (Simple EXE, remaining):** ASPNETHostingBundle8 (RegistryKey existence), VMwareTools (File version, custom `exit /b 3010`), Greenshot (File existence, process kill before install/uninstall)
- **Batch 4 (Compound detection):** Dotnet8 (AND, 2x File existence for x86+x64), MSVC Runtimes (AND, 2x RegistryKeyValue, `PostExecutionBehavior: ForceReboot`)
- **Batch 5 (MSI-wrapped + EXE):** Edge (OR compound, 2x File version for x86+x64 paths), Firefox (File version GreaterEquals), Notepad++ (File version GreaterEquals), PowerBI Desktop (File version GreaterEquals)
- **Batch 6 (Script detection):** Git (Inno Setup, `detection.ps1` with `git.exe --version` parsing), Teams (MSIX bootstrapper, `detection.ps1` with `Get-AppxPackage`)
- **Batch 7 (Temp install products):** AdobeReader (File version for `Acrobat.exe`, reads FileVersionInfo from EXE without temp install), WinSCP (RegistryKeyValue with Is64Bit, temp install for ARP discovery), Wireshark (RegistryKeyValue, temp install with 20-retry polling at 60s intervals), Tableau Desktop/Prep/Reader (File version, temp install for install location + file version discovery)
- WinSCP download URL changed from broken CDN resolution (`winscp.net/eng/downloads.php` page structure changed) to SourceForge direct URL pattern (`sourceforge.net/projects/winscp/files/WinSCP/{version}/WinSCP-{version}-Setup.exe/download`); requires `curl -L` for 302 redirects to mirror
- Local staging now versioned: `C:\temp\ap\7-Zip\25.01\` (with installer + wrappers + manifest) instead of flat `C:\temp\ap\7-Zip\`
- All 26 packager scripts compute `$BaseDownloadRoot = Join-Path $DownloadRoot "<AppSubfolder>"` instead of `Join-Path $env:USERPROFILE "Downloads\_AutoPackager\<AppSubfolder>"`
- Removed all hardcoded `$EstimatedRuntimeMins` and `$MaximumRuntimeMins` variable assignments from packager scripts — values now flow exclusively from parameters
- GUI `Invoke-PackagerGetLatestVersion` now passes `-DownloadRoot` to packager scripts
- GUI replaced `Invoke-PackagerRun` / "Run Selected Packagers" with `Invoke-PackagerStage` / `Invoke-PackagerPackage` and corresponding "Stage Packages" / "Package Apps" buttons
- Configuration save/load (`Save-Configuration`, `Set-ConfigurationInputs`, `AppPackager.configurations.json`) updated to persist `DownloadRoot`, `EstimatedRuntimeMins`, and `MaximumRuntimeMins`
- Replaced the 5-line description panel with a single condensed italic info line to accommodate the additional configuration fields without increasing window size
- GUI layout reorganized: action bar now has 4 buttons instead of 3; button width calculation adjusted from `/3` to `/4`
- GUI version bumped from 0.1.0 to 0.3.0
- GUI layout updated to position "Select Update Available" button adjacent to debug columns checkbox via `Set-UILayout`

### Fixed
- `Write-Log` now accepts empty strings via `[AllowEmptyString()]` attribute — `Write-Log ""` for blank log lines was rejected by PowerShell's mandatory parameter validation
- MSI-based Stage phase (7-Zip) derives ARP detection from MSI properties (`ProductCode` → registry key, `ProductVersion` → `DisplayVersion`) instead of temp install/uninstall — temporary installation of products with shell extensions (e.g., 7-Zip context menu) crashed `explorer.exe` when the extension DLL was unregistered while still loaded
- Content wrapper `.ps1` files use proper `-ArgumentList` array (`@('/i', "`"$msiPath`"", '/qn', '/norestart')`) instead of a single string — `Start-Process` misparses a single string with embedded quotes, causing msiexec to receive no arguments
- Content wrapper `.ps1` files written with `-Encoding ASCII` instead of `-Encoding UTF8` — PowerShell 5.1's `UTF8` encoding always prepends a BOM (`EF BB BF`), which appeared as `∩╗┐` at position 0
- Stage manifest `Detection.DisplayVersion` now stores the raw MSI `ProductVersion` (e.g., `26.00.00.0`) instead of the normalized display version (`26.00`) — the MECM `IsEquals` detection clause must match the actual 4-part registry value

### Changed
- Form icon loading now prefers `apppackager.ico` (multi-resolution `.ico` loaded directly via `System.Drawing.Icon`) over the previous JPG-to-bitmap-to-`GetHicon()` conversion; falls back to JPG/PNG bitmap, then `SystemIcons.Application`

### New Packager Scripts
- `package-putty.ps1` — PuTTY (x64) MSI from `the.earth.li/~sgtatham/putty/latest/w64/`; version scraped from directory listing; ARP detection derived from MSI properties (ProductCode registry key, ProductVersion as DisplayVersion); `New-MsiWrapperContent` + `Write-ContentWrappers`
- `package-filezilla.ps1` — FileZilla Client (x64) NSIS EXE from `download.filezilla-project.org`; version extracted from HTML meta description tag on download page (no AES decryption needed); requires browser-like User-Agent header; fixed ARP registry key `FileZilla Client` (not GUID-based); custom install/uninstall wrappers with NSIS `/S` switch
- `package-keepass.ps1` — KeePass 2.x MSI from SourceForge; version from `keepass.info/update/version2x.txt`; ARP detection derived from MSI properties; 32-bit .NET app (`Is64Bit = $false` for WOW6432Node registry); SourceForge download requires `-L --max-redirs 10` for mirror redirects

### Deprecated
- **Tableau Desktop/Prep/Reader:** Deprecated and moved to `Archive/deprecated/`. Salesforce removed public download access at `downloads.tableau.com/tssoftware/` (returns 404 for all versions). The original pre-refactored scripts used the identical URL pattern. Scripts were structurally complete and passing all verification checks but could not stage due to the upstream change.

### Staging Validation
26 of 26 active products staged and validated successfully.

---

## [0.0.7] - 2026-02-22 4:08 PM

### Added
- `Packagers\AppPackagerCommon.ps1` — shared dot-sourceable helper with structured logging and download retry functions
  - `Write-Log` — timestamped, severity-tagged output (`[yyyy-MM-dd HH:mm:ss] [INFO ] message`); ERROR level also writes to stderr via `$host.UI.WriteErrorLine`; `-Quiet` switch suppresses console output while still writing to the log file
  - `Initialize-Logging` — accepts `-LogPath` to enable file-based logging alongside console output
  - `Invoke-DownloadWithRetry` — wraps `curl.exe` file downloads with 1 retry and 5-second delay; `-ExtraCurlArgs` for per-script flags (`-A "PowerShell"`, `--max-redirs 10`); `-Quiet` for silent version-check mode

### Changed
- All 26 packager scripts now dot-source `AppPackagerCommon.ps1` and use `Write-Log` exclusively (zero remaining `Write-Host`/`Write-Error`/`Write-Warning` calls)
- All 26 packager scripts accept a new `-LogPath` parameter for structured log file output
- All `curl.exe -o` file-download calls replaced with `Invoke-DownloadWithRetry` (transient failures now retry once before failing)
- Simplified `-Quiet` guards in version-resolution functions from `if (-not $Quiet) { Write-Host }` to `Write-Log -Quiet:$Quiet`
- GUI `Invoke-PackagerRun` now passes `-LogPath` to packager scripts, generating a `.structured.log` alongside `.out.log`/`.err.log`
- GUI failure handler now surfaces up to 10 stderr lines directly in the log pane instead of a single-line summary referencing the `.err.log` file

---

## [0.0.7] - 2026-02-22

### Added
- WinForms GUI (`start-apppackager.ps1`) for visual packager management
  - DataGridView with checkbox selection, version columns, and status indicators
  - Three action buttons: Check Latest, Check MECM, Run Selected
  - Save/load named configurations (site code, file share root, work order)
  - Select-all header checkbox
  - Debug columns toggle (CMName, Script)
  - Timestamped log pane
  - Responsive layout with anchored controls
- Application logo and icon assets (`apppackager-logo.jpg`, `apppackager.ico`)
- `-GetLatestVersionOnly` switch on all packager scripts for version-only queries without download or MECM changes
- `-FileServerPath` parameter on all packager scripts for configurable content staging root
- Metadata header tags (`Vendor:`, `App:`, `CMName:`) on all packager scripts for GUI auto-discovery
- README.md with setup instructions, prerequisites, usage, and new-packager guide

### Changed
- Standardized all 26 packager scripts to generate four content wrapper files (`install.bat`, `install.ps1`, `uninstall.bat`, `uninstall.ps1`) in every versioned content folder
  - `.bat` wrappers are thin launchers that call the corresponding `.ps1` via `PowerShell.exe -NonInteractive -ExecutionPolicy Bypass`
  - `.ps1` files use `Start-Process -Wait -PassThru -NoNewWindow` with `exit $proc.ExitCode` to propagate native installer return codes (0, 1603, 3010, etc.) through to MECM
  - VMware Tools `.bat` wrappers hardcode `exit /b 3010` (reboot required) instead of `%ERRORLEVEL%`
  - Wireshark uninstall converted from batch registry enumeration to PowerShell equivalent
  - VMware Tools `.bat` wrappers simplified — removed `timeout /t 300` (handled by `-Wait` on `Start-Process`)
- Updated README.md Content Staging Layout and Adding a New Packager sections to document the four-file standard

### Packager scripts (26 applications)
- 7-Zip (x64)
- Adobe Acrobat Reader DC (x64)
- ASP.NET Core Hosting Bundle 8
- Google Chrome Enterprise (x64)
- .NET Desktop Runtime 8, 9, 10 (x64)
- Microsoft Edge (x64)
- Mozilla Firefox (x64)
- Git for Windows (x64)
- Greenshot
- Microsoft ODBC Driver 18 for SQL Server
- Microsoft OLE DB Driver for SQL Server
- Microsoft Visual C++ 2015-2022 Redistributable (x86+x64)
- Notepad++ (x64)
- Microsoft Power BI Desktop (x64)
- Tableau Desktop, Prep Builder, Reader (x64)
- Microsoft Teams Enterprise (x64)
- VMware Tools (x64)
- Visual Studio Code (x64)
- Cisco Webex (x64)
- Microsoft WebView2 Evergreen Runtime
- WinSCP
- Wireshark (x64)
