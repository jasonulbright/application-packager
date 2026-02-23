# Changelog

All notable changes to AppPackager are documented in this file.

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

### Fixed
- `Write-Log` now accepts empty strings via `[AllowEmptyString()]` attribute — `Write-Log ""` for blank log lines was rejected by PowerShell's mandatory parameter validation
- MSI-based Stage phase (7-Zip) derives ARP detection from MSI properties (`ProductCode` → registry key, `ProductVersion` → `DisplayVersion`) instead of temp install/uninstall — temporary installation of products with shell extensions (e.g., 7-Zip context menu) crashed `explorer.exe` when the extension DLL was unregistered while still loaded
- Content wrapper `.ps1` files use proper `-ArgumentList` array (`@('/i', "`"$msiPath`"", '/qn', '/norestart')`) instead of a single string — `Start-Process` misparses a single string with embedded quotes, causing msiexec to receive no arguments
- Content wrapper `.ps1` files written with `-Encoding ASCII` instead of `-Encoding UTF8` — PowerShell 5.1's `UTF8` encoding always prepends a BOM (`EF BB BF`), which appeared as `∩╗┐` at position 0
- Stage manifest `Detection.DisplayVersion` now stores the raw MSI `ProductVersion` (e.g., `26.00.00.0`) instead of the normalized display version (`26.00`) — the MECM `IsEquals` detection clause must match the actual 4-part registry value

### Deprecated
- **Tableau Desktop/Prep/Reader:** Deprecated and moved to `Archive/deprecated/`. Salesforce removed public download access at `downloads.tableau.com/tssoftware/` (returns 404 for all versions). The original pre-refactored scripts used the identical URL pattern. Scripts were structurally complete and passing all verification checks but could not stage due to the upstream change.

### Staging Validation
23 of 23 active products staged and validated successfully.

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
