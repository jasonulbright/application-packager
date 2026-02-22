# Changelog

All notable changes to AppPackager are documented in this file.

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
