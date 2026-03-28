# AppPackager

PowerShell scripts and a WinForms GUI that automatically package the latest version of common enterprise applications into Microsoft Endpoint Configuration Manager (MECM) applications.

## What It Does

Each packager script operates in two phases:

**Stage** — Downloads the latest installer from the vendor's official source, extracts metadata (version, publisher, detection info), generates install/uninstall wrapper scripts, and writes a `stage-manifest.json`. Everything is built locally under a configurable download root. No network share or MECM required.

**Package** — Reads the stage manifest, copies the content folder to a versioned UNC network share, and creates an MECM Application with the appropriate deployment type and detection method.

The GUI (`start-apppackager.ps1`) provides a visual front-end that discovers packager scripts automatically, lets you check latest versions, query MECM for current versions, and stage or package selected applications.

![AppPackager GUI](screenshot.png)

## Prerequisites

| Requirement | Details |
|---|---|
| **OS** | Windows 10/11 or Windows Server 2016+ |
| **PowerShell** | 5.1 (ships with Windows) |
| **.NET Framework** | 4.8.2 (required by WinForms GUI) |
| **ConfigMgr Console** | Installed locally — provides `ConfigurationManager.psd1` (Package phase only) |
| **MECM Permissions** | RBAC rights to create Applications and Deployment Types (Package phase only) |
| **Local Admin** | Required for packager script execution |
| **Network Share** | Write access to the SCCM content share, e.g., `\\fileserver\sccm$` (Package phase only) |

## Setup

1. Clone the repository:
   ```
   git clone https://github.com/jasonulbright/application-packager.git
   ```

2. Open PowerShell **as Administrator**.

3. Navigate to the project directory:
   ```powershell
   cd application-packager
   ```

4. For the Package phase, ensure the ConfigMgr PSDrive is available in your session:
   ```powershell
   Import-Module (Join-Path $env:SMS_ADMIN_UI_PATH "..\ConfigurationManager.psd1")
   ```

## Usage

### GUI

Launch the WinForms front-end:

```powershell
.\start-apppackager.ps1
```

Or with custom parameters:

```powershell
.\start-apppackager.ps1 -SiteCode "MCM" -PackagersRoot "D:\CM\Packagers"
```

**No network or MECM actions occur on launch.** The GUI loads packager scripts locally and waits for you to act:

- **Check Latest** — queries vendor sources for the latest version of selected applications
- **Check MECM** — queries your ConfigMgr site for the currently deployed version
- **Stage Packages** — downloads installers, extracts metadata, generates wrappers and manifests locally
- **Package Apps** — reads manifests, copies content to network share, creates MECM applications

Settings that rarely change (Site Code, File Share Root, Download Root, Est/Max Runtime, Company Name) are managed via **File > Preferences** and persisted to `AppPackager.preferences.json`. Company Name is also synced to `packager-preferences.json` for use by ODT-based packagers. The Comment field remains on the main form for per-run entry. Window size and position are persisted automatically across sessions.

**File > Applications** lets you show or hide individual packager scripts from the main grid, grouped by vendor in a checkbox TreeView. Hidden applications are persisted to `AppPackager.preferences.json`. On the first Check MECM run, the tool offers to auto-hide applications not found in your MECM environment.

![Preferences](prefs_screenshot.png)

Citrix Workspace App install switches are configured via **File > Citrix Workspace Switches...**. The dialog provides grouped checkboxes for installation options, plugins/add-ons, update/telemetry, store policy, and ADDLOCAL component selection. Settings are persisted to `citrix-workspace-switches.json` and applied during Stage for both CR and LTSR packagers.

![Citrix Workspace Switches](citrix_prefs_screenshot.png)

Additional grid features:
- **Right-click context menu** on any row — Open Log Folder, Open Staged Folder, Open Network Share, Copy Latest Version
- **Ctrl+Click** any row to open the vendor's product page in the default browser
- **Row hover tooltips** — hover over any row to see the application's description from the packager script
- **Select Update Available** button — auto-checks only rows with "Update available" status after a version check
- **Real-time log streaming** — Stage and Package operations stream packager output line-by-line into the log pane as it runs
- **Tooltips** on all interactive controls — hover over any field or button for a description of its purpose

### Command Line

Run a packager script directly:

```powershell
# Stage only — download, extract metadata, generate wrappers + manifest
.\Packagers\package-chrome.ps1 -StageOnly

# Package only — read manifest, copy to network, create MECM app
.\Packagers\package-chrome.ps1 -PackageOnly -SiteCode "MCM" -Comment "WO#12345" -FileServerPath "\\fileserver\sccm$"

# Both phases in sequence (original behavior)
.\Packagers\package-chrome.ps1 -SiteCode "MCM" -Comment "WO#12345" -FileServerPath "\\fileserver\sccm$"

# Check the latest available version without downloading or creating an MECM application
.\Packagers\package-chrome.ps1 -GetLatestVersionOnly
```

All packager scripts accept the same core parameters:

| Parameter | Description |
|---|---|
| `-SiteCode` | ConfigMgr site code PSDrive name (default: `MCM`) |
| `-Comment` | Free-form change/WO text stored on the CM Application Description |
| `-FileServerPath` | UNC root containing the `Applications` folder (default: `\\fileserver\sccm$`) |
| `-DownloadRoot` | Local root folder for staging (default: `C:\temp\ap`) |
| `-EstimatedRuntimeMins` | MECM deployment type estimated runtime (default: `15`) |
| `-MaximumRuntimeMins` | MECM deployment type maximum runtime (default: `30`) |
| `-StageOnly` | Run only the Stage phase |
| `-PackageOnly` | Run only the Package phase |
| `-GetLatestVersionOnly` | Output the latest version string and exit |
| `-LogPath` | Path to a structured log file (timestamps + severity levels) |

## Supported Applications (115)

| Script | Vendor | Application | Detection Type |
|---|---|---|---|
| package-7zip.ps1 | Igor Pavlov | 7-Zip (x64) | RegistryKeyValue |
| package-adobereader.ps1 | Adobe Inc. | Adobe Acrobat Reader DC (x64) | File version |
| package-anaconda.ps1 | Anaconda, Inc. | Anaconda Distribution (x64) | File existence |
| package-anydesk.ps1 | AnyDesk Software GmbH | AnyDesk | File version |
| package-audacity.ps1 | Audacity Team | Audacity (x64) | File version |
| package-aspnethostingbundle8.ps1 | Microsoft | ASP.NET Core Hosting Bundle 8 | RegistryKey existence |
| package-bitwarden.ps1 | Bitwarden Inc. | Bitwarden Desktop (x64) | File version |
| package-brave.ps1 | Brave Software | Brave Browser (x64) | File version |
| package-ccleaner.ps1 | Piriform Software Ltd. | CCleaner | RegistryKeyValue |
| package-citrixworkspacecr.ps1 | Citrix (Cloud Software Group) | Citrix Workspace (CR) (x64) | RegistryKeyValue |
| package-citrixworkspaceltsr.ps1 | Citrix (Cloud Software Group) | Citrix Workspace (LTSR) (x64) | RegistryKeyValue |
| package-chrome.ps1 | Google | Google Chrome Enterprise (x64) | RegistryKeyValue |
| package-corretto-jdk8-x64.ps1 | Amazon | Amazon Corretto JDK 8 (x64) | RegistryKeyValue |
| package-corretto-jdk8-x86.ps1 | Amazon | Amazon Corretto JDK 8 (x86) | RegistryKeyValue |
| package-corretto-jdk11-x64.ps1 | Amazon | Amazon Corretto JDK 11 (x64) | RegistryKeyValue |
| package-corretto-jdk11-x86.ps1 | Amazon | Amazon Corretto JDK 11 (x86) | RegistryKeyValue |
| package-corretto-jdk17.ps1 | Amazon | Amazon Corretto JDK 17 (x64) | RegistryKeyValue |
| package-corretto-jdk21.ps1 | Amazon | Amazon Corretto JDK 21 (x64) | RegistryKeyValue |
| package-corretto-jdk25.ps1 | Amazon | Amazon Corretto JDK 25 (x64) | RegistryKeyValue |
| package-cpuz.ps1 | CPUID | CPU-Z | File version |
| package-cutepdfwriter.ps1 | Acro Software Inc. | CutePDF Writer | RegistryKeyValue |
| package-dbeaver.ps1 | DBeaver Corp | DBeaver Community | File version |
| package-dotnet8.ps1 | Microsoft | .NET Desktop Runtime 8 (x64) | Compound (AND, 2x File existence) |
| package-Dotnet9x64.ps1 | Microsoft | .NET Desktop Runtime 9 (x64) | File existence |
| package-Dotnet10x64.ps1 | Microsoft | .NET Desktop Runtime 10 (x64) | File existence |
| package-drawio.ps1 | JGraph Ltd | draw.io | RegistryKeyValue |
| package-edge.ps1 | Microsoft | Microsoft Edge (x64) | Compound (OR, 2x File version) |
| package-everything.ps1 | Voidtools | Everything (x64) | RegistryKeyValue |
| package-filezilla.ps1 | FileZilla Project | FileZilla Client (x64) | RegistryKeyValue |
| package-firefox.ps1 | Mozilla | Mozilla Firefox (x64) | File version |
| package-foxitreader.ps1 | Foxit Software | Foxit PDF Reader (x64) | File version |
| package-gimp.ps1 | The GIMP Team | GIMP (x64) | RegistryKeyValue |
| package-git.ps1 | Git | Git for Windows (x64) | Script (git.exe --version) |
| package-greenshot.ps1 | Greenshot | Greenshot | File existence |
| package-inkscape.ps1 | Inkscape Project | Inkscape (x64) | RegistryKeyValue |
| package-keepass.ps1 | Dominik Reichl | KeePass | RegistryKeyValue |
| package-libreoffice.ps1 | The Document Foundation | LibreOffice (x64) | RegistryKeyValue |
| package-malwarebytes.ps1 | Malwarebytes | Malwarebytes | RegistryKeyValue |
| package-mremoteng.ps1 | mRemoteNG | mRemoteNG | RegistryKeyValue |
| package-m365apps-x64.ps1 | Microsoft | M365 Apps for Enterprise (x64) | File version (WINWORD.EXE) |
| package-m365apps-x86.ps1 | Microsoft | M365 Apps for Enterprise (x86) | File version (WINWORD.EXE) |
| package-m365visio-x64.ps1 | Microsoft | M365 Visio (x64) | File version (VISIO.EXE) |
| package-m365visio-x86.ps1 | Microsoft | M365 Visio (x86) | File version (VISIO.EXE) |
| package-m365project-x64.ps1 | Microsoft | M365 Project (x64) | File version (WINPROJ.EXE) |
| package-m365project-x86.ps1 | Microsoft | M365 Project (x86) | File version (WINPROJ.EXE) |
| package-msodbcsql18.ps1 | Microsoft | ODBC Driver 18 for SQL Server | RegistryKeyValue |
| package-msoledb.ps1 | Microsoft | OLE DB Driver for SQL Server | RegistryKeyValue |
| package-msvcruntimes.ps1 | Microsoft | VC++ 2015-2022 Redistributable (x86+x64) | Compound (AND, 2x RegistryKeyValue) |
| package-nodejs.ps1 | OpenJS Foundation | Node.js LTS (x64) | RegistryKeyValue |
| package-notepadplusplus.ps1 | Notepad++ | Notepad++ (x64) | File version |
| package-obsidian.ps1 | Obsidian | Obsidian | File existence (per-user) |
| package-opera.ps1 | Opera Software | Opera Browser (x64) | File version |
| package-paintdotnet.ps1 | dotPDN LLC | Paint.NET (x64) | RegistryKeyValue |
| package-pgadmin4.ps1 | pgAdmin Development Team | pgAdmin 4 | RegistryKeyValue |
| package-postman.ps1 | Postman Inc. | Postman | File existence (per-user) |
| package-powerbidesktop.ps1 | Microsoft | Power BI Desktop (x64) | File version |
| package-powershell7.ps1 | Microsoft | PowerShell 7 (x64) | RegistryKeyValue |
| package-powershell7lts.ps1 | Microsoft | PowerShell 7 LTS (x64) | RegistryKeyValue |
| package-powertoys.ps1 | Microsoft Corporation | PowerToys (x64) | File version |
| package-positron.ps1 | Posit Software, PBC | Positron (x64) | File existence |
| package-postgresql13.ps1 | PostgreSQL Global Development Group | PostgreSQL 13 (x64) | File version |
| package-postgresql14.ps1 | PostgreSQL Global Development Group | PostgreSQL 14 (x64) | File version |
| package-postgresql15.ps1 | PostgreSQL Global Development Group | PostgreSQL 15 (x64) | File version |
| package-postgresql16.ps1 | PostgreSQL Global Development Group | PostgreSQL 16 (x64) | File version |
| package-postgresql17.ps1 | PostgreSQL Global Development Group | PostgreSQL 17 (x64) | File version |
| package-putty.ps1 | Simon Tatham | PuTTY (x64) | RegistryKeyValue |
| package-pycharm.ps1 | JetBrains | PyCharm Community Edition (x64) | File existence |
| package-python.ps1 | Python Software Foundation | Python (x64) | File existence |
| package-r.ps1 | The R Foundation | R for Windows (x64) | File existence |
| package-rstudio.ps1 | Posit Software, PBC | RStudio Desktop (x64) | RegistryKeyValue |
| package-sharex.ps1 | ShareX Team | ShareX | File version |
| package-slack.ps1 | Slack Technologies | Slack (x64) | RegistryKeyValue |
| package-soapui.ps1 | SmartBear Software | SoapUI | File existence |
| package-sysinternals.ps1 | Microsoft | Sysinternals Suite | File existence |
| package-ssms.ps1 | Microsoft | SQL Server Management Studio | File version (Ssms.exe) |
| package-teams.ps1 | Microsoft | Microsoft Teams Enterprise (x64) | Script (Get-AppxPackage) |
| package-teamviewer.ps1 | TeamViewer | TeamViewer (x64) | RegistryKeyValue |
| package-teamviewerhost.ps1 | TeamViewer | TeamViewer Host (x64) | File |
| package-tableaudesktop.ps1 | Salesforce (Tableau) | Tableau Desktop (x64) | File version |
| package-tableauprep.ps1 | Salesforce (Tableau) | Tableau Prep Builder (x64) | File version |
| package-tableaureader.ps1 | Salesforce (Tableau) | Tableau Reader (x64) | File version |
| package-temurin-jdk8-x64.ps1 | Eclipse Adoptium | Eclipse Temurin JDK 8 (x64) | RegistryKeyValue |
| package-temurin-jdk8-x86.ps1 | Eclipse Adoptium | Eclipse Temurin JDK 8 (x86) | RegistryKeyValue |
| package-temurin-jdk11-x64.ps1 | Eclipse Adoptium | Eclipse Temurin JDK 11 (x64) | RegistryKeyValue |
| package-temurin-jdk11-x86.ps1 | Eclipse Adoptium | Eclipse Temurin JDK 11 (x86) | RegistryKeyValue |
| package-temurin-jdk17.ps1 | Eclipse Adoptium | Eclipse Temurin JDK 17 (x64) | RegistryKeyValue |
| package-temurin-jdk21.ps1 | Eclipse Adoptium | Eclipse Temurin JDK 21 (x64) | RegistryKeyValue |
| package-temurin-jdk25.ps1 | Eclipse Adoptium | Eclipse Temurin JDK 25 (x64) | RegistryKeyValue |
| package-temurin-jre8-x64.ps1 | Eclipse Adoptium | Eclipse Temurin JRE 8 (x64) | RegistryKeyValue |
| package-temurin-jre8-x86.ps1 | Eclipse Adoptium | Eclipse Temurin JRE 8 (x86) | RegistryKeyValue |
| package-temurin-jre11-x64.ps1 | Eclipse Adoptium | Eclipse Temurin JRE 11 (x64) | RegistryKeyValue |
| package-temurin-jre11-x86.ps1 | Eclipse Adoptium | Eclipse Temurin JRE 11 (x86) | RegistryKeyValue |
| package-temurin-jre17.ps1 | Eclipse Adoptium | Eclipse Temurin JRE 17 (x64) | RegistryKeyValue |
| package-temurin-jre21.ps1 | Eclipse Adoptium | Eclipse Temurin JRE 21 (x64) | RegistryKeyValue |
| package-temurin-jre25.ps1 | Eclipse Adoptium | Eclipse Temurin JRE 25 (x64) | RegistryKeyValue |
| package-thunderbird.ps1 | Mozilla Foundation | Thunderbird (x64) | File version |
| package-treesizefree.ps1 | JAM Software | TreeSize Free | File version |
| package-tortoisegit.ps1 | TortoiseGit | TortoiseGit (x64) | RegistryKeyValue |
| package-tortoisesvn.ps1 | TortoiseSVN | TortoiseSVN (x64) | RegistryKeyValue |
| package-vmwaretools.ps1 | Broadcom | VMware Tools (x64) | File version |
| package-vmwareworkstation.ps1 | Broadcom | VMware Workstation Pro | File version |
| package-vs2026.ps1 | Microsoft | Visual Studio 2026 Enterprise | File version (devenv.exe) |
| package-vs2026community.ps1 | Microsoft | Visual Studio 2026 Community | File version (devenv.exe) |
| package-vim.ps1 | The Vim Project | Vim (x64) | RegistryKeyValue |
| package-vlc.ps1 | VideoLAN | VLC Media Player (x64) | RegistryKeyValue |
| package-vscode.ps1 | Microsoft | Visual Studio Code (x64) | File version |
| package-webex.ps1 | Cisco | Webex (x64) | RegistryKeyValue |
| package-webview2.ps1 | Microsoft | WebView2 Evergreen Runtime | File version |
| package-windirstat.ps1 | WinDirStat Team | WinDirStat (x64) | File version |
| package-winmerge.ps1 | WinMerge | WinMerge (x64) | File version |
| package-winscp.ps1 | WinSCP | WinSCP | RegistryKeyValue |
| package-winrar.ps1 | win.rar GmbH | WinRAR (x64) | RegistryKeyValue |
| package-wireshark.ps1 | Wireshark Foundation | Wireshark (x64) | RegistryKeyValue |
| package-wiztree.ps1 | Antibody Software | WizTree | File version |
| package-zoom.ps1 | Zoom Video Communications | Zoom Workplace (x64) | File existence (per-user) |

## Vendor Version Monitor

The Version Monitor is a headless companion tool that compares MECM-deployed application versions against the latest vendor releases, flags stale packages, and optionally queries the NIST NVD for known CVEs. It produces a self-contained HTML report.

![Vendor Version Monitor Report](screenshot-vendor-version-monitor-report-sample.png)

```powershell
# Full run: MECM + vendor checks + NVD CVE lookups
.\VersionMonitor\Start-VersionMonitor.ps1

# Vendor version checks only (no MECM or NVD dependency)
.\VersionMonitor\Start-VersionMonitor.ps1 -SkipMECM -SkipNVD

# Simulate stale versions for testing report rendering and CVE lookups
.\VersionMonitor\Start-VersionMonitor.ps1 -SimulateStale
```

The monitor discovers all `package-*.ps1` scripts in the sibling `Packagers/` folder and reads metadata directly from their headers — no separate catalog file needed. CPE strings embedded in packager headers enable NVD CVE lookups for stale applications.

| Feature | Details |
|---|---|
| **Packager discovery** | Auto-discovers all 115 packager scripts via relative path |
| **Version checking** | Calls each packager with `-GetLatestVersionOnly` |
| **MECM comparison** | Queries ConfigMgr for deployed versions |
| **NVD CVE lookup** | Queries NIST NVD API for stale apps with CPE headers |
| **Rate limiting** | Sliding-window rate limiter with configurable limits |
| **NVD caching** | JSON cache with configurable TTL (default 6 hours) |
| **HTML report** | Self-contained report with status badges, CVE pills, CVSS scores |
| **Simulation mode** | Override MECM versions via `simulate-overrides.json` for testing |
| **Notifications** | Drop folder copy and webhook stub (extensible) |
| **Log/report cleanup** | Configurable retention for old logs and reports |

Configuration is in `VersionMonitor/monitor-config.json`. Log and report folders default to `VersionMonitor/Logs/` and `VersionMonitor/Reports/` when not specified in config.

### Packager header tags for Version Monitor

Each packager script can include optional metadata tags parsed by the monitor:

```powershell
<#
Vendor: Igor Pavlov
App: 7-Zip (x64)
CMName: 7-Zip
VendorUrl: https://www.7-zip.org/
CPE: cpe:2.3:a:7-zip:7-zip:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.7-zip.org/history.txt
DownloadPageUrl: https://www.7-zip.org/download.html
#>
```

| Tag | Purpose |
|---|---|
| `CPE` | NVD Common Platform Enumeration string for CVE lookups |
| `ReleaseNotesUrl` | Link shown in the HTML report's Links column |
| `DownloadPageUrl` | Link shown in the HTML report's Links column |

## Content Staging Layout

### Local staging (Stage phase)

```
C:\temp\ap\
  <App>\
    staged-version.txt              # Version marker for Package phase
    <Version>\
      installer.msi (or .exe)
      install.bat
      install.ps1
      uninstall.bat
      uninstall.ps1
      stage-manifest.json           # Metadata for Package phase
```

### Network share (Package phase)

```
\\fileserver\sccm$\
  Applications\
    <Vendor>\
      <Application>\
        <Version>\
          installer.msi (or .exe)
          install.bat
          install.ps1
          uninstall.bat
          uninstall.ps1
```

Every content folder contains **four wrapper files** alongside the installer. The `.bat` files are thin wrappers that call the corresponding `.ps1`:

```batch
@echo off
PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0install.ps1"
exit /b %ERRORLEVEL%
```

The `.ps1` files contain the actual install/uninstall logic using `Start-Process -Wait -PassThru -NoNewWindow` and `exit $proc.ExitCode` to propagate native installer return codes (0, 1603, 3010, etc.) through to MECM.

**Why `.bat` wrappers?** MECM's Deployment Type "Hidden" visibility dropdown appends `/q` to install parameters. This conflicts with installers that already specify `/qn` or `/qb`. The `.bat` wrapper with `@echo off` prevents this by hiding the command window without injecting silent flags.

### Stage manifest (`stage-manifest.json`)

Written by the Stage phase, read by the Package phase. Contains all metadata needed to create the MECM application without re-downloading or re-parsing the installer:

```json
{
  "SchemaVersion": 2,
  "StagedAt": "2026-03-28T10:00:00Z",
  "AppName": "7-Zip - 26.00 (x64)",
  "Publisher": "Igor Pavlov",
  "SoftwareVersion": "26.00",
  "InstallerFile": "7z2600-x64.msi",
  "InstallerType": "MSI",
  "InstallArgs": "/qn /norestart",
  "UninstallArgs": "/qn /norestart",
  "ProductCode": "{23170F69-40C1-2702-2600-000001000000}",
  "RunningProcess": ["7zFM", "7zG"],
  "Detection": {
    "Type": "RegistryKeyValue",
    "RegistryKeyRelative": "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{23170F69-...}",
    "ValueName": "DisplayVersion",
    "ExpectedValue": "26.00.00.0",
    "Operator": "IsEquals",
    "Is64Bit": true
  }
}
```

Five detection types are supported: `RegistryKeyValue`, `RegistryKey`, `File`, `Script`, and `Compound` (multiple clauses with AND/OR connectors).

Schema v2 adds optional fields for deployment tool integration (PSADT, Intune, custom wrappers): `InstallerType`, `InstallArgs`, `UninstallArgs`, `UninstallCommand`, `ProductCode`, `RunningProcess`. All fields are optional -- v1 manifests are still accepted.

## Project Structure

```
application-packager/
  start-apppackager.ps1              # WinForms GUI
  apppackager-logo.jpg               # GUI window icon / logo
  apppackager.ico                    # Application icon
  Packagers/
    AppPackagerCommon.psm1           # Shared module (logging, wrappers, MECM helpers)
    AppPackagerCommon.psd1           # Module manifest
    packager-preferences.json        # Universal packager settings (CompanyName, etc.)
    AppPackagerCommon.Tests.ps1      # Pester unit tests (89 tests)
    package-7zip.ps1                 # One script per application (115 total)
    package-chrome.ps1
    ...
  Tests/
    Test-AllPackagers.ps1            # Automated install/uninstall validation harness
  VersionMonitor/
    Start-VersionMonitor.ps1         # Headless version monitor entry point
    monitor-config.json              # Monitor configuration (MECM, NVD, report settings)
    simulate-overrides.json          # Simulated MECM versions for testing
    Module/
      VersionMonitorCommon.psm1      # Monitor module (discovery, comparison, NVD, HTML)
      VersionMonitorCommon.psd1      # Module manifest
    Logs/                            # Auto-created log files
    Reports/                         # Auto-created HTML reports
  CHANGELOG.md
  README.md
```

## Adding a New Packager

1. Create a new file in `Packagers/` named `package-<appname>.ps1`

2. Add metadata tags in the script header (parsed by the GUI):
   ```powershell
   <#
   Vendor: Acme Corp
   App: Acme Widget (x64)
   CMName: Acme Widget
   VendorUrl: https://acme.example.com/widget
   CPE: cpe:2.3:a:acme:widget:*:*:*:*:*:*:*:*
   ReleaseNotesUrl: https://acme.example.com/releases
   DownloadPageUrl: https://acme.example.com/download
   #>
   ```

3. Implement the standard parameter block:
   ```powershell
   param(
       [string]$SiteCode = "MCM",
       [string]$Comment = "WO#00000001234567",
       [string]$FileServerPath = "\\fileserver\sccm$",
       [string]$DownloadRoot = "C:\temp\ap",
       [int]$EstimatedRuntimeMins = 15,
       [int]$MaximumRuntimeMins = 30,
       [string]$LogPath,
       [switch]$GetLatestVersionOnly,
       [switch]$StageOnly,
       [switch]$PackageOnly
   )
   ```

4. Import the shared module:
   ```powershell
   Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
   Initialize-Logging -LogPath $LogPath
   ```

5. Implement `Invoke-Stage<App>`:
   - Download the installer
   - Extract metadata (version, publisher, detection info)
   - Generate wrapper content and call `Write-ContentWrappers`
   - Call `Write-StageManifest` with detection block
   - Write `staged-version.txt`

6. Implement `Invoke-Package<App>`:
   - Read `staged-version.txt` and `stage-manifest.json` via `Read-StageManifest`
   - Copy content to network share via `Get-NetworkAppRoot`
   - Call `New-MECMApplicationFromManifest`

7. Wire up the main block:
   ```powershell
   if ($StageOnly) { Invoke-StageAcmeWidget }
   elseif ($PackageOnly) { Invoke-PackageAcmeWidget }
   else { Invoke-StageAcmeWidget; Invoke-PackageAcmeWidget }
   ```

8. The `-GetLatestVersionOnly` switch must output **only** the version string to stdout and exit.

The GUI will automatically discover and display the new script on next launch.

## Shared Module (`AppPackagerCommon.psm1`)

All packager scripts import the shared module which provides:

| Function | Purpose |
|---|---|
| `Write-Log` | Timestamped, severity-tagged logging to console and optional file |
| `Initialize-Logging` | Sets up log file output |
| `Invoke-DownloadWithRetry` | curl.exe download wrapper with 1 retry and 5s delay |
| `Test-IsAdmin` | Checks for administrator elevation |
| `Connect-CMSite` | Imports ConfigMgr module and sets PSDrive location |
| `Initialize-Folder` | Creates directory if missing |
| `Test-NetworkShareAccess` | Verifies UNC path is writable |
| `Get-MsiPropertyMap` | Reads MSI properties (ProductName, ProductVersion, Manufacturer, ProductCode) |
| `Find-UninstallEntry` | Searches ARP registry keys by DisplayName pattern |
| `Write-ContentWrappers` | Generates install/uninstall .bat + .ps1 wrapper files |
| `New-MsiWrapperContent` | Returns MSI install/uninstall .ps1 content strings |
| `New-ExeWrapperContent` | Returns EXE install/uninstall .ps1 content strings |
| `Get-NetworkAppRoot` | Constructs and initializes the network share path |
| `Write-StageManifest` / `Read-StageManifest` | JSON manifest serialization |
| `New-MECMApplicationFromManifest` | Creates MECM Application + deployment type from manifest |
| `Remove-CMApplicationRevisionHistoryByCIId` | Trims old application revisions |
| `Get-PackagerPreferences` | Reads `packager-preferences.json` for universal settings (e.g., CompanyName) |
| `New-OdtConfigXml` | Generates full ODT configuration XML for M365 download/install phases |
| `Get-LatestTemurinRelease` | Queries Adoptium API for latest Eclipse Temurin MSI (JRE/JDK, x64/x86) |
| `Get-LatestCorrettoRelease` | Queries GitHub releases for latest Amazon Corretto MSI (JDK, x64/x86) |

## License

This project is licensed under the [MIT License](LICENSE).

## Author

Jason Ulbright
