# application-packager

PowerShell scripts that automatically package the latest version of common enterprise
applications into Microsoft Endpoint Configuration Manager (MECM) applications.

## What it does

Each packager script:
1. Fetches the latest version from the vendor's official source (API, download page, or release feed)
2. Downloads the installer
3. Stages content to a versioned UNC network share
4. Creates an MECM Application with the appropriate deployment type and detection method

## Supported applications

| Script | Vendor | Detection method |
|---|---|---|
| package-7zip.ps1 | 7-Zip | File version |
| package-adobereader.ps1 | Adobe Inc. | File version |
| package-aspnethostingbundle8.ps1 | Microsoft | Windows Installer (MSI) |
| package-chrome.ps1 | Google | Windows Installer (MSI) |
| package-dotnet8.ps1 | Microsoft | Windows Installer (MSI) |
| package-Dotnet9x64.ps1 | Microsoft | Windows Installer (MSI) |
| package-Dotnet10x64.ps1 | Microsoft | Windows Installer (MSI) |
| package-edge.ps1 | Microsoft | Windows Installer (MSI) |
| package-firefox.ps1 | Mozilla | File version |
| package-git.ps1 | Git Development Community | PowerShell registry script |
| package-greenshot.ps1 | Greenshot | File version |
| package-msodbcsql18.ps1 | Microsoft | Windows Installer (MSI) |
| package-msoledb.ps1 | Microsoft | Windows Installer (MSI) |
| package-msvcruntimes.ps1 | Microsoft | Windows Installer (MSI) |
| package-notepadplusplus.ps1 | Notepad++ | File version |
| package-powerbidesktop.ps1 | Microsoft | File version |
| package-tableaudesktop.ps1 | Salesforce (Tableau) | File version |
| package-tableauprep.ps1 | Salesforce (Tableau) | File version |
| package-tableaureader.ps1 | Salesforce (Tableau) | File version |
| package-teams.ps1 | Microsoft | PowerShell registry script |
| package-vmwaretools.ps1 | VMware | File version |
| package-vscode.ps1 | Microsoft | File version |
| package-webex.ps1 | Cisco | File version |
| package-webview2.ps1 | Microsoft | Windows Installer (MSI) |
| package-winscp.ps1 | WinSCP | File version |
| package-wireshark.ps1 | Wireshark Foundation | File version |

## Usage

```powershell
# Package the latest version of an application
.\Packagers\package-chrome.ps1 -SiteCode "MCM" -Comment "WO#12345" -FileServerPath "\\fileserver\sccm$"

# Check the latest available version without downloading or creating an MECM application
.\Packagers\package-chrome.ps1 -GetLatestVersionOnly
```

## Requirements

- PowerShell 5.1
- ConfigMgr Admin Console installed (for `ConfigurationManager.psd1`)
- RBAC rights to create Applications and Deployment Types in MECM
- Local administrator
- Write access to the SCCM content share (`FileServerPath`)
