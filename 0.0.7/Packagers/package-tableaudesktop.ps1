<#
Vendor: Tableau
App: Tableau Desktop (x64)
CMName: Tableau 20

.SYNOPSIS
    Automates downloading the latest Tableau Desktop installer (x64) and creating an MECM application.

.DESCRIPTION
    Creates an MECM application for Tableau Desktop, using registry-based detection
    and batch file installation. Temporarily installs the product to extract registry data (DisplayName, DisplayVersion, Publisher, InstallLocation)
    and file version from the detection executable. Application and deployment type names use the registry DisplayName.
    Stores installers and batch files in \\fileshare\sccm$\Applications\Tableau\<Product>\<Version>.
    MECM settings: 30-minute max runtime, 10-minute estimated runtime, system installation, no user logon requirement.
    Checks for existing files locally and on network share to skip redundant downloads/copies.
    Uses Add-CMScriptDeploymentType for the detection clause.

.NOTES
    - Run with admin privileges for registry access, installation, and MECM operations.
    - Requires PowerShell 5.1 and Configuration Manager module.
    - Temporarily installs and uninstalls product to extract registry data and file versions.
    - No residual files or registry entries are left after uninstallation.
    - Assumes installations always succeed in a clean packaging environment.
#>

param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

# --- Configuration Variables ---
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Base URL for Tableau software downloads
$BaseDownloadUrl = "https://downloads.tableau.com/tssoftware/"

# URL to Tableau's release notes page
$ReleaseNotesUrl = "https://www.tableau.com/support/releases"

# Download and local temp directories
$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\Tableau_installers"
$LocalTempRoot = "C:\temp"
$TableauRootNetworkPath = Join-Path $FileServerPath "Applications\Tableau"

# Define the Tableau product and its metadata
$TableauProducts = @(
    @{
        Name = "Tableau {0}" # e.g., "Tableau 2025.2.3"
        Subfolder = "Desktop"
        FileNamePattern = "TableauDesktop-64bit-{0}.exe"
        ProgramsAndFeaturesNamePrefix = "Tableau 20"
        DetectionFile = "tableau.exe"
        InstallBatContent = '"%~dp0TableauDesktop-64bit-{0}.exe" /install /quiet /norestart ACCEPTEULA=1 REMOVEINSTALLEDAPP=1 SENDTELEMETRY=0'
        UninstallBatContent = '"%~dp0TableauDesktop-64bit-{0}.exe" /uninstall /quiet /norestart'
        RegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    }
)

# --- Functions ---

function Test-IsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Warning "Failed to check admin privileges: $($_.Exception.Message)"
        return $false
    }
}

function Connect-CMSite {
    param([Parameter(Mandatory)][string]$SiteCode)

    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
        Write-Host "Connected to CM site: $SiteCode"
        return $true
    }
    catch {
        Write-Error "Failed to connect to CM site: $($_.Exception.Message)"
        return $false
    }
}

function Compare-TableauVersion {
    param (
        [string]$Version1,
        [string]$Version2
    )
    $v1Parts = $Version1 -split '-' | ForEach-Object { [int]$_ }
    $v2Parts = $Version2 -split '-' | ForEach-Object { [int]$_ }
    if ($v1Parts.Count -lt 3) { $v1Parts += 0 }
    if ($v2Parts.Count -lt 3) { $v2Parts += 0 }
    if ($v1Parts[0] -gt $v2Parts[0]) { return 1 }
    if ($v1Parts[0] -lt $v2Parts[0]) { return -1 }
    if ($v1Parts[1] -gt $v2Parts[1]) { return 1 }
    if ($v1Parts[1] -lt $v2Parts[1]) { return -1 }
    if ($v1Parts[2] -gt $v2Parts[2]) { return 1 }
    if ($v1Parts[2] -lt $v2Parts[2]) { return -1 }
    return 0
}

function Get-LatestTableauVersion {
    param([switch]$Quiet)

    $originalLocation = Get-Location
    if (-not $Quiet) { Write-Host "Current location before getting latest version: ${originalLocation}" }
    try {
        Set-Location $PSScriptRoot -ErrorAction Stop
        if (-not $Quiet) { Write-Host "Set location to script directory for version check: ${PSScriptRoot}" }
        if (-not $Quiet) { Write-Host "Attempting to find the latest Tableau version from: $ReleaseNotesUrl" }
        $HtmlContent = (curl.exe -L --fail --silent --show-error $ReleaseNotesUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch release notes: $ReleaseNotesUrl" }
        $versionPattern = '\b(20\d{2})[.-](\d+)(?:[.-](\d+))?\b'
        $matches = [regex]::Matches($HtmlContent, $versionPattern)

        if ($matches.Count -eq 0) {
            Write-Error "No version matches found in release notes."
            return $null
        }

        $versions = foreach ($m in $matches) {
            $year = $m.Groups[1].Value
            $minor = $m.Groups[2].Value
            $patch = $m.Groups[3].Value
            if (-not $patch) { $patch = "0" }
            "{0}-{1}-{2}" -f $year, $minor, $patch
        }

        $versions = $versions | Select-Object -Unique

        $latest = $versions | Sort-Object -Descending -Property @{ Expression = { ($_ -split '-')[0] -as [int] } },
                                                 @{ Expression = { ($_ -split '-')[1] -as [int] } },
                                                 @{ Expression = { ($_ -split '-')[2] -as [int] } } |
                  Select-Object -First 1

        if (-not $Quiet) { Write-Host "Latest Tableau version found (dash format): $latest" }
        return $latest
    }
    catch {
        Write-Error "Error determining latest Tableau version: $($_.Exception.Message)"
        return $null
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
    }
}

function Get-InstalledAppRegistryInfo {
    param(
        [Parameter(Mandatory)] [string]$DisplayNamePrefix
    )

    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $uninstallPaths) {
        $apps = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName -like "$DisplayNamePrefix*" } |
            Sort-Object -Property DisplayVersion -Descending

        if ($apps -and $apps.Count -gt 0) {
            return $apps | Select-Object -First 1
        }
    }

    return $null
}

function Get-FileVersion {
    param(
        [Parameter(Mandatory)] [string]$FilePath
    )
    if (-not (Test-Path -LiteralPath $FilePath)) { return $null }
    try {
        return (Get-Item -LiteralPath $FilePath).VersionInfo.FileVersion
    }
    catch {
        return $null
    }
}

function New-ContentFolder {
    param(
        [Parameter(Mandatory)] [string]$Path
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-InstallUninstallBats {
    param(
        [Parameter(Mandatory)] [string]$FolderPath,
        [Parameter(Mandatory)] [string]$InstallContent,
        [Parameter(Mandatory)] [string]$UninstallContent
    )

    $installBat = Join-Path $FolderPath "install.bat"
    $uninstallBat = Join-Path $FolderPath "uninstall.bat"

    Set-Content -LiteralPath $installBat -Value "@echo off`r`nsetlocal`r`n$InstallContent`r`nexit /b 0`r`n" -Encoding ASCII
    Set-Content -LiteralPath $uninstallBat -Value "@echo off`r`nsetlocal`r`n$UninstallContent`r`nexit /b 0`r`n" -Encoding ASCII
}

function Download-InstallerIfNeeded {
    param(
        [Parameter(Mandatory)] [string]$DownloadUrl,
        [Parameter(Mandatory)] [string]$LocalFilePath
    )

    if (Test-Path -LiteralPath $LocalFilePath) {
        Write-Host "Installer already exists locally: $LocalFilePath"
        return
    }

    Write-Host "Downloading installer from: $DownloadUrl"
    curl.exe -L --fail --silent --show-error -o $LocalFilePath $DownloadUrl
    if ($LASTEXITCODE -ne 0) { throw "Download failed: $DownloadUrl" }
}

function Copy-InstallerToNetworkIfNeeded {
    param(
        [Parameter(Mandatory)] [string]$LocalFilePath,
        [Parameter(Mandatory)] [string]$NetworkFilePath
    )

    if (Test-Path -LiteralPath $NetworkFilePath) {
        Write-Host "Installer already exists on network share: $NetworkFilePath"
        return
    }

    Copy-Item -LiteralPath $LocalFilePath -Destination $NetworkFilePath -Force -ErrorAction Stop
}

function Install-ForMetadataExtraction {
    param(
        [Parameter(Mandatory)] [string]$InstallerPath,
        [Parameter(Mandatory)] [string]$InstallArgs
    )

    Write-Host "Installing temporarily for metadata extraction: $InstallerPath"
    Start-Process -FilePath $InstallerPath -ArgumentList $InstallArgs -Wait -NoNewWindow
}

function Uninstall-AfterMetadataExtraction {
    param(
        [Parameter(Mandatory)] [string]$InstallerPath,
        [Parameter(Mandatory)] [string]$UninstallArgs
    )

    Write-Host "Uninstalling after metadata extraction: $InstallerPath"
    Start-Process -FilePath $InstallerPath -ArgumentList $UninstallArgs -Wait -NoNewWindow
}

function Remove-CMApplicationRevisionHistoryByCIId {
    param(
        [Parameter(Mandatory)][UInt32]$CI_ID,
        [UInt32]$KeepLatest = 1
    )
    $history = Get-CMApplicationRevisionHistory -Id $CI_ID -ErrorAction SilentlyContinue
    if (-not $history) { return }
    $revs = @()
    foreach ($h in @($history)) {
        if ($h.PSObject.Properties.Name -contains 'Revision')  { $revs += [UInt32]$h.Revision;   continue }
        if ($h.PSObject.Properties.Name -contains 'CIVersion') { $revs += [UInt32]$h.CIVersion; continue }
    }
    $revs = $revs | Sort-Object -Unique -Descending
    if ($revs.Count -le $KeepLatest) { return }
    foreach ($rev in ($revs | Select-Object -Skip $KeepLatest)) {
        Remove-CMApplicationRevisionHistory -Id $CI_ID -Revision $rev -Force -ErrorAction Stop
    }
}

function New-MECMApplicationForTableau {
    param(
        [Parameter(Mandatory)] [string]$AppName,
        [Parameter(Mandatory)] [string]$Publisher,
        [Parameter(Mandatory)] [string]$SoftwareVersion,
        [Parameter(Mandatory)] [string]$ContentLocation,
        [Parameter(Mandatory)] [string]$InstallCommand,
        [Parameter(Mandatory)] [string]$UninstallCommand,
        [Parameter(Mandatory)] [string]$DetectionExePath,
        [Parameter(Mandatory)] [string]$DetectionExeVersion
    )

    if (-not (Connect-CMSite -SiteCode $SiteCode)) {
        throw "Unable to connect to CM site drive."
    }

    $existingApp = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
    if ($existingApp) {
        Write-Warning "Application already exists: $AppName"
        return
    }

    $cmApp = New-CMApplication -Name $AppName -Publisher $Publisher -SoftwareVersion $SoftwareVersion -Description $Comment -LocalizedApplicationName $AppName -ErrorAction Stop

    $detectionClause = New-CMDetectionClauseFile `
        -Path (Split-Path -Path $DetectionExePath -Parent) `
        -FileName (Split-Path -Path $DetectionExePath -Leaf) `
        -PropertyType Version `
        -ExpressionOperator GreaterEquals `
        -ExpectedValue $DetectionExeVersion `
        -Value

    Add-CMScriptDeploymentType `
        -ApplicationName $AppName `
        -DeploymentTypeName $AppName `
        -InstallCommand $InstallCommand `
        -UninstallCommand $UninstallCommand `
        -ContentLocation $ContentLocation `
        -InstallationBehaviorType InstallForSystem `
        -LogonRequirementType WhetherOrNotUserLoggedOn `
        -EstimatedRuntimeMins 10 `
        -MaximumRuntimeMins 30 `
        -AddDetectionClause $detectionClause `
        -ErrorAction Stop | Out-Null
    Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

    Write-Host "Created MECM application: $AppName"
}

# --- Latest-only mode ---
if ($GetLatestVersionOnly) {
    try {
        $v = Get-LatestTableauVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output ($v -replace '-', '.')
        exit 0
    }
    catch { exit 1 }
}

# --- Main Script ---

try {
    $originalLocation = Get-Location
    if (-not (Test-IsAdmin)) {
        Write-Error "This script must be run with admin privileges. Please run PowerShell as Administrator."
        exit 1
    }

    Set-Location $PSScriptRoot -ErrorAction Stop

    $LatestTableauVersion = Get-LatestTableauVersion
    if (-not $LatestTableauVersion) {
        Write-Error "Exiting script because the latest Tableau version could not be determined."
        exit 1
    }

    $LatestTableauVersionDotted = $LatestTableauVersion -replace '-', '.'

    foreach ($product in $TableauProducts) {

        $ProductName = $product.Name -f $LatestTableauVersionDotted
        $Subfolder = $product.Subfolder
        $InstallerFileName = $product.FileNamePattern -f $LatestTableauVersionDotted
        $ProgramsAndFeaturesNamePrefix = $product.ProgramsAndFeaturesNamePrefix
        $DetectionFile = $product.DetectionFile
        $InstallBatContent = $product.InstallBatContent -f $LatestTableauVersionDotted
        $UninstallBatContent = $product.UninstallBatContent -f $LatestTableauVersionDotted

        Write-Host ""
        Write-Host ("=" * 60)
        Write-Host "Processing: $ProductName"
        Write-Host ("=" * 60)

        $LocalDownloadPath = Join-Path $BaseDownloadRoot $Subfolder
        $LocalInstallerPath = Join-Path $LocalDownloadPath $InstallerFileName

        $NetworkProductRoot = Join-Path $TableauRootNetworkPath $Subfolder
        $NetworkVersionPath = Join-Path $NetworkProductRoot $LatestTableauVersionDotted
        $NetworkInstallerPath = Join-Path $NetworkVersionPath $InstallerFileName

        New-ContentFolder -Path $LocalDownloadPath
        New-ContentFolder -Path $LocalTempRoot
        New-ContentFolder -Path $NetworkProductRoot
        New-ContentFolder -Path $NetworkVersionPath

        $DownloadUrl = "{0}{1}" -f $BaseDownloadUrl, $InstallerFileName

        Download-InstallerIfNeeded -DownloadUrl $DownloadUrl -LocalFilePath $LocalInstallerPath
        Copy-InstallerToNetworkIfNeeded -LocalFilePath $LocalInstallerPath -NetworkFilePath $NetworkInstallerPath

        Write-InstallUninstallBats -FolderPath $NetworkVersionPath -InstallContent $InstallBatContent -UninstallContent $UninstallBatContent

        # Temporary install for metadata extraction (uses same args as BAT wrapper)
        Install-ForMetadataExtraction -InstallerPath $LocalInstallerPath -InstallArgs "/install /quiet /norestart ACCEPTEULA=1 REMOVEINSTALLEDAPP=1 SENDTELEMETRY=0"

        $regInfo = Get-InstalledAppRegistryInfo -DisplayNamePrefix $ProgramsAndFeaturesNamePrefix
        if (-not $regInfo) {
            Write-Error "Could not find installed application in registry after install: $ProgramsAndFeaturesNamePrefix"
            Uninstall-AfterMetadataExtraction -InstallerPath $LocalInstallerPath -UninstallArgs "/uninstall /quiet /norestart"
            exit 1
        }

        $displayName = $regInfo.DisplayName
        $displayVersion = $regInfo.DisplayVersion
        $publisher = $regInfo.Publisher
        $installLocation = $regInfo.InstallLocation

        if (-not $installLocation) {
            $installLocation = "C:\Program Files\Tableau"
        }

        $detectionExe = Join-Path $installLocation $DetectionFile
        $detectionExeVersion = Get-FileVersion -FilePath $detectionExe

        if (-not $detectionExeVersion) {
            Write-Warning "Could not determine file version for detection file: $detectionExe"
            $detectionExeVersion = $displayVersion
        }

        Uninstall-AfterMetadataExtraction -InstallerPath $LocalInstallerPath -UninstallArgs "/uninstall /quiet /norestart"

        $installCommand = "install.bat"
        $uninstallCommand = "uninstall.bat"

        if (-not $publisher) { $publisher = "Tableau" }

        New-MECMApplicationForTableau `
            -AppName $displayName `
            -Publisher $publisher `
            -SoftwareVersion $displayVersion `
            -ContentLocation $NetworkVersionPath `
            -InstallCommand $installCommand `
            -UninstallCommand $uninstallCommand `
            -DetectionExePath $detectionExe `
            -DetectionExeVersion $detectionExeVersion
    }

    Write-Host ""
    Write-Host "Script execution complete."
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    exit 1
}
finally {
    Set-Location $originalLocation -ErrorAction SilentlyContinue
}
