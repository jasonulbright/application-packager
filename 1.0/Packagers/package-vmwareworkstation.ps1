<#
Vendor: Broadcom
App: VMware Workstation Pro
CMName: VMware Workstation
VendorUrl: https://www.vmware.com/products/desktop-hypervisor/workstation-and-fusion
CPE: cpe:2.3:a:vmware:workstation:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://docs.vmware.com/en/VMware-Workstation-Pro/index.html
DownloadPageUrl: https://www.vmware.com/products/desktop-hypervisor/workstation-and-fusion

.SYNOPSIS
    Packages VMware Workstation Pro for MECM.

.DESCRIPTION
    Stages a manually downloaded VMware Workstation Pro installer, generates
    content wrappers and detection metadata, and creates an MECM Application
    with file-based version detection.

    Supports two-phase operation:
      -StageOnly    Read version from installer, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    NOTE: VMware Workstation must be downloaded manually from the Broadcom
    Support Portal (login required). Place the installer EXE at:
        C:\temp\ap\VMwareWorkstation\VMware-workstation-full-*.exe

    The version is read from the EXE's FileVersionInfo. The installer is a
    custom bootstrapper that wraps an MSI; it supports /s for silent install.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers. Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 60

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Reads the version from the locally placed installer and outputs it.
    Returns exit code 1 if no installer is found.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
    - RBAC permissions to create Applications and Deployment Types
    - Local administrator
    - Write access to FileServerPath
    - VMware Workstation installer manually downloaded from Broadcom
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 60,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$VendorFolder = "Broadcom"
$AppFolder    = "VMware Workstation"

$BaseDownloadRoot = Join-Path $DownloadRoot "VMwareWorkstation"

# --- Functions ---


function Find-LocalInstaller {
    <#
    .SYNOPSIS
        Finds the most recent VMware Workstation installer in the download root.
    #>
    $exes = Get-ChildItem -Path $BaseDownloadRoot -Filter "VMware-workstation-full-*.exe" -File -ErrorAction SilentlyContinue
    if (-not $exes -or $exes.Count -eq 0) { return $null }
    return ($exes | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
}


function Get-InstallerVersion {
    param([string]$ExePath)

    $fi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($ExePath)
    $v = $fi.FileVersion
    if ([string]::IsNullOrWhiteSpace($v)) {
        $v = $fi.ProductVersion
    }
    if ([string]::IsNullOrWhiteSpace($v)) {
        throw "Could not read version from installer: $ExePath"
    }
    return $v.Trim()
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageVMwareWorkstation {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "VMware Workstation Pro - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Find locally placed installer ---
    $installer = Find-LocalInstaller
    if (-not $installer) {
        Write-Log ("No VMware Workstation installer found. Please download manually from:" +
            "`n  https://support.broadcom.com/" +
            "`nand place at:" +
            "`n  {0}\VMware-workstation-full-*.exe" -f $BaseDownloadRoot) -Level ERROR
        exit 1
    }

    $installerFileName = $installer.Name
    $localExe = $installer.FullName

    # --- Read version from EXE ---
    $version = Get-InstallerVersion -ExePath $localExe

    Write-Log "Installer                    : $localExe"
    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $installerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # VMware bootstrapper: /s for silent, /v passes args to inner MSI
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs @"
'/s', '/v"/qn EULAS_AGREED=1 AUTOSOFTWAREUPDATE=0 DATACOLLECTION=0"'
"@ `
        -UninstallCommand "`$env:ProgramFiles\VMware\VMware Workstation\vmware-installer.exe" `
        -UninstallArgs "'/s', '/v`"/qn`"'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $detectionPath = "{0}\VMware\VMware Workstation" -f ${env:ProgramFiles(x86)}

    $appName   = "VMware Workstation Pro $version"
    $publisher = "Broadcom"

    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : vmware.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "vmware.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $false
        }
    }

    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageVMwareWorkstation {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "VMware Workstation Pro - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }
    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()

    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    $localFiles = Get-ChildItem -Path $localContentPath -File -ErrorAction Stop
    foreach ($f in $localFiles) {
        if ($f.Name -eq "stage-manifest.json") { continue }
        $dest = Join-Path $networkContentPath $f.Name
        if (-not (Test-Path -LiteralPath $dest)) {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
            Write-Log "Copied to network            : $($f.Name)"
        }
        else {
            Write-Log "Already on network           : $($f.Name)"
        }
    }

    New-MECMApplicationFromManifest `
        -Manifest $manifest `
        -SiteCode $SiteCode `
        -Comment $Comment `
        -NetworkContentPath $networkContentPath `
        -EstimatedRuntimeMins $EstimatedRuntimeMins `
        -MaximumRuntimeMins $MaximumRuntimeMins
}


# --- Latest-only mode ---
if ($GetLatestVersionOnly) {
    try {
        Initialize-Folder -Path $BaseDownloadRoot
        $installer = Find-LocalInstaller
        if (-not $installer) {
            Write-Error "No VMware Workstation installer found in $BaseDownloadRoot"
            exit 1
        }
        $v = Get-InstallerVersion -ExePath $installer.FullName
        Write-Output $v
        exit 0
    }
    catch {
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "VMware Workstation Pro Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "(Manual download required)"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageVMwareWorkstation
    }
    elseif ($PackageOnly) {
        Invoke-PackageVMwareWorkstation
    }
    else {
        Invoke-StageVMwareWorkstation
        Invoke-PackageVMwareWorkstation
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
