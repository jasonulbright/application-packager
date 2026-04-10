<#
Vendor: Postman Inc.
App: Postman
CMName: Postman
VendorUrl: https://www.postman.com/
ReleaseNotesUrl: https://www.postman.com/release-notes/postman-app/
DownloadPageUrl: https://www.postman.com/downloads/

.SYNOPSIS
    Packages Postman for MECM as a per-user install.

.DESCRIPTION
    Downloads the latest Postman installer from the static Postman CDN URL,
    stages content to a versioned local folder with file-existence detection
    metadata, and creates an MECM Application configured for per-user deployment.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    This is a USER-level install (not system). The MECM deployment type uses
    InstallForUser behavior and OnlyWhenUserLoggedOn logon requirement.
    The installer is a Squirrel-based package; -s runs it silently.

    The download URL is version-agnostic (always serves latest), so the
    installer is always re-downloaded. The version is read from the EXE's
    FileVersionInfo.

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
    Maximum allowed runtime in minutes. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Downloads the installer, reads its version, outputs the version string,
    and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
    - RBAC permissions to create Applications and Deployment Types
    - Local administrator
    - Write access to FileServerPath
#>

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


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$InstallerUrl      = "https://dl.pstmn.io/download/latest/win64"
$InstallerFileName = "Postman-win64-Setup.exe"

$VendorFolder = "Postman"
$AppFolder    = "Postman"

$BaseDownloadRoot = Join-Path $DownloadRoot "Postman"

# --- Functions ---


function Get-PostmanVersion {
    <#
    .SYNOPSIS
        Downloads the latest Postman installer and reads its version.
    #>
    param([switch]$Quiet)

    Write-Log "Installer URL                : $InstallerUrl" -Quiet:$Quiet

    try {
        Initialize-Folder -Path $BaseDownloadRoot
        $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
        Write-Log "Downloading Postman..." -Quiet:$Quiet
        Invoke-DownloadWithRetry -Url $InstallerUrl -OutFile $localExe

        $fvi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe)
        $version = $fvi.ProductVersion
        if ([string]::IsNullOrWhiteSpace($version)) {
            $version = $fvi.FileVersion
        }
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Could not read version from Postman installer."
        }

        $version = $version.Trim()
        Write-Log "Postman version              : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Postman version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePostman {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Postman - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download (always re-download; URL is version-agnostic) ---
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Downloading Postman..."
    Invoke-DownloadWithRetry -Url $InstallerUrl -OutFile $localExe

    # --- Read version from EXE ---
    $fvi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe)
    $version = $fvi.ProductVersion
    if ([string]::IsNullOrWhiteSpace($version)) { $version = $fvi.FileVersion }
    if ([string]::IsNullOrWhiteSpace($version)) { throw "Could not read version from Postman installer." }
    $version = $version.Trim()

    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # Squirrel installer: -s for silent
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''-s'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        '$updateExe = Join-Path $env:LOCALAPPDATA ''Postman\Update.exe''',
        'if (Test-Path -LiteralPath $updateExe) {',
        '    $proc = Start-Process -FilePath $updateExe -ArgumentList @(''--uninstall'', ''-s'') -Wait -PassThru -NoNewWindow',
        '    exit $proc.ExitCode',
        '}',
        'exit 0'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    $detectionPath = "%LOCALAPPDATA%\Postman"

    $appName   = "Postman $version"
    $publisher = "Postman Inc."

    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : Postman.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName                  = $appName
        Publisher                = $publisher
        SoftwareVersion          = $version
        InstallerFile            = $InstallerFileName
        InstallerType            = "EXE"
        InstallArgs              = "-s"
        UninstallArgs            = "--uninstall -s"
        RunningProcess           = @("Postman")
        InstallationBehaviorType = "InstallForUser"
        LogonRequirementType     = "OnlyWhenUserLoggedOn"
        Detection                = @{
            Type         = "File"
            FilePath     = $detectionPath
            FileName     = "Postman.exe"
            PropertyType = "Existence"
            Is64Bit      = $true
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

function Invoke-PackagePostman {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Postman - PACKAGE phase"
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
    Write-Log "InstallationBehaviorType     : $($manifest.InstallationBehaviorType)"
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
        $ProgressPreference = 'SilentlyContinue'
        $v = Get-PostmanVersion -Quiet
        if (-not $v) { exit 1 }
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
    Write-Log "Postman Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "InstallerUrl                 : $InstallerUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StagePostman
    }
    elseif ($PackageOnly) {
        Invoke-PackagePostman
    }
    else {
        Invoke-StagePostman
        Invoke-PackagePostman
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
