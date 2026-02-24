<#
Vendor: TeamViewer
App: TeamViewer Host (x64)
CMName: TeamViewer Host
VendorUrl: https://www.teamviewer.com/

.SYNOPSIS
    Packages TeamViewer Host (x64) EXE for MECM.

.DESCRIPTION
    Downloads the latest TeamViewer Host x64 setup EXE from TeamViewer's static
    download URL, stages content to a versioned local folder with file-based
    version detection metadata, and creates an MECM Application.

    Supports two-phase operation:
      -StageOnly    Download EXE, derive version from file properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    TeamViewer Host is the unattended-access variant of TeamViewer. It is
    deployed to endpoints that need to be remotely managed by IT without a
    user initiating the session. The Full client (package-teamviewer.ps1) is
    deployed to IT staff who initiate support sessions.

    NOTE: The Host MSI is no longer available from TeamViewer's CDN. This
    packager uses the EXE installer with /S silent flag.

    GetLatestVersionOnly downloads the EXE, reads the file version, and exits.
    No lighter-weight version API is available.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\TeamViewer\TeamViewer Host\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\TeamViewerHost).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download EXE, derive version from file
    properties, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Downloads the TeamViewer Host EXE, reads the file version, outputs the
    version string, and exits. No MECM changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
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
$ExeDownloadUrl    = "https://download.teamviewer.com/download/version_15x/TeamViewer_Host_Setup_x64.exe"
$InstallerFileName = "TeamViewer_Host_Setup_x64.exe"

$VendorFolder = "TeamViewer"
$AppFolder    = "TeamViewer Host"

$BaseDownloadRoot = Join-Path $DownloadRoot "TeamViewerHost"

# --- Functions ---


function Get-LatestTeamViewerHostVersion {
    param([switch]$Quiet)

    Write-Log "TeamViewer Host EXE URL      : $ExeDownloadUrl" -Quiet:$Quiet

    try {
        Initialize-Folder -Path $BaseDownloadRoot

        $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
        if (-not (Test-Path -LiteralPath $localExe)) {
            Write-Log "Downloading TeamViewer Host EXE..." -Quiet:$Quiet
            Invoke-DownloadWithRetry -Url $ExeDownloadUrl -OutFile $localExe
        }

        $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe)
        $version = $fileInfo.ProductVersion
        if ([string]::IsNullOrWhiteSpace($version)) {
            $version = $fileInfo.FileVersion
        }
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Cannot read version from TeamViewer Host EXE."
        }

        # Trim any trailing whitespace or extra segments
        $version = $version.Trim()

        Write-Log "Latest TeamViewer Host ver   : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get TeamViewer Host version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTeamViewerHost {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeamViewer Host (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading TeamViewer Host EXE..."
        Invoke-DownloadWithRetry -Url $ExeDownloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # --- Get version from file properties ---
    $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe)
    $version = $fileInfo.ProductVersion
    if ([string]::IsNullOrWhiteSpace($version)) {
        $version = $fileInfo.FileVersion
    }
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Cannot read version from TeamViewer Host EXE."
    }
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
    $installPs1 = @(
        "`$exePath = Join-Path `$PSScriptRoot '$InstallerFileName'"
        "`$proc = Start-Process -FilePath `$exePath -ArgumentList @('/S') -Wait -PassThru -NoNewWindow"
        "exit `$proc.ExitCode"
    ) -join "`r`n"

    $uninstallPs1 = @(
        "`$uninstaller = `"C:\Program Files\TeamViewer\uninstall.exe`""
        "if (-not (Test-Path -LiteralPath `$uninstaller)) { Write-Error 'TeamViewer Host uninstaller not found.'; exit 1 }"
        "`$proc = Start-Process -FilePath `$uninstaller -ArgumentList @('/S') -Wait -PassThru -NoNewWindow"
        "exit `$proc.ExitCode"
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    $detectionPath = "{0}\TeamViewer" -f $env:ProgramFiles

    $appName   = "TeamViewer Host $version"
    $publisher = "TeamViewer"

    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : TeamViewer.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "TeamViewer.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
        }
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageTeamViewerHost {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeamViewer Host (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    # --- Resolve version from local staging ---
    Initialize-Folder -Path $BaseDownloadRoot

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }
    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()

    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    # --- Read manifest ---
    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
    Write-Log ""

    # --- Network share ---
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    # --- Copy staged content to network ---
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

    # --- MECM application ---
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
        $v = Get-LatestTeamViewerHostVersion -Quiet
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
    Write-Log "TeamViewer Host (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ExeDownloadUrl               : $ExeDownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageTeamViewerHost
    }
    elseif ($PackageOnly) {
        Invoke-PackageTeamViewerHost
    }
    else {
        Invoke-StageTeamViewerHost
        Invoke-PackageTeamViewerHost
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
