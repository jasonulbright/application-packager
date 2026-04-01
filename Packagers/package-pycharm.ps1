<#
Vendor: JetBrains
App: PyCharm Community Edition (x64)
CMName: PyCharm Community
VendorUrl: https://www.jetbrains.com/pycharm/
CPE: cpe:2.3:a:jetbrains:pycharm:*:*:*:*:community:*:*:*
ReleaseNotesUrl: https://www.jetbrains.com/pycharm/whatsnew/
DownloadPageUrl: https://www.jetbrains.com/pycharm/download/

.SYNOPSIS
    Packages PyCharm Community Edition (x64) for MECM.

.DESCRIPTION
    Queries the JetBrains releases API for the latest PyCharm Community Edition,
    downloads the Windows EXE installer, stages content to a versioned local
    folder with file-based version detection metadata, and creates an MECM
    Application with file-based detection.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\JetBrains\PyCharm Community\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\PyCharm).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, generate content wrappers
    and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Queries the JetBrains API for the latest PyCharm Community version, outputs
    the version string, and exits. No download or MECM changes are made.

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
$JetBrainsApiUrl = "https://data.services.jetbrains.com/products/releases?code=PCC&latest=true&type=release"

$VendorFolder = "JetBrains"
$AppFolder    = "PyCharm Community"

$BaseDownloadRoot = Join-Path $DownloadRoot "PyCharm"

# --- Functions ---


function Get-LatestPyCharmRelease {
    param([switch]$Quiet)

    Write-Log "JetBrains API URL            : $JetBrainsApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $JetBrainsApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch PyCharm release info" }

        $releases = ConvertFrom-Json $json
        $latest = $releases.PCC[0]

        if (-not $latest) { throw "No PyCharm Community releases found in API response." }

        $version = $latest.version
        $build   = $latest.build

        # Find Windows EXE download
        $downloadUrl = $latest.downloads.windows.link
        if (-not $downloadUrl) { throw "No Windows download link found for PyCharm $version." }

        $fileName = "pycharm-community-${version}.exe"

        Write-Log "Latest PyCharm version       : $version (build $build)" -Quiet:$Quiet

        return [PSCustomObject]@{
            Version     = $version
            Build       = $build
            FileName    = $fileName
            DownloadUrl = $downloadUrl
        }
    }
    catch {
        Write-Log "Failed to get PyCharm release info: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StagePyCharm {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PyCharm Community Edition - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestPyCharmRelease
    if (-not $releaseInfo) { throw "Could not resolve PyCharm release info." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.FileName
    $downloadUrl       = $releaseInfo.DownloadUrl

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Download URL                 : $downloadUrl"
        Write-Log ""
        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

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
    # PyCharm uses NSIS installer: /S for silent, /D= for install dir (no quotes, must be last)
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    # Uninstall via the bundled uninstaller
    $uninstallContent = (
        '$uninstaller = ''C:\Program Files\JetBrains\PyCharm Community Edition *\bin\Uninstall.exe''',
        '$uninstPath = (Resolve-Path $uninstaller -ErrorAction SilentlyContinue | Select-Object -First 1).Path',
        'if ($uninstPath) {',
        '    $proc = Start-Process -FilePath $uninstPath -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        '    exit $proc.ExitCode',
        '} else {',
        '    Write-Error "PyCharm uninstaller not found"',
        '    exit 1',
        '}'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Detection: file-based on pycharm64.exe ---
    $detectionPath = "C:\Program Files\JetBrains\PyCharm Community Edition $version\bin"
    $detectionFile = "pycharm64.exe"

    Write-Log ""
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : $detectionFile"
    Write-Log ""

    # --- Write stage manifest ---
    $appName   = "PyCharm Community - $version"
    $publisher = "JetBrains"

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallCommand = "C:\Program Files\JetBrains\PyCharm Community Edition $version\bin\Uninstall.exe"
        UninstallArgs   = "/S"
        RunningProcess  = @("pycharm64")
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = $detectionFile
            PropertyType  = "Existence"
            Is64Bit       = $true
        }
    }

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackagePyCharm {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PyCharm Community Edition - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    # --- Resolve version from local staging ---
    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestPyCharmRelease -Quiet
    if (-not $releaseInfo) { throw "Could not resolve PyCharm version for Package phase." }

    $version          = $releaseInfo.Version
    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    # --- Read manifest ---
    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Type               : $($manifest.Detection.Type)"
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
            Write-Log "Copied                       : $($f.Name)"
        }
        else {
            Write-Log "Exists, skipped              : $($f.Name)"
        }
    }

    Write-Log ""

    # --- MECM application ---
    New-MECMApplicationFromManifest `
        -Manifest $manifest `
        -SiteCode $SiteCode `
        -Comment $Comment `
        -NetworkContentPath $networkContentPath `
        -EstimatedRuntimeMins $EstimatedRuntimeMins `
        -MaximumRuntimeMins $MaximumRuntimeMins

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Package complete"
    Write-Log ("=" * 60)
}


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if ($GetLatestVersionOnly) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $releaseInfo = Get-LatestPyCharmRelease -Quiet
        if (-not $releaseInfo -or -not $releaseInfo.Version) { exit 1 }
        Write-Output $releaseInfo.Version
        exit 0
    }
    catch {
        exit 1
    }
}

$prefs = Get-PackagerPreferences
if ($prefs.CompanyName) {
    Write-Log "Company name                 : $($prefs.CompanyName)"
}

if ($PackageOnly) {
    Invoke-PackagePyCharm
    exit 0
}

$localContentPath = Invoke-StagePyCharm

if (-not $StageOnly) {
    Invoke-PackagePyCharm
}
