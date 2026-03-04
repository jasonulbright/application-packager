<#
Vendor: Opera Software
App: Opera Browser (x64)
CMName: Opera Browser
VendorUrl: https://www.opera.com/
CPE: cpe:2.3:a:opera:opera_browser:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://blogs.opera.com/desktop/changelog/
DownloadPageUrl: https://www.opera.com/download

.SYNOPSIS
    Packages Opera Browser (x64) for MECM.

.DESCRIPTION
    Downloads the latest Opera Browser x64 offline installer from the Opera CDN,
    stages content to a versioned local folder with file-based version detection
    metadata, and creates an MECM Application.

    Supports two-phase operation:
      -StageOnly    Download installer, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The latest version is resolved by scraping the Opera CDN desktop directory
    listing for the latest versioned folder.

    Opera's ARP registry key changes with every version update, so file-based
    detection on opera.exe is used instead.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Opera Software\Opera Browser (x64)\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Scrapes the Opera CDN for the latest version, outputs the version string,
    and exits. No MECM changes are made.

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
$CdnDirectoryUrl = "https://get.geo.opera.com/pub/opera/desktop/"

$VendorFolder = "Opera Software"
$AppFolder    = "Opera Browser (x64)"

$BaseDownloadRoot = Join-Path $DownloadRoot "Opera"

# --- Functions ---


function Get-LatestOperaVersion {
    param([switch]$Quiet)

    Write-Log "CDN directory URL            : $CdnDirectoryUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $CdnDirectoryUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Opera CDN directory listing." }

        # Parse version folders from directory listing (4-part versions like 127.0.5778.76)
        $versions = [regex]::Matches($html, 'href="(\d+\.\d+\.\d+\.\d+)/"') |
            ForEach-Object { $_.Groups[1].Value } |
            Sort-Object { [version]$_ } -Descending

        if (-not $versions -or $versions.Count -eq 0) {
            throw "No version folders found in CDN directory listing."
        }

        $version = $versions[0]
        $fileName = "Opera_${version}_Setup_x64.exe"
        $downloadUrl = "${CdnDirectoryUrl}${version}/win/$fileName"

        Write-Log "Latest Opera version         : $version" -Quiet:$Quiet

        return [PSCustomObject]@{
            Version     = $version
            FileName    = $fileName
            DownloadUrl = $downloadUrl
        }
    }
    catch {
        Write-Log "Failed to get Opera version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageOpera {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Opera Browser (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get latest version ---
    $versionInfo = Get-LatestOperaVersion
    if (-not $versionInfo) { throw "Could not determine latest Opera version." }

    $version          = $versionInfo.Version
    $installerFileName = $versionInfo.FileName
    $downloadUrl      = $versionInfo.DownloadUrl

    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading Opera..."
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
    $wrapperContent = New-ExeWrapperContent `
        -InstallerFileName $installerFileName `
        -InstallArgs "'/silent', '/allusers=1', '/launchopera=0', '/setdefaultbrowser=0'" `
        -UninstallCommand "C:\Program Files\Opera\launcher.exe" `
        -UninstallArgs "'--uninstall', '/silent'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    # Opera's ARP key changes with every version, so use file-based detection
    $detectionPath = "{0}\Opera" -f $env:ProgramFiles

    $appName   = "Opera Browser $version (x64)"
    $publisher = "Opera Software"

    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : opera.exe"
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
            FileName      = "opera.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
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

function Invoke-PackageOpera {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Opera Browser (x64) - PACKAGE phase"
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
        $ProgressPreference = 'SilentlyContinue'
        $info = Get-LatestOperaVersion -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
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
    Write-Log "Opera Browser (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "CdnDirectoryUrl              : $CdnDirectoryUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageOpera
    }
    elseif ($PackageOnly) {
        Invoke-PackageOpera
    }
    else {
        Invoke-StageOpera
        Invoke-PackageOpera
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
