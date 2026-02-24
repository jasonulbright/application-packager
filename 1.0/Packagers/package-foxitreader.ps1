<#
Vendor: Foxit Software
App: Foxit PDF Reader (x64)
CMName: Foxit PDF Reader
VendorUrl: https://www.foxit.com/pdf-reader.html

.SYNOPSIS
    Packages Foxit PDF Reader (x64) for MECM.

.DESCRIPTION
    Downloads the latest Foxit PDF Reader x64 EXE from Foxit's CDN, stages
    content to a versioned local folder with file-based version detection
    metadata, and creates an MECM Application with file-based detection.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    NOTE: Foxit's enterprise MSI requires registration. This packager uses the
    free EXE installer (~370MB) with InnoSetup silent flags. The version is
    resolved from the Chocolatey API.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Foxit Software\Foxit PDF Reader\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\FoxitReader).
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
    Queries the Chocolatey API for the latest Foxit Reader version, outputs
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
$ChocolateyApiUrl = "https://community.chocolatey.org/api/v2/Packages()?`$filter=Id%20eq%20%27foxitreader%27%20and%20IsLatestVersion"

$VendorFolder = "Foxit Software"
$AppFolder    = "Foxit PDF Reader"

$BaseDownloadRoot = Join-Path $DownloadRoot "FoxitReader"

# --- Functions ---


function Get-LatestFoxitReaderVersion {
    param([switch]$Quiet)

    Write-Log "Chocolatey API               : foxitreader" -Quiet:$Quiet

    try {
        $xml = (curl.exe -L --fail --silent --show-error $ChocolateyApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Chocolatey API." }

        # Parse version from OData XML response
        if ($xml -match '<d:Version[^>]*>([^<]+)</d:Version>') {
            $version = $Matches[1]
        }
        else {
            throw "Could not parse version from Chocolatey API response."
        }

        Write-Log "Latest Foxit Reader version  : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Foxit Reader version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Get-FoxitDownloadUrl {
    param([string]$Version)

    # CDN URL pattern: major.minor version in path, abbreviated in filename
    # e.g., version 2025.3.0.35737 -> path 2025.3.0, filename FoxitPDFReader20253
    $parts = $Version.Split('.')
    $majorMinorPatch = "{0}.{1}.{2}" -f $parts[0], $parts[1], $parts[2]
    $filenamePart = "{0}{1}" -f $parts[0], $parts[1]

    $url = "https://cdn01.foxitsoftware.com/product/reader/desktop/win/${majorMinorPatch}/FoxitPDFReader${filenamePart}_L10N_Setup_Prom_x64.exe"
    return $url
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageFoxitReader {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Foxit PDF Reader (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $fullVersion = Get-LatestFoxitReaderVersion
    if (-not $fullVersion) { throw "Could not resolve Foxit Reader version." }

    $downloadUrl = Get-FoxitDownloadUrl -Version $fullVersion
    $installerFileName = [System.IO.Path]::GetFileName($downloadUrl)

    # Display version without build number (e.g., "2025.3.0" from "2025.3.0.35737")
    $parts = $fullVersion.Split('.')
    $displayVersion = "{0}.{1}.{2}" -f $parts[0], $parts[1], $parts[2]

    Write-Log "Full version                 : $fullVersion"
    Write-Log "Display version              : $displayVersion"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading Foxit PDF Reader (~370MB)..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $displayVersion
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
        -InstallArgs "'/VERYSILENT', '/NORESTART', '/MERGETASKS=`"!desktopicon`"'" `
        -UninstallCommand 'C:\Program Files\Foxit Software\Foxit PDF Reader\unins000.exe' `
        -UninstallArgs "'/VERYSILENT', '/NORESTART'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $detectionPath = "{0}\Foxit Software\Foxit PDF Reader" -f $env:ProgramFiles

    $appName   = "Foxit PDF Reader $displayVersion (x64)"
    $publisher = "Foxit Software Inc."

    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : FoxitPDFReader.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $displayVersion
        InstallerFile   = $installerFileName
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "FoxitPDFReader.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $displayVersion
            Is64Bit       = $true
        }
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $displayVersion -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageFoxitReader {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Foxit PDF Reader (x64) - PACKAGE phase"
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
        $v = Get-LatestFoxitReaderVersion -Quiet
        if (-not $v) { exit 1 }
        # Output display version (without build number)
        $parts = $v.Split('.')
        $display = "{0}.{1}.{2}" -f $parts[0], $parts[1], $parts[2]
        Write-Output $display
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
    Write-Log "Foxit PDF Reader (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageFoxitReader
    }
    elseif ($PackageOnly) {
        Invoke-PackageFoxitReader
    }
    else {
        Invoke-StageFoxitReader
        Invoke-PackageFoxitReader
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
