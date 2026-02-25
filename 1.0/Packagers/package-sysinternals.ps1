<#
Vendor: Microsoft Corporation
App: Sysinternals Suite
CMName: Sysinternals Suite
VendorUrl: https://learn.microsoft.com/en-us/sysinternals/

.SYNOPSIS
    Packages Sysinternals Suite for MECM.

.DESCRIPTION
    Downloads the latest Sysinternals Suite ZIP from Microsoft, stages content to
    a versioned local folder with file-existence detection metadata, and creates
    an MECM Application.

    Supports two-phase operation:
      -StageOnly    Download ZIP, resolve version from Chocolatey, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    Sysinternals Suite is a ZIP archive of standalone tools -- there is no
    traditional installer. The install wrapper extracts the ZIP to
    C:\Program Files\Sysinternals. The uninstall wrapper removes the folder.

    The suite uses date-based versioning (e.g. 2026.2.4) tracked via the
    Chocolatey community repository. Individual tools have their own independent
    version numbers.

    NOTE: Each Sysinternals tool shows a EULA dialog on first run. Users can
    suppress this by passing -accepteula to any tool, or an administrator can
    pre-accept by setting registry keys under HKCU:\Software\Sysinternals.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\Sysinternals Suite\<Version>

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
    Queries the Chocolatey API for the latest Sysinternals Suite date-based
    version, outputs the version string, and exits. No MECM changes are made.

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
$ZipDownloadUrl   = "https://download.sysinternals.com/files/SysinternalsSuite.zip"
$ChocolateyApiUrl = "https://community.chocolatey.org/api/v2/FindPackagesById()?`$orderby=Version%20desc&`$top=1&id=%27sysinternals%27"
$ZipFileName      = "SysinternalsSuite.zip"

$VendorFolder = "Microsoft"
$AppFolder    = "Sysinternals Suite"

$BaseDownloadRoot = Join-Path $DownloadRoot "Sysinternals"

# --- Functions ---


function Get-LatestSysinternalsVersion {
    param([switch]$Quiet)

    Write-Log "Chocolatey API URL           : $ChocolateyApiUrl" -Quiet:$Quiet

    try {
        $xml = (curl.exe -L --fail --silent --show-error $ChocolateyApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Chocolatey API." }

        if ($xml -match '<d:Version[^>]*>([^<]+)</d:Version>') {
            $version = $Matches[1].Trim()
        }
        else {
            throw "Could not parse version from Chocolatey API response."
        }

        if ([string]::IsNullOrWhiteSpace($version)) { throw "Empty version in Chocolatey response." }

        Write-Log "Latest Sysinternals version  : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Sysinternals version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageSysinternals {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Sysinternals Suite - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get latest version ---
    $version = Get-LatestSysinternalsVersion
    if (-not $version) { throw "Could not determine latest Sysinternals version." }

    # --- Download ---
    $localZip = Join-Path $BaseDownloadRoot $ZipFileName

    Write-Log "Download URL                 : $ZipDownloadUrl"
    Write-Log "Local ZIP path               : $localZip"
    Write-Log "Version                      : $version"
    Write-Log ""

    # Always re-download since URL is static (always latest) and we have a new version
    Write-Log "Downloading Sysinternals Suite..."
    Invoke-DownloadWithRetry -Url $ZipDownloadUrl -OutFile $localZip

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedZip = Join-Path $localContentPath $ZipFileName
    if (-not (Test-Path -LiteralPath $stagedZip)) {
        Copy-Item -LiteralPath $localZip -Destination $stagedZip -Force -ErrorAction Stop
        Write-Log "Copied ZIP to staged folder  : $stagedZip"
    }
    else {
        Write-Log "Staged ZIP exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # Install: extract ZIP to Program Files, no traditional installer
    $installPs1 = @(
        "`$zipPath = Join-Path `$PSScriptRoot '$ZipFileName'"
        "`$installDir = Join-Path `$env:ProgramFiles 'Sysinternals'"
        "if (-not (Test-Path `$installDir)) { New-Item -Path `$installDir -ItemType Directory -Force | Out-Null }"
        "Expand-Archive -Path `$zipPath -DestinationPath `$installDir -Force"
        "exit 0"
    ) -join "`r`n"

    $uninstallPs1 = @(
        "`$installDir = Join-Path `$env:ProgramFiles 'Sysinternals'"
        "if (Test-Path `$installDir) { Remove-Item -Path `$installDir -Recurse -Force }"
        "exit 0"
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    $detectionPath = "{0}\Sysinternals" -f $env:ProgramFiles

    $appName   = "Sysinternals Suite $version"
    $publisher = "Microsoft Corporation"

    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : procmon.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $ZipFileName
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "procmon.exe"
            PropertyType  = "DateModified"
            Operator      = "GreaterEquals"
            ExpectedValue = (Get-Date).ToString("yyyy-MM-dd")
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

function Invoke-PackageSysinternals {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Sysinternals Suite - PACKAGE phase"
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
        $v = Get-LatestSysinternalsVersion -Quiet
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
    Write-Log "Sysinternals Suite Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ZipDownloadUrl               : $ZipDownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageSysinternals
    }
    elseif ($PackageOnly) {
        Invoke-PackageSysinternals
    }
    else {
        Invoke-StageSysinternals
        Invoke-PackageSysinternals
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
