<#
Vendor: Piriform Software Ltd.
App: CCleaner
CMName: CCleaner
VendorUrl: https://www.ccleaner.com/

.SYNOPSIS
    Packages CCleaner (Free) for MECM.

.DESCRIPTION
    Downloads the latest CCleaner Free slim installer from the CCleaner CDN,
    stages content to a versioned local folder with ARP-based detection metadata,
    and creates an MECM Application.

    Supports two-phase operation:
      -StageOnly    Download installer, resolve version, generate wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The slim variant is used to avoid bundled third-party offers. The version is
    resolved from the Chocolatey community API.

    NOTE: CCleaner 7 may use an online/stub installer from the Avast CDN. If the
    slim download URL returns 404, check the CCleaner builds page for an updated
    offline installer URL.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Piriform\CCleaner\<Version>

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
    Queries the Chocolatey API for the latest CCleaner version, outputs the
    version string, and exits. No MECM changes are made.

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
$ChocolateyApiUrl = "https://community.chocolatey.org/api/v2/FindPackagesById()?`$orderby=Version%20desc&`$top=1&id=%27ccleaner%27"
$DownloadBaseUrl  = "https://download.ccleaner.com/slim"

$VendorFolder = "Piriform"
$AppFolder    = "CCleaner"

$BaseDownloadRoot = Join-Path $DownloadRoot "CCleaner"

# --- Functions ---


function Get-LatestCCleanerVersion {
    param([switch]$Quiet)

    Write-Log "Chocolatey API URL           : $ChocolateyApiUrl" -Quiet:$Quiet

    try {
        $xml = (curl.exe -L --fail --silent --show-error $ChocolateyApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Chocolatey API." }

        if ($xml -match '<d:Version[^>]*>([^<]+)</d:Version>') {
            $chocoVersion = $Matches[1].Trim()
        }
        else {
            throw "Could not parse version from Chocolatey API response."
        }

        if ([string]::IsNullOrWhiteSpace($chocoVersion)) { throw "Empty version in Chocolatey response." }

        Write-Log "Latest CCleaner version      : $chocoVersion" -Quiet:$Quiet
        return $chocoVersion
    }
    catch {
        Write-Log "Failed to get CCleaner version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageCCleaner {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CCleaner - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get latest version ---
    $version = Get-LatestCCleanerVersion
    if (-not $version) { throw "Could not determine latest CCleaner version." }

    # Construct download URL from version
    # Version format examples: "6.29" -> "629", "7.04" -> "704"
    $parts = $version.Split('.')
    $major = $parts[0]
    $minor = if ($parts.Length -ge 2) { $parts[1].PadLeft(2, '0') } else { "00" }
    $verDigits = "$major$minor"
    $installerFileName = "ccsetup${verDigits}_slim.exe"
    $downloadUrl = "$DownloadBaseUrl/$installerFileName"

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading CCleaner..."
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
        -InstallArgs "'/S'" `
        -UninstallCommand "C:\Program Files\CCleaner\uninst.exe" `
        -UninstallArgs "'/S'"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\CCleaner"

    $appName   = "CCleaner $version"
    $publisher = "Piriform Software Ltd."

    Write-Log "ARP Registry Key             : $arpRegistryKey"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $version
            Operator            = "GreaterEquals"
            Is64Bit             = $true
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

function Invoke-PackageCCleaner {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "CCleaner - PACKAGE phase"
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
        $v = Get-LatestCCleanerVersion -Quiet
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
    Write-Log "CCleaner Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "Chocolatey API URL           : $ChocolateyApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageCCleaner
    }
    elseif ($PackageOnly) {
        Invoke-PackageCCleaner
    }
    else {
        Invoke-StageCCleaner
        Invoke-PackageCCleaner
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
