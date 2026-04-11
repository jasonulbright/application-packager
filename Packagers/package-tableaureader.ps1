<#
Vendor: Tableau
App: Tableau Reader (x64)
CMName: Tableau Reader

.SYNOPSIS
    Packages Tableau Reader (x64) for MECM.

.DESCRIPTION
    Downloads the latest Tableau Reader x64 installer from the official
    Tableau download server, stages content to a versioned local folder,
    temporarily installs the product to extract registry metadata and file
    versions, and creates a stage manifest with file-based version detection.
    Detection uses tabreader.exe version from the registry-discovered
    install location.

    Supports two-phase operation:
      -StageOnly    Download, temp install for metadata extraction, generate
                    content wrappers and stage manifest, then uninstall
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Tableau\Reader\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\TableauReader).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download, temp install for metadata extraction,
    generate content wrappers and stage manifest, then uninstall.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Tableau Reader version string and exits.

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
$BaseDownloadUrl = "https://downloads.tableau.com/tssoftwareregistered/"
$ReleaseNotesUrl = "https://www.tableau.com/support/releases"

$VendorFolder = "Tableau"
$AppFolder    = "Reader"

$InstallerPrefix        = "TableauReader-64bit"
$DetectionFileName      = "tabreader.exe"
$RegistryPrefix         = "Tableau 20"
$DisplayNameMustContain = "Reader"

$InstallArgs   = "/install /quiet /norestart ACCEPTEULA=1 SENDTELEMETRY=0"
$UninstallArgs = "/uninstall /quiet /norestart"

$BaseDownloadRoot = Join-Path $DownloadRoot "TableauReader"

# --- Functions ---


function Get-LatestTableauVersion {
    param([switch]$Quiet)

    Write-Log "Tableau release notes URL     : $ReleaseNotesUrl" -Quiet:$Quiet

    try {
        $HtmlContent = (curl.exe -L --fail --silent --show-error $ReleaseNotesUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch release notes: $ReleaseNotesUrl" }

        $versionPattern = '\b(20\d{2})[.-](\d+)(?:[.-](\d+))?\b'
        $regexMatches = [regex]::Matches($HtmlContent, $versionPattern)

        if ($regexMatches.Count -eq 0) {
            throw "No version matches found in release notes."
        }

        $versions = foreach ($m in $regexMatches) {
            $year  = $m.Groups[1].Value
            $minor = $m.Groups[2].Value
            $patch = $m.Groups[3].Value
            if (-not $patch) { $patch = "0" }
            "{0}.{1}.{2}" -f $year, $minor, $patch
        }

        $versions = $versions | Select-Object -Unique

        $latest = $versions | Sort-Object -Descending -Property @{
                Expression = { [int](($_ -split '\.')[0]) }
            }, @{
                Expression = { [int](($_ -split '\.')[1]) }
            }, @{
                Expression = { [int](($_ -split '\.')[2]) }
            } | Select-Object -First 1

        if (-not $latest) {
            throw "Could not determine latest Tableau version."
        }

        Write-Log "Latest Tableau version        : $latest" -Quiet:$Quiet
        return $latest
    }
    catch {
        Write-Log "Failed to get Tableau version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Get-InstalledAppRegistryInfo {
    param(
        [Parameter(Mandatory)][string]$DisplayNamePrefix,
        [Parameter(Mandatory)][string]$DisplayNameMustContain
    )

    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $uninstallPaths) {
        $apps = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object {
                $_.DisplayName -and
                $_.DisplayName -like "$DisplayNamePrefix*" -and
                $_.DisplayName -like "*$DisplayNameMustContain*"
            } |
            Sort-Object -Property DisplayVersion -Descending

        if ($apps -and $apps.Count -gt 0) {
            return $apps | Select-Object -First 1
        }
    }

    return $null
}

function Get-FileVersion {
    param([Parameter(Mandatory)][string]$FilePath)

    if (-not (Test-Path -LiteralPath $FilePath)) { return $null }
    try {
        return (Get-Item -LiteralPath $FilePath).VersionInfo.FileVersion
    }
    catch {
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTableauReader {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Tableau Reader (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestTableauVersion
    if (-not $version) { throw "Could not resolve Tableau Reader version." }

    $installerFileName = "${InstallerPrefix}-${version}.exe"
    $downloadUrl       = "${BaseDownloadUrl}${installerFileName}"

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
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        ('$proc = Start-Process -FilePath $exePath -ArgumentList @(''/install'', ''/quiet'', ''/norestart'', ''ACCEPTEULA=1'', ''SENDTELEMETRY=0'') -Wait -PassThru -NoNewWindow'),
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/uninstall'', ''/quiet'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Temporary install for metadata extraction ---
    Write-Log ""
    Write-Log "Installing temporarily for metadata extraction..."
    Start-Process -FilePath $localExe -ArgumentList $InstallArgs -Wait -NoNewWindow

    $regInfo = Get-InstalledAppRegistryInfo -DisplayNamePrefix $RegistryPrefix -DisplayNameMustContain $DisplayNameMustContain
    if (-not $regInfo) {
        Write-Log "Could not find installed application in registry: $RegistryPrefix (*$DisplayNameMustContain*)" -Level ERROR
        Write-Log "Uninstalling after failed metadata extraction..."
        Start-Process -FilePath $localExe -ArgumentList $UninstallArgs -Wait -NoNewWindow
        exit 1
    }

    $displayName     = $regInfo.DisplayName
    $displayVersion  = $regInfo.DisplayVersion
    $publisher       = $regInfo.Publisher
    $installLocation = $regInfo.InstallLocation

    if (-not $installLocation) {
        $installLocation = "C:\Program Files\Tableau"
    }

    $detectionExe        = Join-Path $installLocation $DetectionFileName
    $detectionExeVersion = Get-FileVersion -FilePath $detectionExe

    if (-not $detectionExeVersion) {
        Write-Log "Could not determine file version for: $detectionExe" -Level WARN
        $detectionExeVersion = $displayVersion
    }

    Write-Log "Uninstalling after metadata extraction..."
    Start-Process -FilePath $localExe -ArgumentList $UninstallArgs -Wait -NoNewWindow

    if (-not $publisher) { $publisher = "Tableau" }

    # --- Write stage manifest ---
    $detectionPath = Split-Path -Path $detectionExe -Parent
    $detectionFile = Split-Path -Path $detectionExe -Leaf

    Write-Log ""
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : $detectionFile"
    Write-Log "Detection version            : $detectionExeVersion"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $displayName
        Publisher       = $publisher
        SoftwareVersion = $displayVersion
        InstallerFile   = $installerFileName
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = $detectionFile
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $detectionExeVersion
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

function Invoke-PackageTableauReader {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Tableau Reader (x64) - PACKAGE phase"
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
    Write-Log "Detection Version            : $($manifest.Detection.ExpectedValue)"
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
        $v = Get-LatestTableauVersion -Quiet
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
    Write-Log "Tableau Reader (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ReleaseNotesUrl              : $ReleaseNotesUrl"
    Write-Log ""

    if ($StageOnly) {
        $contentPath = Invoke-StageTableauReader
        Write-Output $contentPath
    }
    elseif ($PackageOnly) {
        Invoke-PackageTableauReader
    }
    else {
        Invoke-StageTableauReader
        Invoke-PackageTableauReader
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
