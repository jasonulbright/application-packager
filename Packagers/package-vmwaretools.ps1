<#
Vendor: Broadcom
App: VMware Tools (x64)
CMName: VMWare Tools
VendorUrl: https://www.vmware.com/products/cloud-infrastructure/desktop-hypervisor/workstation-and-fusion
CPE: cpe:2.3:a:vmware:tools:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://docs.vmware.com/en/VMware-Tools/index.html
DownloadPageUrl: https://packages.vmware.com/tools/releases/

.SYNOPSIS
    Packages VMware Tools (x64) for MECM.

.DESCRIPTION
    Downloads the latest VMware Tools Windows x64 installer from the official
    VMware/Broadcom CDN (using the "latest" release symlink), stages content to
    a versioned local folder with file-version-based detection metadata, and
    creates an MECM Application with file-based detection.
    Detection uses vmtoolsd.exe version >= packaged version in the Program Files
    install path.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    NOTE: The install.bat and uninstall.bat wrappers always exit with code 3010
    to signal a required reboot to MECM.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Broadcom\VMware Tools\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\VMwareTools).
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
    Outputs only the latest available VMware Tools version string and exits.

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
$LatestUrl = "https://packages.vmware.com/tools/releases/latest/windows/x64/"

$VendorFolder = "Broadcom"
$AppFolder    = "VMware Tools"

$BaseDownloadRoot = Join-Path $DownloadRoot "VMwareTools"

# --- Functions ---


function Get-LatestVMwareToolsRelease {
    param([switch]$Quiet)

    Write-Log "Latest installer URL         : $LatestUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $LatestUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch VMware Tools latest listing: $LatestUrl" }

        $fileName = ([regex]::Matches($html, 'href="(VMware-tools-[^"]*-x64\.exe)"') |
            Select-Object -First 1).Groups[1].Value

        if (-not $fileName) {
            throw "Could not find the installer filename in the latest directory."
        }

        if ($fileName -notmatch 'VMware-tools-(\d+\.\d+\.\d+)') {
            throw "Could not extract version from installer filename: $fileName"
        }
        $version = $Matches[1]

        Write-Log "Latest VMware Tools version  : $version" -Quiet:$Quiet
        Write-Log "Installer filename           : $fileName" -Quiet:$Quiet

        return [PSCustomObject]@{
            Version           = $version
            InstallerFileName = $fileName
        }
    }
    catch {
        Write-Log "Failed to get VMware Tools release info: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageVMwareTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "VMware Tools (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $releaseInfo = Get-LatestVMwareToolsRelease
    if (-not $releaseInfo) { throw "Could not resolve VMware Tools release info." }

    $version           = $releaseInfo.Version
    $installerFileName = $releaseInfo.InstallerFileName

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $installerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        $downloadUrl = "{0}{1}" -f $LatestUrl, $installerFileName
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
    # VMware Tools uses /S (silent) /v (pass args to msiexec) with embedded quoted msiexec args
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/S'', ''/v'', ''"/qn REBOOT=R ADDLOCAL=ALL REMOVE=FileIntrospection,NetworkIntrospection"'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $installerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/S'', ''/v'', ''"/qn REBOOT=R REMOVE=ALL"'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent `
        -InstallBatExitCode '3010' `
        -UninstallBatExitCode '3010'

    # --- Write stage manifest ---
    $detectionPath = "{0}\VMware\VMware Tools" -f $env:ProgramFiles

    $appName   = "VMWare Tools $version"
    $publisher = "Broadcom"

    Write-Log ""
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : vmtoolsd.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = '/S /v"/qn REBOOT=R ADDLOCAL=ALL REMOVE=FileIntrospection,NetworkIntrospection"'
        UninstallArgs   = '/S /v"/qn REBOOT=R REMOVE=ALL"'
        RunningProcess  = @("vmtoolsd")
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "vmtoolsd.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
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

function Invoke-PackageVMwareTools {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "VMware Tools (x64) - PACKAGE phase"
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
        $releaseInfo = Get-LatestVMwareToolsRelease -Quiet
        if (-not $releaseInfo) { exit 1 }
        Write-Output $releaseInfo.Version
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
    Write-Log "VMware Tools (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "LatestUrl                    : $LatestUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageVMwareTools
    }
    elseif ($PackageOnly) {
        Invoke-PackageVMwareTools
    }
    else {
        Invoke-StageVMwareTools
        Invoke-PackageVMwareTools
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
