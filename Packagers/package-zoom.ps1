<#
Vendor: Zoom Video Communications
App: Zoom Workplace (x64)
CMName: Zoom Workplace
VendorUrl: https://zoom.us/download
CPE: cpe:2.3:a:zoom:zoom:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://support.zoom.com/hc/en/article?id=zm_kb&sysparm_article=KB0060720
DownloadPageUrl: https://zoom.us/download

.SYNOPSIS
    Packages Zoom Workplace (x64) for MECM as a per-user install.

.DESCRIPTION
    Downloads the latest Zoom Workplace EXE installer from zoom.us, downloads
    the CleanZoom utility for uninstall, stages content to a versioned local
    folder with file-existence detection metadata, and creates an MECM
    Application configured for per-user deployment.

    This is a USER-level install (not system). The MECM deployment type uses
    InstallForUser behavior and OnlyWhenUserLoggedOn logon requirement.
    Detection checks for Zoom.exe in the user's AppData\Roaming path.

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
    Content is staged under: <FileServerPath>\Applications\Zoom Video Communications\Zoom Workplace\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Zoom).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer and CleanZoom, generate
    content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Zoom version string and exits.

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
$InstallerUrl  = "https://zoom.us/client/latest/ZoomInstaller.exe"
$CleanZoomUrl  = "https://assets.zoom.us/docs/msi-templates/CleanZoom.zip"

$VendorFolder = "Zoom Video Communications"
$AppFolder    = "Zoom Workplace"

$BaseDownloadRoot   = Join-Path $DownloadRoot "Zoom"
$InstallerFileName  = "ZoomInstaller.exe"
$CleanZoomFileName  = "CleanZoom.exe"

# --- Functions ---


function Get-LatestZoomVersion {
    <#
    .SYNOPSIS
        Resolves the latest Zoom Workplace version by following the redirect
        from the "latest" download URL and extracting the version from the
        effective URL path. Falls back to downloading the installer and
        reading its ProductVersion file property.
    #>
    param([switch]$Quiet)

    Write-Log "Zoom installer URL           : $InstallerUrl" -Quiet:$Quiet

    try {
        # Follow redirects, discard body, capture effective URL
        $effectiveUrl = & curl.exe -s -L -o NUL -w "%{url_effective}" $InstallerUrl 2>$null
        if ($LASTEXITCODE -ne 0) { throw "curl failed to follow redirect." }

        Write-Log "Effective URL                : $effectiveUrl" -Quiet:$Quiet

        # Expected: https://zoom.us/client/6.6.0/ZoomInstaller.exe (or 4-part like 6.5.1.6476)
        if ($effectiveUrl -match '/client/(\d+\.\d+\.\d+(?:\.\d+)?)/') {
            $version = $Matches[1]
            Write-Log "Zoom version (from URL)      : $version" -Quiet:$Quiet
            return $version
        }

        # Fallback: download installer and read FileVersionInfo
        Write-Log "URL does not contain version. Downloading installer to extract version..." -Level WARN -Quiet:$Quiet

        Initialize-Folder -Path $BaseDownloadRoot
        $tempExe = Join-Path $BaseDownloadRoot $InstallerFileName
        Invoke-DownloadWithRetry -Url $InstallerUrl -OutFile $tempExe -Quiet:$Quiet

        $fvi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($tempExe)
        $version = $fvi.ProductVersion
        if ([string]::IsNullOrWhiteSpace($version)) { throw "Could not read ProductVersion from installer." }

        Write-Log "Zoom version (from EXE)      : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to resolve Zoom version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageZoom {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Zoom Workplace (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $version = Get-LatestZoomVersion
    if (-not $version) { throw "Could not resolve Zoom version." }

    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Download installer ---
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading Zoom installer..."
        Invoke-DownloadWithRetry -Url $InstallerUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # --- Download CleanZoom ---
    $cleanZoomZip = Join-Path $BaseDownloadRoot "CleanZoom.zip"
    $cleanZoomExe = Join-Path $BaseDownloadRoot $CleanZoomFileName

    if (-not (Test-Path -LiteralPath $cleanZoomExe)) {
        Write-Log "Downloading CleanZoom utility..."
        Invoke-DownloadWithRetry -Url $CleanZoomUrl -OutFile $cleanZoomZip

        # Extract CleanZoom.exe from the ZIP
        Write-Log "Extracting CleanZoom.exe..."
        $extractPath = Join-Path $BaseDownloadRoot "CleanZoom_extract"
        if (Test-Path -LiteralPath $extractPath) {
            Remove-Item -LiteralPath $extractPath -Recurse -Force
        }
        Expand-Archive -LiteralPath $cleanZoomZip -DestinationPath $extractPath -Force

        $extracted = Get-ChildItem -Path $extractPath -Filter "CleanZoom.exe" -Recurse | Select-Object -First 1
        if (-not $extracted) { throw "CleanZoom.exe not found in ZIP archive." }

        Copy-Item -LiteralPath $extracted.FullName -Destination $cleanZoomExe -Force -ErrorAction Stop
        Remove-Item -LiteralPath $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $cleanZoomZip -Force -ErrorAction SilentlyContinue

        Write-Log "CleanZoom.exe extracted       : $cleanZoomExe"
    }
    else {
        Write-Log "CleanZoom.exe exists. Skipping download."
    }
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    # Copy installer
    $stagedExe = Join-Path $localContentPath $InstallerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied installer to staged   : $stagedExe"
    }
    else {
        Write-Log "Staged installer exists. Skipping copy."
    }

    # Copy CleanZoom
    $stagedCleanZoom = Join-Path $localContentPath $CleanZoomFileName
    if (-not (Test-Path -LiteralPath $stagedCleanZoom)) {
        Copy-Item -LiteralPath $cleanZoomExe -Destination $stagedCleanZoom -Force -ErrorAction Stop
        Write-Log "Copied CleanZoom to staged   : $stagedCleanZoom"
    }
    else {
        Write-Log "Staged CleanZoom exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $installContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/silent'', ''/install'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $CleanZoomFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/silent'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    # Per-user install: detection via file existence in user's AppData
    $detectionPath = "%APPDATA%\Zoom\bin"

    $appName   = "Zoom Workplace - $version (x64)"
    $publisher = "Zoom Video Communications"

    Write-Log ""
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : Zoom.exe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName                  = $appName
        Publisher                = $publisher
        SoftwareVersion          = $version
        InstallerFile            = $InstallerFileName
        InstallerType            = "EXE"
        InstallArgs              = "/silent /install"
        UninstallArgs            = "/silent"
        RunningProcess           = @("Zoom")
        InstallationBehaviorType = "InstallForUser"
        LogonRequirementType     = "OnlyWhenUserLoggedOn"
        Detection                = @{
            Type         = "File"
            FilePath     = "%APPDATA%\Zoom\bin"
            FileName     = "Zoom.exe"
            PropertyType = "Existence"
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

function Invoke-PackageZoom {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Zoom Workplace (x64) - PACKAGE phase"
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
    Write-Log "InstallationBehaviorType     : $($manifest.InstallationBehaviorType)"
    Write-Log "LogonRequirementType         : $($manifest.LogonRequirementType)"
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
        $ver = Get-LatestZoomVersion -Quiet
        if (-not $ver) { exit 1 }
        Write-Output $ver
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
    Write-Log "Zoom Workplace (x64) Auto-Packager starting"
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
        Invoke-StageZoom
    }
    elseif ($PackageOnly) {
        Invoke-PackageZoom
    }
    else {
        Invoke-StageZoom
        Invoke-PackageZoom
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
