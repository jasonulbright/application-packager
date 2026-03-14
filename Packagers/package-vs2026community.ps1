<#
Vendor: Microsoft
App: Visual Studio 2026 Community
CMName: Visual Studio 2026 Community
VendorUrl: https://visualstudio.microsoft.com/
CPE: cpe:2.3:a:microsoft:visual_studio:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/en-us/visualstudio/releases/2026/release-notes
DownloadPageUrl: https://visualstudio.microsoft.com/downloads/

.SYNOPSIS
    Packages Visual Studio 2026 Community offline layout for MECM.

.DESCRIPTION
    Downloads the Visual Studio 2026 Community bootstrapper, creates an offline
    layout with configurable workloads, stages content to a versioned local folder,
    and creates an MECM Application.

    The install is user-interactive (no --quiet flag). Each developer customizes
    their installation through the VS Installer UI. The --noWeb flag ensures the
    installer uses only the offline layout content. The MECM deployment type is
    configured with RequireUserInteraction and OnlyWhenUserLoggedOn so the
    installer UI is visible to the logged-on user.

    Community edition is free for individuals, open-source projects, academic
    research, education, and small organizations (up to 5 users).

    Supports two-phase operation:
      -StageOnly    Download bootstrapper, create offline layout, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 60

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 120

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (Package phase)
    - Local administrator
    - Write access to FileServerPath (Package phase)
    - Internet access (Stage phase)
    - Significant disk space for offline layout (10-50 GB depending on workloads)
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 60,
    [int]$MaximumRuntimeMins = 120,
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

# VS 2026 Community bootstrapper (Stable channel, major version 18.x)
# See: https://learn.microsoft.com/en-us/visualstudio/install/create-a-network-installation-of-visual-studio
$BootstrapperUrl  = "https://aka.ms/vs/stable/vs_community.exe"
$BootstrapperName = "vs_community.exe"

$VendorFolder     = "Microsoft"
$AppFolder        = "Visual Studio 2026 Community"
$BaseDownloadRoot = Join-Path $DownloadRoot "VS2026Community"

# Detection path
$DetectionPath = "{0}\Microsoft Visual Studio\2026\Community\Common7\IDE" -f $env:ProgramFiles
$DetectionExe  = "devenv.exe"

# Workloads included in the offline layout.
# Use @('--all') for everything (~35-50 GB) or specify workloads (~10-15 GB).
$LayoutArgs = @(
    '--all',
    '--lang', 'en-US'
)

# --- Functions ---


function Get-BootstrapperVersion {
    <#
    .SYNOPSIS
        Downloads the VS bootstrapper and reads its file version.
    #>
    param([switch]$Quiet)

    $bootstrapperPath = Join-Path $BaseDownloadRoot $BootstrapperName

    Write-Log "Bootstrapper URL             : $BootstrapperUrl" -Quiet:$Quiet
    Write-Log "Downloading bootstrapper..."  -Quiet:$Quiet
    Invoke-DownloadWithRetry -Url $BootstrapperUrl -OutFile $bootstrapperPath -Quiet:$Quiet

    if (-not (Test-Path -LiteralPath $bootstrapperPath)) {
        throw "Bootstrapper download failed: $bootstrapperPath"
    }

    $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($bootstrapperPath)

    # ProductVersion may be a display string (e.g., "Visual Studio 2026").
    # Prefer FileVersion which is always numeric (e.g., "18.3.11512.155").
    $version = $versionInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($version)) {
        $version = $versionInfo.ProductVersion
    }
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Could not read version from bootstrapper: $bootstrapperPath"
    }

    # Trim any trailing metadata (e.g., "+abcdef" hash suffixes)
    if ($version -match '^([0-9]+\.[0-9]+\.[0-9]+(\.[0-9]+)?)') {
        $version = $Matches[1]
    }

    Write-Log "Bootstrapper version         : $version" -Quiet:$Quiet
    return @{ Version = $version; Path = $bootstrapperPath }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageVS2026Community {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Visual Studio 2026 Community - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download bootstrapper and get version ---
    $bsInfo = Get-BootstrapperVersion
    $version = $bsInfo.Version
    $bootstrapperPath = $bsInfo.Path

    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Create offline layout ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    Write-Log "Creating offline layout (this may take a long time)..."
    Write-Log "Layout path                  : $localContentPath"
    Write-Log "Workload args                : $($LayoutArgs -join ' ')"
    Write-Log ""

    $layoutCmdArgs = @('--layout', $localContentPath) + $LayoutArgs + @('--wait')
    $layoutProc = Start-Process -FilePath $bootstrapperPath -ArgumentList $layoutCmdArgs -Wait -PassThru -NoNewWindow
    if ($layoutProc.ExitCode -ne 0) {
        throw "VS layout creation failed with exit code $($layoutProc.ExitCode)."
    }
    Write-Log "Offline layout creation complete."
    Write-Log ""

    # --- Generate content wrappers ---
    # Install: Interactive, offline only (--noWeb), NO --quiet so user can customize.
    # No -NoNewWindow so the VS Installer window is visible to the user.
    $installPs1 = @(
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $BootstrapperName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''--noWeb'') -Wait -PassThru',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    # Uninstall: Silent
    $uninstallPs1 = @(
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $BootstrapperName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''uninstall'', ''--quiet'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    $appName   = "Visual Studio 2026 Community - $version"
    $publisher = "Microsoft"

    Write-Log "Detection path               : $DetectionPath"
    Write-Log "Detection file               : $DetectionExe"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName                  = $appName
        Publisher                = $publisher
        SoftwareVersion          = $version
        InstallerFile            = $BootstrapperName
        LogonRequirementType     = "OnlyWhenUserLoggedOn"
        RequireUserInteraction   = $true
        Detection                = @{
            Type          = "File"
            FilePath      = $DetectionPath
            FileName      = $DetectionExe
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
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

function Invoke-PackageVS2026Community {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Visual Studio 2026 Community - PACKAGE phase"
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
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    # --- Copy staged content to network (recursive for layout directories) ---
    $items = Get-ChildItem -Path $localContentPath -ErrorAction Stop
    foreach ($item in $items) {
        if ($item.Name -eq "stage-manifest.json") { continue }
        $dest = Join-Path $networkContentPath $item.Name
        if ($item.PSIsContainer) {
            if (-not (Test-Path -LiteralPath $dest)) {
                Copy-Item -Path $item.FullName -Destination $dest -Recurse -Force -ErrorAction Stop
                Write-Log "Copied directory to network  : $($item.Name)"
            }
            else {
                Write-Log "Directory exists on network  : $($item.Name)"
            }
        }
        else {
            if (-not (Test-Path -LiteralPath $dest)) {
                Copy-Item -LiteralPath $item.FullName -Destination $dest -Force -ErrorAction Stop
                Write-Log "Copied to network            : $($item.Name)"
            }
            else {
                Write-Log "Already on network           : $($item.Name)"
            }
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
        Initialize-Folder -Path $BaseDownloadRoot
        $bsInfo = Get-BootstrapperVersion -Quiet
        if (-not $bsInfo.Version) { exit 1 }
        Write-Output $bsInfo.Version
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
    Write-Log "Visual Studio 2026 Community Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "BootstrapperUrl              : $BootstrapperUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageVS2026Community
    }
    elseif ($PackageOnly) {
        Invoke-PackageVS2026Community
    }
    else {
        Invoke-StageVS2026Community
        Invoke-PackageVS2026Community
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
