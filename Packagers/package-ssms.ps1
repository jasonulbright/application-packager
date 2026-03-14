<#
Vendor: Microsoft
App: SQL Server Management Studio
CMName: SQL Server Management Studio
VendorUrl: https://learn.microsoft.com/sql/ssms/download-sql-server-management-studio-ssms
CPE: cpe:2.3:a:microsoft:sql_server_management_studio:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/en-us/sql/ssms/release-notes-ssms
DownloadPageUrl: https://learn.microsoft.com/en-us/sql/ssms/download-sql-server-management-studio-ssms

.SYNOPSIS
    Packages SQL Server Management Studio (SSMS) for MECM.

.DESCRIPTION
    Downloads the latest SSMS bootstrapper from Microsoft, creates an offline
    layout using the VS Installer backend (--layout), stages content to a
    versioned local folder, and creates an MECM Application with silent install.

    SSMS 22+ uses the Visual Studio Installer bootstrapper. The offline layout
    ensures installation works without internet connectivity on endpoints.

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
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 60

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
    - Disk space for offline layout (~1-2 GB)
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
$SsmsDownloadUrl  = "https://aka.ms/ssms/22/release/vs_SSMS.exe"
$BootstrapperName = "vs_SSMS.exe"

$VendorFolder     = "Microsoft"
$AppFolder        = "SQL Server Management Studio"
$BaseDownloadRoot = Join-Path $DownloadRoot "SSMS"

# --- Functions ---


function Get-SsmsBootstrapperVersion {
    <#
    .SYNOPSIS
        Downloads the SSMS bootstrapper and reads its file version.
    #>
    param([switch]$Quiet)

    $bootstrapperPath = Join-Path $BaseDownloadRoot $BootstrapperName

    Write-Log "SSMS download URL            : $SsmsDownloadUrl" -Quiet:$Quiet
    Write-Log "Downloading SSMS bootstrapper..." -Quiet:$Quiet
    Invoke-DownloadWithRetry -Url $SsmsDownloadUrl -OutFile $bootstrapperPath -Quiet:$Quiet

    if (-not (Test-Path -LiteralPath $bootstrapperPath)) {
        throw "SSMS bootstrapper download failed: $bootstrapperPath"
    }

    $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($bootstrapperPath)

    # ProductVersion contains the SSMS version (e.g. 22.3.0).
    # FileVersion contains the VS Installer engine version - do not use.
    $version = $versionInfo.ProductVersion
    if ([string]::IsNullOrWhiteSpace($version)) {
        $version = $versionInfo.FileVersion
    }
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Could not read version from SSMS bootstrapper: $bootstrapperPath"
    }

    # Trim any trailing metadata (e.g., "+abcdef" hash suffixes)
    if ($version -match '^([0-9]+\.[0-9]+(\.[0-9]+(\.[0-9]+)?)?)') {
        $version = $Matches[1]
    }

    Write-Log "SSMS version                 : $version" -Quiet:$Quiet
    return @{ Version = $version; Path = $bootstrapperPath }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageSsms {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SQL Server Management Studio - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Download bootstrapper and get version ---
    $bsInfo = Get-SsmsBootstrapperVersion
    $version = $bsInfo.Version
    $bootstrapperPath = $bsInfo.Path

    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Create offline layout ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    Write-Log "Creating offline layout (this may take several minutes)..."
    Write-Log "Layout path                  : $localContentPath"
    Write-Log ""

    $layoutProc = Start-Process -FilePath $bootstrapperPath -ArgumentList @('--layout', $localContentPath, '--lang', 'en-US', '--wait') -Wait -PassThru -NoNewWindow
    if ($layoutProc.ExitCode -ne 0) {
        throw "SSMS layout creation failed with exit code $($layoutProc.ExitCode)."
    }
    Write-Log "Offline layout creation complete."
    Write-Log ""

    # --- Generate content wrappers ---
    # Install: Silent, runs from layout
    $installPs1 = @(
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $BootstrapperName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''--quiet'', ''--norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    # Uninstall: Silent
    $uninstallPs1 = @(
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $BootstrapperName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''uninstall'', ''--quiet'', ''--norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Determine detection ---
    # SSMS registers in ARP under a product-specific GUID.
    # Use the bootstrapper's ProductName and version for file-based detection instead,
    # since the GUID changes between major versions.
    # SSMS installs ssms.exe in a known path.
    $detectionPath = "{0}\Microsoft SQL Server Management Studio {1}\Release\Common7\IDE" -f $env:ProgramFiles, ($version -replace '\..*$', '')
    $detectionFile = "Ssms.exe"

    # If the major-version folder naming doesn't match, fall back to a broader check.
    # The detection path format may need adjustment based on the actual SSMS install path.
    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : $detectionFile"
    Write-Log ""

    # --- Write stage manifest ---
    $appName   = "SQL Server Management Studio - $version"
    $publisher = "Microsoft"

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $BootstrapperName
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = $detectionFile
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

function Invoke-PackageSsms {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SQL Server Management Studio - PACKAGE phase"
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
        $bsInfo = Get-SsmsBootstrapperVersion -Quiet
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
    Write-Log "SQL Server Management Studio Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "SsmsDownloadUrl              : $SsmsDownloadUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageSsms
    }
    elseif ($PackageOnly) {
        Invoke-PackageSsms
    }
    else {
        Invoke-StageSsms
        Invoke-PackageSsms
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
