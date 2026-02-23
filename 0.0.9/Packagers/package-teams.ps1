<#
Vendor: Microsoft
App: Microsoft Teams Enterprise (x64)
CMName: Microsoft Teams Enterprise

.SYNOPSIS
    Packages Microsoft Teams Enterprise (system-wide) MSIX for MECM.

.DESCRIPTION
    Downloads teamsbootstrapper.exe and the Teams Enterprise MSIX (x64) from
    Microsoft, stages both to a versioned local folder with script-based
    detection metadata, and creates an MECM Application with a PowerShell
    detection script.
    Detection uses Get-AppxPackage -AllUsers to check the provisioned MSTeams
    package version >= packaged version.

    Install:   teamsbootstrapper.exe -p -o "MSTeams-x64.msix"
    Uninstall: teamsbootstrapper.exe -x

    NOTE: The bootstrapper and MSIX are always re-downloaded to ensure the
    latest version is used. Version is read from AppxManifest.xml inside the
    MSIX.

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
    Content is staged under: <FileServerPath>\Applications\Microsoft\Teams\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Teams).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download bootstrapper and MSIX, generate content
    wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with script-based detection.

.PARAMETER GetLatestVersionOnly
    Downloads the MSIX to a local staging folder, extracts the version from
    AppxManifest.xml, outputs the version string, and exits.

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
$BootstrapperUrl      = "https://go.microsoft.com/fwlink/?linkid=2243204"
$MsixUrl              = "https://go.microsoft.com/fwlink/?linkid=2196106"
$BootstrapperFileName = "teamsbootstrapper.exe"
$MsixFileName         = "MSTeams-x64.msix"
$AppXPackageName      = "MSTeams"

$VendorFolder = "Microsoft"
$AppFolder    = "Teams"

$BaseDownloadRoot = Join-Path $DownloadRoot "Teams"

# --- Functions ---


function Get-MsixVersion {
    param([Parameter(Mandatory)][string]$MsixPath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::OpenRead($MsixPath)
    try {
        $manifestEntry = $zip.Entries |
            Where-Object { $_.FullName -eq "AppxManifest.xml" } |
            Select-Object -First 1

        if (-not $manifestEntry) {
            throw "AppxManifest.xml not found inside MSIX: $MsixPath"
        }

        $reader = [System.IO.StreamReader]::new($manifestEntry.Open())
        try {
            $xmlContent = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $zip.Dispose()
    }

    $xml     = [xml]$xmlContent
    $version = $xml.Package.Identity.Version
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Version attribute missing or empty in AppxManifest.xml."
    }
    return $version
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTeams {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Microsoft Teams Enterprise (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # Always re-download bootstrapper and MSIX to ensure latest version
    $localBootstrapper = Join-Path $BaseDownloadRoot $BootstrapperFileName
    $localMsix         = Join-Path $BaseDownloadRoot $MsixFileName

    Write-Log "Downloading bootstrapper..."
    Invoke-DownloadWithRetry -Url $BootstrapperUrl -OutFile $localBootstrapper

    Write-Log "Downloading MSIX..."
    Invoke-DownloadWithRetry -Url $MsixUrl -OutFile $localMsix

    # --- Get version from MSIX ---
    $version = Get-MsixVersion -MsixPath $localMsix

    Write-Log ""
    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedBootstrapper = Join-Path $localContentPath $BootstrapperFileName
    if (-not (Test-Path -LiteralPath $stagedBootstrapper)) {
        Copy-Item -LiteralPath $localBootstrapper -Destination $stagedBootstrapper -Force -ErrorAction Stop
        Write-Log "Copied bootstrapper to staged: $stagedBootstrapper"
    }
    else {
        Write-Log "Staged bootstrapper exists. Skipping copy."
    }

    $stagedMsix = Join-Path $localContentPath $MsixFileName
    if (-not (Test-Path -LiteralPath $stagedMsix)) {
        Copy-Item -LiteralPath $localMsix -Destination $stagedMsix -Force -ErrorAction Stop
        Write-Log "Copied MSIX to staged        : $stagedMsix"
    }
    else {
        Write-Log "Staged MSIX exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $installContent = (
        ('$bsPath = Join-Path $PSScriptRoot ''{0}''' -f $BootstrapperFileName),
        ('$msixPath = Join-Path $PSScriptRoot ''{0}''' -f $MsixFileName),
        '$proc = Start-Process -FilePath $bsPath -ArgumentList @(''-p'', ''-o'', $msixPath) -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = (
        ('$bsPath = Join-Path $PSScriptRoot ''{0}''' -f $BootstrapperFileName),
        '$proc = Start-Process -FilePath $bsPath -ArgumentList @(''-x'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Detection script ---
    $detectionScript = (
        ('$pkg = Get-AppxPackage -Name "{0}" -AllUsers |' -f $AppXPackageName),
        '    Sort-Object { [version]$_.Version } -Descending |',
        '    Select-Object -First 1',
        ('if ($pkg -and [version]$pkg.Version -ge [version]"' + $version + '") {'),
        '    Write-Output "Installed: $($pkg.Version)"',
        '}'
    ) -join "`r`n"

    # Write detection.ps1 to content folder for reference / manual testing
    $detectionPs1Path = Join-Path $localContentPath "detection.ps1"
    if (-not (Test-Path -LiteralPath $detectionPs1Path)) {
        Set-Content -LiteralPath $detectionPs1Path -Value $detectionScript -Encoding ASCII -ErrorAction Stop
        Write-Log "Created wrapper              : detection.ps1"
    }
    else {
        Write-Log "Wrapper exists, skipped      : detection.ps1"
    }

    # --- Write stage manifest ---
    $appName   = "Microsoft Teams Enterprise - $version"
    $publisher = "Microsoft Corporation"

    Write-Log ""
    Write-Log "Detection method             : PowerShell script (Get-AppxPackage)"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFiles  = @($BootstrapperFileName, $MsixFileName)
        Detection       = @{
            Type           = "Script"
            ScriptLanguage = "PowerShell"
            ScriptText     = $detectionScript
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

function Invoke-PackageTeams {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Microsoft Teams Enterprise (x64) - PACKAGE phase"
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
        Initialize-Folder -Path $BaseDownloadRoot
        $tempMsix = Join-Path $BaseDownloadRoot $MsixFileName
        Invoke-DownloadWithRetry -Url $MsixUrl -OutFile $tempMsix -Quiet
        $version = Get-MsixVersion -MsixPath $tempMsix
        Write-Output $version
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
    Write-Log "Microsoft Teams Enterprise (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "BootstrapperUrl              : $BootstrapperUrl"
    Write-Log "MsixUrl                      : $MsixUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageTeams
    }
    elseif ($PackageOnly) {
        Invoke-PackageTeams
    }
    else {
        Invoke-StageTeams
        Invoke-PackageTeams
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
