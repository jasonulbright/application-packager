<#
Vendor: Citrix (Cloud Software Group)
App: Citrix Workspace (CR) (x64)
CMName: Citrix Workspace CR
VendorUrl: https://www.citrix.com/downloads/workspace-app/

.SYNOPSIS
    Packages Citrix Workspace App Current Release (CR) for MECM.

.DESCRIPTION
    Downloads the latest Citrix Workspace App Current Release (monthly channel)
    from the Citrix CDN, stages content to a versioned local folder with
    ARP-based registry detection metadata, and creates an MECM Application.

    Supports two-phase operation:
      -StageOnly    Download installer, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The latest version is resolved via the Citrix auto-update catalog XML. The
    installer is always available at a permanent CDN URL that serves the
    latest CR build.

    Detection uses the fixed ARP registry key CitrixOnlinePluginPackWeb under
    Wow6432Node. The DisplayVersion value contains the 4-part internal version
    (e.g. 25.11.10.50).

    Install switches are read from citrix-workspace-switches.json (if present
    in the Packagers folder) to allow enterprise customization of SSO, plugins,
    ADDLOCAL components, store URL, etc.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Citrix\Citrix Workspace (CR) (x64)\<Version>

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
    Queries the Citrix catalog for the latest CR version, outputs the version
    string, and exits. No MECM changes are made.

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
$CatalogUrl       = "https://downloadplugins.citrix.com/ReceiverUpdates/Prod/catalog_win.xml"
$InstallerUrl     = "https://downloadplugins.citrix.com/Windows/CitrixWorkspaceApp.exe"
$InstallerFileName = "CitrixWorkspaceApp.exe"

$VendorFolder = "Citrix"
$AppFolder    = "Citrix Workspace (CR) (x64)"

$BaseDownloadRoot = Join-Path $DownloadRoot "CitrixWorkspaceCR"

# ARP detection key (fixed across versions)
$ArpKeyPath = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\CitrixOnlinePluginPackWeb"

# --- Functions ---


function Get-LatestCitrixCRVersion {
    param([switch]$Quiet)

    Write-Log "Citrix catalog URL           : $CatalogUrl" -Quiet:$Quiet

    try {
        [xml]$catalog = (curl.exe -L --fail --silent --show-error $CatalogUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Citrix update catalog." }

        $current = $catalog.Catalog.Installers.Installer | Where-Object { $_.Stream -eq 'Current' }
        if (-not $current) { throw "No 'Current' stream found in catalog." }

        $version = $current.Version
        if ([string]::IsNullOrWhiteSpace($version) -or $version -eq '0.0.0.0') {
            throw "Invalid version in catalog: '$version'"
        }

        Write-Log "Latest CWA CR version        : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get CWA CR version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


function Get-CwaSwitchesConfig {
    $switchesPath = Join-Path $PSScriptRoot "citrix-workspace-switches.json"
    if (-not (Test-Path -LiteralPath $switchesPath)) { return $null }

    try {
        $raw = Get-Content -LiteralPath $switchesPath -Raw -ErrorAction Stop
        return ($raw | ConvertFrom-Json -ErrorAction Stop)
    }
    catch {
        Write-Log "Warning: Could not read CWA switches config: $($_.Exception.Message)" -Level WARN
        return $null
    }
}


function Build-CwaInstallArgs {
    param($Config)

    $args = @('/silent', '/noreboot')

    if (-not $Config) { return $args }

    # Installation options
    if ($Config.Installation.CleanInstall -eq $true)    { $args += '/CleanInstall' }
    if ($Config.Installation.IncludeSSON -eq $true)     { $args += '/includeSSON' }
    if ($Config.Installation.AppProtection -eq $true)   { $args += '/includeappprotection' }

    # Key=Value installation switches
    if ($Config.Installation.EnableSSON -eq $true)          { $args += 'ENABLE_SSON=Yes' }
    if ($Config.Installation.SessionPreLaunch -eq $true)    { $args += 'ENABLEPRELAUNCH=True' }
    if ($Config.Installation.SelfServiceMode -eq $true)     { $args += 'SELFSERVICEMODE=True' }
    elseif ($null -ne $Config.Installation.SelfServiceMode) { $args += 'SELFSERVICEMODE=False' }

    # Store configuration
    if (-not [string]::IsNullOrWhiteSpace($Config.Store.Url)) {
        $storeName = if ([string]::IsNullOrWhiteSpace($Config.Store.Name)) { 'Store' } else { $Config.Store.Name }
        $storeUrl  = $Config.Store.Url.TrimEnd('/')
        if ($storeUrl -notlike '*/discovery') { $storeUrl = "$storeUrl/discovery" }
        $args += ('STORE0="{0};{1};On;{0}"' -f $storeName, $storeUrl)
    }

    # Plugins
    if ($Config.Plugins.MSTeamsPlugin -eq $false)       { $args += 'InstallMSTeamsPlugin=N' }
    if ($Config.Plugins.ZoomPlugin -eq $false)          { $args += 'Installzoomplugin=N' }
    if ($Config.Plugins.WebExPlugin -eq $true)          { $args += 'ADDONS=WebexVDIPlugin' }
    if ($Config.Plugins.UberAgent -eq $true) {
        $args += '/InstallUberAgent'
        if ($Config.Plugins.UberAgentSkipUpgrade -eq $true) { $args += '/SkipUberAgentUpgrade' }
    }
    if ($Config.Plugins.EPAClient -eq $false)           { $args += 'InstallEPAClient=N' }
    if ($Config.Plugins.SessionRecording -eq $true)     { $args += '/InstallSRAgent' }

    # Update & Telemetry
    if (-not [string]::IsNullOrWhiteSpace($Config.UpdateAndTelemetry.AutoUpdateCheck)) {
        $args += ('AutoUpdateCheck={0}' -f $Config.UpdateAndTelemetry.AutoUpdateCheck)
    }
    if ($Config.UpdateAndTelemetry.EnableCEIP -eq $false)    { $args += 'EnableCEIP=False' }
    if ($Config.UpdateAndTelemetry.EnableTracing -eq $false) { $args += 'EnableTracing=false' }

    # Store policy
    if (-not [string]::IsNullOrWhiteSpace($Config.StorePolicy.AllowAddStore)) {
        $args += ('ALLOWADDSTORE={0}' -f $Config.StorePolicy.AllowAddStore)
    }
    if (-not [string]::IsNullOrWhiteSpace($Config.StorePolicy.AllowSavePwd)) {
        $args += ('ALLOWSAVEPWD={0}' -f $Config.StorePolicy.AllowSavePwd)
    }

    # ADDLOCAL components
    if ($Config.Components.Customize -eq $true) {
        $components = @()
        if ($Config.Components.ReceiverInside -ne $false) { $components += 'ReceiverInside' }
        if ($Config.Components.ICA_Client -ne $false)     { $components += 'ICA_Client' }
        if ($Config.Components.AM -ne $false)             { $components += 'AM' }
        if ($Config.Components.SelfService -eq $true)     { $components += 'SelfService' }
        if ($Config.Components.DesktopViewer -eq $true)   { $components += 'DesktopViewer' }
        if ($Config.Components.WebHelper -eq $true)       { $components += 'WebHelper' }
        if ($Config.Components.BCR_Client -eq $true)      { $components += 'BCR_Client' }
        if ($Config.Components.USB -eq $true)             { $components += 'USB' }
        if ($Config.Components.SSON -eq $true)            { $components += 'SSON' }
        if ($components.Count -gt 0) {
            $args += ('ADDLOCAL={0}' -f ($components -join ','))
        }
    }

    return $args
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageCitrixCR {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Citrix Workspace (CR) (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get latest version ---
    $version = Get-LatestCitrixCRVersion
    if (-not $version) { throw "Could not determine latest Citrix Workspace CR version." }

    Write-Log "Download URL                 : $InstallerUrl"
    Write-Log "Version (catalog)            : $version"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Log "Downloading Citrix Workspace CR..."
        Invoke-DownloadWithRetry -Url $InstallerUrl -OutFile $localExe
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    # --- Read internal version from EXE for detection ---
    $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe)
    $internalVersion = $fileInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($internalVersion)) {
        $internalVersion = $fileInfo.ProductVersion
    }
    if ([string]::IsNullOrWhiteSpace($internalVersion)) {
        Write-Log "Could not read internal version from EXE; using catalog version." -Level WARN
        $internalVersion = $version
    }
    Write-Log "Internal version (detection) : $internalVersion"

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedExe = Join-Path $localContentPath $InstallerFileName
    if (-not (Test-Path -LiteralPath $stagedExe)) {
        Copy-Item -LiteralPath $localExe -Destination $stagedExe -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedExe"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $switchesConfig = Get-CwaSwitchesConfig
    $installArgs = Build-CwaInstallArgs -Config $switchesConfig

    $installArgsQuoted = ($installArgs | ForEach-Object { "''{0}''" -f $_ }) -join ', '
    $installContent = @(
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFileName),
        ('$proc = Start-Process -FilePath $exePath -ArgumentList @({0}) -Wait -PassThru -NoNewWindow' -f $installArgsQuoted),
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallContent = @(
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/silent'', ''/uninstall'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    # --- Write stage manifest ---
    $appName   = "Citrix Workspace $version (CR) (x64)"
    $publisher = "Citrix (Cloud Software Group)"

    Write-Log "ARP key path                 : $ArpKeyPath"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $appName
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        Detection       = @{
            Type          = "RegistryKeyValue"
            Hive          = "HKLM"
            KeyPath       = $ArpKeyPath
            ValueName     = "DisplayVersion"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $internalVersion
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

function Invoke-PackageCitrixCR {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Citrix Workspace (CR) (x64) - PACKAGE phase"
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
        $v = Get-LatestCitrixCRVersion -Quiet
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
    Write-Log "Citrix Workspace (CR) (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "CDN URL                      : $InstallerUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageCitrixCR
    }
    elseif ($PackageOnly) {
        Invoke-PackageCitrixCR
    }
    else {
        Invoke-StageCitrixCR
        Invoke-PackageCitrixCR
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
