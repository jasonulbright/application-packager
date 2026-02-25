<#
Vendor: Citrix (Cloud Software Group)
App: Citrix Workspace (LTSR) (x64)
CMName: Citrix Workspace LTSR
VendorUrl: https://www.citrix.com/downloads/workspace-app/

.SYNOPSIS
    Packages Citrix Workspace App Long Term Service Release (LTSR) for MECM.

.DESCRIPTION
    Downloads the latest Citrix Workspace App LTSR from the Citrix CDN,
    stages content to a versioned local folder with ARP-based registry
    detection metadata, and creates an MECM Application.

    Supports two-phase operation:
      -StageOnly    Download installer, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

    The latest LTSR version is resolved via the Chocolatey community API. The
    LTSR download URL changes with each Cumulative Update release; the script
    attempts to resolve it from the Citrix download page. If the automated
    URL resolution fails, place the installer manually in the download root
    folder (C:\temp\ap\CitrixWorkspaceLTSR\CitrixWorkspaceApp.exe).

    Detection uses the fixed ARP registry key CitrixOnlinePluginPackWeb under
    Wow6432Node. The DisplayVersion value contains the 4-part internal version
    (e.g. 25.7.1000.1025).

    Install switches are read from citrix-workspace-switches.json (if present
    in the Packagers folder) to allow enterprise customization of SSO, plugins,
    ADDLOCAL components, store URL, etc.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Citrix\Citrix Workspace (LTSR) (x64)\<Version>

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
    Queries the Chocolatey API for the latest LTSR version, outputs the
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
$ChocolateyApiUrl = "https://community.chocolatey.org/api/v2/FindPackagesById()?`$orderby=Version%20desc&`$top=1&id=%27citrix-workspace-ltsr%27"
$InstallerFileName = "CitrixWorkspaceApp.exe"

# Citrix LTSR download page (scraped for the download link)
$LtsrDownloadPageUrl = "https://www.citrix.com/downloads/workspace-app/workspace-app-for-windows-long-term-service-release/workspace-app-for-windows-LTSR-Latest.html"

$VendorFolder = "Citrix"
$AppFolder    = "Citrix Workspace (LTSR) (x64)"

$BaseDownloadRoot = Join-Path $DownloadRoot "CitrixWorkspaceLTSR"

# ARP detection key (fixed across versions, shared with CR)
$ArpKeyPath = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\CitrixOnlinePluginPackWeb"

# --- Functions ---


function Get-LatestCitrixLTSRVersion {
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

        Write-Log "Latest CWA LTSR version      : $chocoVersion" -Quiet:$Quiet
        return $chocoVersion
    }
    catch {
        Write-Log "Chocolatey lookup failed: $($_.Exception.Message)" -Level WARN

        # Fallback: check for a manually placed installer in the download root
        Write-Log "Checking for local installer to read version..." -Quiet:$Quiet
        $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
        if (Test-Path -LiteralPath $localExe) {
            $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe)
            $version = $fileInfo.FileVersion
            if (-not [string]::IsNullOrWhiteSpace($version)) {
                Write-Log "Local installer version      : $version" -Quiet:$Quiet
                return $version
            }
        }

        Write-Log "Failed to get CWA LTSR version." -Level ERROR
        return $null
    }
}


function Resolve-LtsrDownloadUrl {
    # Try to extract the download URL from the Citrix LTSR download page.
    # Returns the URL string on success, $null on failure.
    Write-Log "Resolving LTSR download URL from Citrix page..."

    try {
        $html = (curl.exe -L --fail --silent --show-error $LtsrDownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch LTSR download page." }

        # Look for a direct download link to CitrixWorkspaceApp.exe
        # Page uses protocol-relative URLs (//downloads.citrix.com/...)
        if ($html -match '(?:https?:)?//downloads\.citrix\.com/\d+/CitrixWorkspaceApp\.exe') {
            $url = $Matches[0]
            # Normalize protocol-relative URL to https
            if ($url.StartsWith('//')) { $url = "https:$url" }
            Write-Log "Resolved LTSR download URL   : $url"
            return $url
        }

        Write-Log "Could not find download URL in page HTML." -Level WARN
        return $null
    }
    catch {
        Write-Log "URL resolution failed: $($_.Exception.Message)" -Level WARN
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

function Invoke-StageCitrixLTSR {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Citrix Workspace (LTSR) (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get latest version ---
    $version = Get-LatestCitrixLTSRVersion
    if (-not $version) { throw "Could not determine latest Citrix Workspace LTSR version." }

    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Download ---
    $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
    Write-Log "Local installer path         : $localExe"

    if (-not (Test-Path -LiteralPath $localExe)) {
        # Try to resolve the LTSR download URL from Citrix page
        $downloadUrl = Resolve-LtsrDownloadUrl
        if (-not $downloadUrl) {
            throw ("LTSR installer not found locally and could not resolve download URL. " +
                   "Please download the LTSR installer manually from:`n" +
                   "  $LtsrDownloadPageUrl`n" +
                   "and place it at:`n" +
                   "  $localExe")
        }

        Write-Log "Downloading Citrix Workspace LTSR..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe
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
        Write-Log "Could not read internal version from EXE; using API version." -Level WARN
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
    $appName   = "Citrix Workspace $version (LTSR) (x64)"
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

function Invoke-PackageCitrixLTSR {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Citrix Workspace (LTSR) (x64) - PACKAGE phase"
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
        $v = Get-LatestCitrixLTSRVersion -Quiet
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
    Write-Log "Citrix Workspace (LTSR) (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "LTSR download page           : $LtsrDownloadPageUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageCitrixLTSR
    }
    elseif ($PackageOnly) {
        Invoke-PackageCitrixLTSR
    }
    else {
        Invoke-StageCitrixLTSR
        Invoke-PackageCitrixLTSR
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
