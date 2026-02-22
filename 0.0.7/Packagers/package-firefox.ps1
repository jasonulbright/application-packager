<#
Vendor: Mozilla
App: Mozilla Firefox (x64)
CMName: Mozilla Firefox

.SYNOPSIS
    Automates downloading the latest Mozilla Firefox x64 MSI and creating an MECM application.

.DESCRIPTION
    Creates one MECM application for Mozilla Firefox x64 with file-version-based detection.
    Downloads the MSI from Mozilla's release servers, copies to the SCCM content share,
    creates install/uninstall batch files, and registers the application in MECM.
    Application name format: "Mozilla Firefox <version> (x64)".

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Work order or comment string applied to the MECM application description.

.PARAMETER FileServerPath
    UNC root of the SCCM content share (e.g., "\\fileserver\sccm$").

.PARAMETER GetLatestVersionOnly
    Outputs only the latest version string and exits.

.NOTES
    Requires:
      - PowerShell 5.1
      - ConfigMgr Admin Console installed (for ConfigurationManager.psd1)
      - RBAC rights to create Applications and Deployment Types
#>

param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Stop"

# --- Configuration ---
$BaseDownloadRoot       = Join-Path $env:USERPROFILE "Downloads"
$FirefoxNetworkRoot     = Join-Path $FileServerPath "Applications\Mozilla\Firefox"
$FirefoxVersionsJsonUrl = "https://product-details.mozilla.org/1.0/firefox_versions.json"
$FirefoxDownloadBase    = "https://releases.mozilla.org/pub/firefox/releases"
$Publisher              = "Mozilla"
$DetectionFolder        = "C:\Program Files\Mozilla Firefox"
$DetectionFile          = "firefox.exe"

# --- Functions ---

function Test-IsAdmin {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Warning "Failed to check admin privileges: $($_.Exception.Message)"
        return $false
    }
}

function Connect-CMSite {
    param([Parameter(Mandatory)][string]$SiteCode)
    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
        Write-Host "Connected to CM site: $SiteCode"
        return $true
    }
    catch {
        Write-Error "Failed to connect to CM site: $($_.Exception.Message)"
        return $false
    }
}

function Test-NetworkShareAccess {
    param([Parameter(Mandatory)][string]$Path)
    $originalLocation = Get-Location
    try {
        if (-not (Test-Path -LiteralPath $Path -ErrorAction Stop)) {
            Write-Error "Network path '$Path' does not exist or is inaccessible."
            return $false
        }
        $testFile = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
        Set-Content -LiteralPath $testFile -Value "Test" -Encoding ASCII -ErrorAction Stop
        Remove-Item -LiteralPath $testFile -ErrorAction Stop
        return $true
    }
    catch {
        Write-Error "Failed to access network share '$Path': $($_.Exception.Message)"
        return $false
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
    }
}

function Get-LatestFirefoxVersion {
    param([switch]$Quiet)
    try {
        $jsonText = (curl.exe -L --fail --silent --show-error $FirefoxVersionsJsonUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Firefox version info: $FirefoxVersionsJsonUrl" }
        $json = ConvertFrom-Json $jsonText
        $version = $json.LATEST_FIREFOX_VERSION
        if ([string]::IsNullOrWhiteSpace($version)) { throw "LATEST_FIREFOX_VERSION field was empty." }
        if (-not $Quiet) { Write-Host "Latest Firefox version: $version" }
        return $version
    }
    catch {
        Write-Error "Failed to retrieve latest Firefox version: $($_.Exception.Message)"
        return $null
    }
}

function Remove-CMApplicationRevisionHistoryByCIId {
    param(
        [Parameter(Mandatory)][UInt32]$CI_ID,
        [UInt32]$KeepLatest = 1
    )
    $history = Get-CMApplicationRevisionHistory -Id $CI_ID -ErrorAction SilentlyContinue
    if (-not $history) { return }
    $revs = @()
    foreach ($h in @($history)) {
        if ($h.PSObject.Properties.Name -contains 'Revision')  { $revs += [UInt32]$h.Revision;   continue }
        if ($h.PSObject.Properties.Name -contains 'CIVersion') { $revs += [UInt32]$h.CIVersion; continue }
    }
    $revs = $revs | Sort-Object -Unique -Descending
    if ($revs.Count -le $KeepLatest) { return }
    foreach ($rev in ($revs | Select-Object -Skip $KeepLatest)) {
        Remove-CMApplicationRevisionHistory -Id $CI_ID -Revision $rev -Force -ErrorAction Stop
    }
}

function New-MECMFirefoxApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$NetworkPath
    )
    $originalLocation = Get-Location
    try {
        if (-not (Connect-CMSite -SiteCode $SiteCode)) {
            Write-Error "Failed to connect to CM site."
            return
        }

        $existingApp = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if ($existingApp) {
            $dts = Get-CMDeploymentType -ApplicationName $AppName -ErrorAction SilentlyContinue
            if ($dts -and $dts.Count -gt 0) {
                Write-Warning "Application '$AppName' already exists with $($dts.Count) deployment type(s). Skipping."
                return
            }
            $cmApp = $existingApp
        }
        else {
            Write-Host "Creating application: $AppName"
            $cmApp = New-CMApplication `
                -Name $AppName `
                -Publisher $Publisher `
                -SoftwareVersion $Version `
                -LocalizedApplicationName $AppName `
                -Description $Comment `
                -ErrorAction Stop
        }

        $detectionClause = New-CMDetectionClauseFile `
            -Path $DetectionFolder `
            -FileName $DetectionFile `
            -PropertyType Version `
            -ExpressionOperator GreaterEquals `
            -ExpectedValue $Version `
            -Value

        $params = @{
            ApplicationName           = $AppName
            DeploymentTypeName        = "$AppName Script DT"
            InstallCommand            = "install.bat"
            UninstallCommand          = "uninstall.bat"
            ContentLocation           = $NetworkPath
            InstallationBehaviorType  = "InstallForSystem"
            LogonRequirementType      = "WhetherOrNotUserLoggedOn"
            MaximumRuntimeMins        = 30
            EstimatedRuntimeMins      = 10
            ContentFallback           = $true
            SlowNetworkDeploymentMode = "Download"
            AddDetectionClause        = $detectionClause
            ErrorAction               = "Stop"
        }

        Add-CMScriptDeploymentType @params | Out-Null
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Host "Created MECM application: $AppName"
    }
    catch {
        Write-Error "Failed to create MECM application: $($_.Exception.Message)"
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
    }
}

# --- Latest-only mode ---
if ($GetLatestVersionOnly) {
    $v = Get-LatestFirefoxVersion -Quiet
    if (-not $v) { exit 1 }
    Write-Output $v
    exit 0
}

# --- Main ---
try {
    if (-not (Test-IsAdmin)) {
        Write-Error "This script must be run as Administrator."
        exit 1
    }

    Set-Location $PSScriptRoot -ErrorAction Stop
    Write-Host "Set initial location to script directory: $PSScriptRoot"

    $Version = Get-LatestFirefoxVersion
    if (-not $Version) { exit 1 }

    $AppName     = "Mozilla Firefox $Version (x64)"
    $MsiFileName = "Firefox Setup $Version.msi"
    $DownloadUrl = "$FirefoxDownloadBase/$Version/win64/en-US/" + ($MsiFileName -replace ' ', '%20')

    $LocalFolder        = Join-Path $BaseDownloadRoot "Firefox_$Version"
    $LocalMsi           = Join-Path $LocalFolder $MsiFileName
    $NetworkVersionPath = Join-Path $FirefoxNetworkRoot $Version
    $NetworkMsi         = Join-Path $NetworkVersionPath $MsiFileName

    if (-not (Test-Path -LiteralPath $LocalFolder)) {
        New-Item -ItemType Directory -Path $LocalFolder -Force | Out-Null
    }

    if (-not (Test-NetworkShareAccess -Path $FirefoxNetworkRoot)) {
        Write-Error "Network share '$FirefoxNetworkRoot' is inaccessible. Exiting."
        exit 1
    }

    if (-not (Test-Path -LiteralPath $NetworkVersionPath)) {
        Write-Host "Creating network directory: $NetworkVersionPath"
        New-Item -ItemType Directory -Path $NetworkVersionPath -Force -ErrorAction Stop | Out-Null
    }

    # Download MSI
    if (-not (Test-Path -LiteralPath $LocalMsi)) {
        Write-Host "Downloading: $DownloadUrl"
        curl.exe -L --fail --silent --show-error -o $LocalMsi $DownloadUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $DownloadUrl" }
        Write-Host "Downloaded: $LocalMsi"
    }
    else {
        Write-Host "MSI already exists locally: $LocalMsi"
    }

    # Copy to network share
    if (-not (Test-Path -LiteralPath $NetworkMsi)) {
        Write-Host "Copying MSI to network share..."
        Copy-Item -LiteralPath $LocalMsi -Destination $NetworkVersionPath -Force -ErrorAction Stop
        Write-Host "Copied to: $NetworkVersionPath"
    }
    else {
        Write-Host "MSI already exists on network share: $NetworkMsi"
    }

    # Create batch files
    $installBat   = Join-Path $NetworkVersionPath "install.bat"
    $uninstallBat = Join-Path $NetworkVersionPath "uninstall.bat"
    Set-Content -LiteralPath $installBat   -Value "start /wait msiexec.exe /i `"%~dp0$MsiFileName`" /qn /norestart" -Encoding ASCII
    Set-Content -LiteralPath $uninstallBat -Value "start /wait msiexec.exe /x `"%~dp0$MsiFileName`" /qn /norestart" -Encoding ASCII
    Write-Host "Created install.bat and uninstall.bat in $NetworkVersionPath"

    New-MECMFirefoxApplication `
        -AppName     $AppName `
        -Version     $Version `
        -NetworkPath $NetworkVersionPath

    Write-Host ""
    Write-Host "Script execution complete."
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    exit 1
}
