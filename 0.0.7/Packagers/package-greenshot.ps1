<#
Vendor: Greenshot
App: Greenshot
CMName: Greenshot

.SYNOPSIS
    Automates downloading the latest Greenshot installer and creating a MECM application.
.DESCRIPTION
    Creates a single MECM application for Greenshot, using file-based detection
    and batch file installation.
    Application name matches Programs and Features (e.g., "Greenshot 1.2.10.6").
    Uses static metadata (Publisher: "Greenshot", SoftwareVersion: download version).
    MECM settings: 10-minute max runtime, 5-minute estimated runtime, system installation, no user logon requirement.
    Checks for existing files locally and on network share to skip redundant downloads/copies.
    Uses Add-CMScriptDeploymentType for the detection clause.
.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM"). The PSDrive is assumed to already exist.
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
      - Local administrator
      - Write access to FileServerPath
#>

param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Stop"

# --- Configuration Variables ---
# URL to Greenshot's GitHub API for the latest release
$GitHubApiUrl = "https://api.github.com/repos/greenshot/greenshot/releases/latest"
# Define the download directory
$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads"
$GreenshotRootNetworkPath = Join-Path $FileServerPath "Applications\Greenshot"
# --- Functions ---
function Test-IsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
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
function Get-LatestGreenshotVersion {
    param([switch]$Quiet)
    if (-not $Quiet) { Write-Host "Fetching Greenshot release information from: ${GitHubApiUrl}" }
    try {
        $ApiResponse = Invoke-RestMethod -Uri $GitHubApiUrl -ErrorAction Stop
        $latestVersionWithV = $ApiResponse.tag_name
        $latestVersion = $latestVersionWithV -replace '^v'
        $directDownloadUrl = $null
        foreach ($asset in $ApiResponse.assets) {
            if ($asset.name -like "Greenshot-INSTALLER*.exe") {
                $directDownloadUrl = $asset.browser_download_url
                break
            }
        }
        if (-not $directDownloadUrl) {
            Write-Error "Could not find the installer download link in the GitHub API response."
            exit 1
        }
        Write-Host "Found latest Greenshot version: ${latestVersion}"
        return [PSCustomObject]@{
            Version = $latestVersion
            DownloadUrl = $directDownloadUrl
        }
    }
    catch {
        Write-Error "Failed to get Greenshot version: $($_.Exception.Message)"
        exit 1
    }
}
function Create-BatchFiles {
    param ([string]$NetworkPath, [string]$Version, [string]$FileName)
    $originalLocation = Get-Location
    Write-Host "Current location before batch file creation: ${originalLocation}"
    try {
        Set-Location $PSScriptRoot -ErrorAction Stop
        Write-Host "Set location to script directory for batch file creation: ${PSScriptRoot}"
        $InstallBatContent = @"
@echo off
setlocal
start /wait "" "%~dp0$FileName" /SP- /ALLUSERS /VERYSILENT /SUPPRESSMESSAGEBOXES /NORUN /FORCECLOSEAPPLICATIONS /NORESTART /LOG
exit /b 0
"@
        $UninstallBatContent = @"
@echo off
setlocal
REM Close any active Greenshot
taskkill /IM "Greenshot.exe" /F

start /wait "" "C:\Program Files\Greenshot\unins000.exe" /SILENT
exit /b 0
"@
        $InstallBatPath = Join-Path $NetworkPath "install.bat"
        $UninstallBatPath = Join-Path $NetworkPath "uninstall.bat"
        Set-Content -Path $InstallBatPath -Value $InstallBatContent -Encoding ASCII -ErrorAction Stop
        Set-Content -Path $UninstallBatPath -Value $UninstallBatContent -Encoding ASCII -ErrorAction Stop
        Write-Host "Created install.bat and uninstall.bat at ${NetworkPath}"
    }
    catch {
        Write-Error "Failed to create batch files in '${NetworkPath}': $($_.Exception.Message)"
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
        Write-Host "Restored location to: ${originalLocation}"
    }
}
function Test-NetworkShareAccess {
    param ([string]$Path)
    $originalLocation = Get-Location
    Write-Host "Current location before network share validation: ${originalLocation}"
    try {
        Set-Location $PSScriptRoot -ErrorAction Stop
        Write-Host "Set location to script directory for network share validation: ${PSScriptRoot}"
        if (-not $Path) { Write-Error "Network path is null or empty."; return $false }
        if (-not (Test-Path $Path -ErrorAction Stop)) { Write-Error "Network path '${Path}' does not exist or is inaccessible."; return $false }
        $testFile = Join-Path $Path "test_$(Get-Random).txt"
        Set-Content -Path $testFile -Value "Test" -ErrorAction Stop
        Remove-Item -Path $testFile -ErrorAction Stop
        Write-Host "Network share '${Path}' is accessible and writable."
        return $true
    }
    catch {
        Write-Error "Failed to access network share '${Path}': $($_.Exception.Message)"
        return $false
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
        Write-Host "Restored location to: ${originalLocation}"
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

function New-MECMApplication {
    param (
        [string]$AppName,
        [string]$Version,
        [string]$NetworkPath,
        [string]$FileName,
        [string]$Publisher
    )
    $originalLocation = Get-Location
    Write-Host "Current location before MECM application creation: ${originalLocation}"
    Write-Host "Application Name: ${AppName}"
    try {
        if (-not (Test-IsAdmin)) { Write-Error "Script must be run with admin privileges."; return }
        if (-not (Connect-CMSite -SiteCode $SiteCode)) { Write-Error "Failed to connect to CM site."; return }
        Write-Host "Checking for existing application: ${AppName}"
        $existingApp = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if ($existingApp) {
            Write-Host "Application '${AppName}' already exists. Checking deployment types..."
            $deploymentTypes = Get-CMDeploymentType -ApplicationName $AppName -ErrorAction SilentlyContinue
            if ($deploymentTypes -and $deploymentTypes.Count -gt 0) {
                Write-Warning "Application '${AppName}' already exists with $($deploymentTypes.Count) deployment type(s). Skipping creation."
                return
            } else {
                Write-Host "Application '${AppName}' exists but has no deployment types. Continuing to add deployment type..."
                $app = $existingApp
            }
        } else {
            Write-Host "Creating application: ${AppName}"
            $app = New-CMApplication -Name $AppName `
                -Publisher $Publisher `
                -SoftwareVersion $Version `
                -Description $Comment `
                -LocalizedApplicationName $AppName `
                -ErrorAction Stop
        }
        Create-BatchFiles -NetworkPath $NetworkPath -Version $Version -FileName $FileName
        $detectionClauses = @()
        # File detection for Greenshot (main executable)
        $detectionClause = New-CMDetectionClauseFile -Path "$env:ProgramFiles\Greenshot" -FileName "Greenshot.exe" -Existence
        $detectionClauses += $detectionClause
        $params = @{
            ApplicationName = $AppName
            DeploymentTypeName = "${AppName}"
            InstallCommand = "install.bat"
            ContentLocation = $NetworkPath
            UninstallCommand = "uninstall.bat"
            InstallationBehaviorType = "InstallForSystem"
            LogonRequirementType = "WhetherOrNotUserLoggedOn"
            MaximumRuntimeMins = 20
            EstimatedRuntimeMins = 10
            AddDetectionClause = $detectionClauses
            ErrorAction = "Stop"
            ContentFallback = $true
            SlowNetworkDeploymentMode = "Download"
        }
        Add-CMScriptDeploymentType @params
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$app.CI_ID) -KeepLatest 1
        Write-Host "Created MECM application: ${AppName} with file-based detection"
    }
    catch {
        Write-Error "Failed to create MECM application: $($_.Exception.Message)"
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
        Write-Host "Restored location to: ${originalLocation}"
    }
}
# --- Main Script ---
try {
    if (-not (Test-IsAdmin)) {
        Write-Error "This script must be run with admin privileges. Please run PowerShell as Administrator."
        exit 1
    }
    Set-Location $PSScriptRoot -ErrorAction Stop
    Write-Host "Set initial location to script directory: ${PSScriptRoot}"
    if ($GetLatestVersionOnly) {
        $info = Get-LatestGreenshotVersion -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        return
    }
    $DownloadInfo = Get-LatestGreenshotVersion
    if (-not $DownloadInfo) { exit 1 }
    $Version = $DownloadInfo.Version
    $DownloadUrl = $DownloadInfo.DownloadUrl
    $FileName = [System.IO.Path]::GetFileName($DownloadUrl)
    $DownloadFolderName = "Greenshot_${Version}_installers"
    $DownloadPath = Join-Path $BaseDownloadRoot $DownloadFolderName
    if (-not (Test-Path $DownloadPath)) {
        Write-Host "Creating download directory: ${DownloadPath}"
        New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
    }
    $AppName = "Greenshot $Version"
    $NetworkPath = Join-Path $GreenshotRootNetworkPath $Version
    $Publisher = "Greenshot"
    if (-not (Test-NetworkShareAccess -Path $GreenshotRootNetworkPath)) {
        Write-Error "Network share '${GreenshotRootNetworkPath}' is inaccessible. Skipping '${AppName}'."
        exit 1
    }
    if (-not (Test-Path $NetworkPath)) {
        Write-Host "Creating network directory: ${NetworkPath}"
        New-Item -ItemType Directory -Path $NetworkPath -Force -ErrorAction Stop | Out-Null
    }
    $OutputPath = Join-Path $DownloadPath $FileName
    $NetworkFilePath = Join-Path $NetworkPath $FileName
    if (Test-Path $OutputPath) {
        Write-Host "${FileName} already exists at ${OutputPath}, skipping download."
    }
    else {
        Write-Host "Downloading ${AppName} (${FileName}) from ${DownloadUrl}"
        try {
            curl.exe -L --fail --silent --show-error -o $OutputPath $DownloadUrl
            if ($LASTEXITCODE -ne 0) { throw "Download failed: $DownloadUrl" }
            Write-Host "Downloaded ${AppName} (${FileName}) to ${OutputPath}"
        }
        catch {
            Write-Error "Failed to download ${FileName}: $($_.Exception.Message)"
            exit 1
        }
    }
    if (Test-Path $NetworkFilePath) {
        Write-Host "${FileName} already exists at ${NetworkFilePath}, skipping copy."
    }
    else {
        try {
            Set-Location $PSScriptRoot -ErrorAction Stop
            Write-Host "Copying ${FileName} to ${NetworkPath}"
            Copy-Item -Path $OutputPath -Destination $NetworkPath -Force -ErrorAction Stop
            Write-Host "Copied ${FileName} to ${NetworkPath}"
        }
        catch {
            Write-Error "Failed to copy file to '${NetworkPath}': $($_.Exception.Message)"
            exit 1
        }
    }
    New-MECMApplication -AppName $AppName `
        -Version $Version `
        -NetworkPath $NetworkPath `
        -FileName $FileName `
        -Publisher $Publisher
    Write-Host "Script execution complete."
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}