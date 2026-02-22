<#
Vendor: Google
App: Google Chrome Enterprise (x64)
CMName: Google Chrome Enterprise

.SYNOPSIS
    Packages the latest Google Chrome Enterprise (x64) MSI for MECM.

.DESCRIPTION
    Downloads the latest Google Chrome Enterprise x64 MSI from Google's static enterprise
    download URL, stages content to a versioned network folder, and creates an MECM
    Application with Windows Installer (MSI) detection.

    Install:   msiexec.exe /i googlechromestandaloneenterprise64.msi /qn /norestart
    Uninstall: msiexec.exe /x googlechromestandaloneenterprise64.msi /qn /norestart
    Detection: Windows Installer ProductCode + ProductVersion (IsEquals)

    NOTE: Google's enterprise MSI URL always serves the current stable release.
    The version is read from MSI properties after download.

    GetLatestVersionOnly queries Google's VersionHistory API (tiny JSON response)
    and exits without downloading the installer.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM"). The PSDrive is assumed to already exist.

.PARAMETER Comment
    Work order or comment string applied to the MECM application description.

.PARAMETER FileServerPath
    UNC root of the SCCM content share (e.g., "\\fileserver\sccm$").

.PARAMETER GetLatestVersionOnly
    Queries the Chrome VersionHistory API for the current stable version, outputs the
    version string, and exits. No download or MECM changes are made.

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

# --- Configuration ---
$ChromeVersionApiUrl  = "https://versionhistory.googleapis.com/v1/chrome/platforms/win64/channels/stable/versions?order_by=version+desc&pageSize=1"
$MsiDownloadUrl       = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
$MsiFileName          = "googlechromestandaloneenterprise64.msi"
$BaseDownloadRoot     = Join-Path $env:USERPROFILE "Downloads\_AutoPackager"
$NetworkRootPath      = Join-Path $FileServerPath "Applications\Google\Chrome Enterprise"
$Publisher            = "Google LLC"
$EstimatedRuntimeMins = 10
$MaximumRuntimeMins   = 20

# --- Functions ---

function Test-IsAdmin {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Warning "Admin check failed: $($_.Exception.Message)"
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

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Test-NetworkShareAccess {
    param([Parameter(Mandatory)][string]$Path)
    $originalLocation = Get-Location
    try {
        Set-Location C: -ErrorAction Stop
        if (-not (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) {
            Write-Error "Network path does not exist or is inaccessible: $Path"
            return $false
        }
        $tmp = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
        Set-Content -LiteralPath $tmp -Value "test" -Encoding ASCII -ErrorAction Stop
        Remove-Item -LiteralPath $tmp -ErrorAction Stop
        return $true
    }
    catch {
        Write-Error "Network share is not writable: $Path ($($_.Exception.Message))"
        return $false
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
    }
}

function Get-ChromeStableVersion {
    param([switch]$Quiet)
    if (-not $Quiet) { Write-Host "Querying Chrome VersionHistory API: $ChromeVersionApiUrl" }
    $json = (curl.exe -L --fail --silent --show-error $ChromeVersionApiUrl) -join ''
    if ($LASTEXITCODE -ne 0) { throw "Failed to query Chrome version API: $ChromeVersionApiUrl" }
    $data = ConvertFrom-Json $json
    $version = $data.versions[0].version
    if ([string]::IsNullOrWhiteSpace($version)) { throw "Could not parse Chrome stable version from API response." }
    if (-not $Quiet) { Write-Host "Latest Chrome stable version: $version" }
    return $version
}

function Get-MsiPropertyMap {
    param([Parameter(Mandatory)][string]$MsiPath)
    $installer = $null
    $db        = $null
    $view      = $null
    $record    = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @($MsiPath, 0))
        $wanted = @("ProductName", "ProductVersion", "Manufacturer", "ProductCode")
        $map = @{}
        foreach ($p in $wanted) {
            $sql    = "SELECT ``Value`` FROM ``Property`` WHERE ``Property``='$p'"
            $view   = $db.GetType().InvokeMember("OpenView",   "InvokeMethod", $null, $db,   @($sql))
            $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null
            $record = $view.GetType().InvokeMember("Fetch",     "InvokeMethod", $null, $view, $null)
            if ($null -ne $record) {
                $map[$p] = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
            } else {
                $map[$p] = $null
            }
        }
        return $map
    }
    finally {
        foreach ($o in @($record, $view, $db, $installer)) {
            if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) {
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) | Out-Null
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
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

# --- GetLatestVersionOnly mode ---
if ($GetLatestVersionOnly) {
    try {
        $version = Get-ChromeStableVersion -Quiet
        Write-Output $version
        exit 0
    }
    catch {
        Write-Error "Failed to retrieve Chrome version: $($_.Exception.Message)"
        exit 1
    }
}

# --- Main ---
$originalLocation = Get-Location

try {
    if (-not (Test-IsAdmin)) {
        Write-Error "This script must be run as Administrator."
        exit 1
    }

    Set-Location C: -ErrorAction Stop

    Write-Host ""
    Write-Host ("=" * 60)
    Write-Host "Google Chrome Enterprise Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser         : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    Write-Host ("Machine           : {0}"     -f $env:COMPUTERNAME)
    Write-Host "SiteCode          : $SiteCode"
    Write-Host "BaseDownloadRoot  : $BaseDownloadRoot"
    Write-Host "NetworkRootPath   : $NetworkRootPath"
    Write-Host ""

    Ensure-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $NetworkRootPath)) {
        throw "Network root path not accessible: $NetworkRootPath"
    }

    Set-Location C: -ErrorAction Stop

    # 1. Download MSI (always resolves to latest stable)
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Host "Downloading Chrome Enterprise MSI..."
    curl.exe -L --fail --silent --show-error -o $localMsi $MsiDownloadUrl
    if ($LASTEXITCODE -ne 0) { throw "MSI download failed: $MsiDownloadUrl" }
    Write-Host "Downloaded: $localMsi"

    # 2. Extract MSI properties
    Write-Host "Reading MSI properties..."
    $props = Get-MsiPropertyMap -MsiPath $localMsi
    $productName    = $props["ProductName"]
    $productVersion = $props["ProductVersion"]
    $productCode    = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productVersion)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))    { throw "MSI ProductCode missing." }

    Write-Host "ProductName    : $productName"
    Write-Host "ProductVersion : $productVersion"
    Write-Host "ProductCode    : $productCode"
    Write-Host ""

    # 3. Create versioned content folder
    $contentPath = Join-Path $NetworkRootPath $productVersion
    Ensure-Folder -Path $contentPath

    # 4. Copy MSI
    $netMsi = Join-Path $contentPath $MsiFileName
    if (-not (Test-Path -LiteralPath $netMsi)) {
        Write-Host "Copying MSI to network..."
        Copy-Item -LiteralPath $localMsi -Destination $netMsi -Force -ErrorAction Stop
    } else {
        Write-Host "Network MSI already exists. Skipping copy."
    }

    # 5. Write install.bat and uninstall.bat
    $installBatPath   = Join-Path $contentPath "install.bat"
    $uninstallBatPath = Join-Path $contentPath "uninstall.bat"

    if (-not (Test-Path -LiteralPath $installBatPath)) {
        $installBat = @"
@echo off
setlocal
start /wait "" msiexec.exe /i "%~dp0$MsiFileName" /qn /norestart
exit /b 0
"@
        Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
    }

    if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
        $uninstallBat = @"
@echo off
setlocal
start /wait "" msiexec.exe /x "%~dp0$MsiFileName" /qn /norestart
exit /b 0
"@
        Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
    }

    Write-Host "install.bat and uninstall.bat ready."

    # 6. Connect to Configuration Manager
    $appName = "Google Chrome Enterprise $productVersion"
    Write-Host "CM Application Name : $appName"
    Write-Host ""

    if (-not (Connect-CMSite -SiteCode $SiteCode)) {
        throw "Cannot proceed without CM connection."
    }

    # 7. Check for existing application
    $existingApp = Get-CMApplication -Name $appName -ErrorAction SilentlyContinue
    if ($existingApp) {
        Write-Warning "Application '$appName' already exists (CI_ID: $($existingApp.CI_ID)). Exiting."
        exit 1
    }

    # 8. Create application
    Write-Host "Creating application '$appName'..." -ForegroundColor Yellow
    $cmApp = New-CMApplication `
        -Name $appName `
        -Publisher $Publisher `
        -SoftwareVersion $productVersion `
        -LocalizedApplicationName $appName `
        -Description $Comment `
        -AutoInstall $true `
        -ErrorAction Stop

    Write-Host "Application CI_ID: $($cmApp.CI_ID)"

    # 9. MSI detection (ProductCode + exact version)
    $clause = New-CMDetectionClauseWindowsInstaller `
        -ProductCode $productCode `
        -Value `
        -ExpressionOperator IsEquals `
        -ExpectedValue $productVersion

    # 10. Add Script Deployment Type
    Write-Host "Adding deployment type '$appName'..."
    Add-CMScriptDeploymentType `
        -ApplicationName $appName `
        -DeploymentTypeName $appName `
        -ContentLocation $contentPath `
        -InstallCommand "install.bat" `
        -UninstallCommand "uninstall.bat" `
        -InstallationBehaviorType InstallForSystem `
        -LogonRequirementType WhetherOrNotUserLoggedOn `
        -UserInteractionMode Hidden `
        -EstimatedRuntimeMins $EstimatedRuntimeMins `
        -MaximumRuntimeMins $MaximumRuntimeMins `
        -AddDetectionClause @($clause) `
        -ContentFallback $true `
        -SlowNetworkDeploymentMode Download `
        -RebootBehavior NoAction `
        -ErrorAction Stop | Out-Null

    # 11. Revision history cleanup
    Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

    Write-Host "Google Chrome Enterprise $productVersion packaged successfully." -ForegroundColor Green
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}
finally {
    Set-Location $originalLocation -ErrorAction SilentlyContinue
    Write-Host "Restored initial location to: ${originalLocation}"
}
