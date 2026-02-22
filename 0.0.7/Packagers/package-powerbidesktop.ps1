<#
Vendor: Microsoft
App: Microsoft Power BI Desktop (x64)
CMName: Microsoft Power BI Desktop

.SYNOPSIS
    Packages the latest Microsoft Power BI Desktop (x64) for MECM.

.DESCRIPTION
    Parses the Microsoft Download Center page to retrieve the current Power BI Desktop
    x64 installer URL and version, downloads the EXE, stages content to a versioned
    network folder, and creates an MECM Application with file version-based detection.

    Install:   PBIDesktopSetup_x64.exe -quiet ACCEPT_EULA=1
    Uninstall: PowerShell registry lookup (uninstall.ps1)
    Detection: PBIDesktop.exe file version >= packaged version

    GetLatestVersionOnly fetches only the Microsoft Download Center page (small HTML)
    to read the current version — no installer download is performed.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM"). The PSDrive is assumed to already exist.

.PARAMETER Comment
    Work order or comment string applied to the MECM application description.

.PARAMETER FileServerPath
    UNC root of the SCCM content share (e.g., "\\fileserver\sccm$").

.PARAMETER GetLatestVersionOnly
    Parses the Microsoft Download Center page for the current version, outputs the
    version string, and exits. No download or MECM changes are made.

.NOTES
    Requires:
      - PowerShell 5.1
      - ConfigMgr Admin Console installed (for ConfigurationManager.psd1)
      - RBAC rights to create Applications and Deployment Types
      - Local administrator
      - Write access to FileServerPath

    Power BI Desktop is released monthly. The installer URL embedded in the download
    page changes with each release; there is no stable fwlink for the x64 EXE.
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
$DownloadPageUrl      = "https://www.microsoft.com/en-us/download/details.aspx?id=58494"
$ExeFileName          = "PBIDesktopSetup_x64.exe"
$BaseDownloadRoot     = Join-Path $env:USERPROFILE "Downloads\_AutoPackager"
$NetworkRootPath      = Join-Path $FileServerPath "Applications\Microsoft\Power BI Desktop"
$Publisher            = "Microsoft Corporation"
$EstimatedRuntimeMins = 15
$MaximumRuntimeMins   = 30

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

function Get-PowerBIDownloadInfo {
    # Lightweight: parses the Microsoft Download Center page — no installer download.
    param([switch]$Quiet)
    if (-not $Quiet) { Write-Host "Fetching Power BI Desktop download page: $DownloadPageUrl" }

    $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch download page: $DownloadPageUrl" }

    $urlMatch = [regex]::Match($html, '"url"\s*:\s*"(https://download\.microsoft\.com/download/[^"]+PBIDesktopSetup_x64\.exe)"')
    if (-not $urlMatch.Success) { throw "Could not locate PBIDesktopSetup_x64.exe URL in download page." }

    $verMatch = [regex]::Match($html, '"[Vv]ersion"\s*:\s*"(\d+\.\d+\.\d+(?:\.\d+)?)"')
    if (-not $verMatch.Success) { throw "Could not parse version from download page." }

    $dlUrl   = $urlMatch.Groups[1].Value
    $version = $verMatch.Groups[1].Value

    if (-not $Quiet) {
        Write-Host "Download URL : $dlUrl"
        Write-Host "Version      : $version"
    }

    return @{ Url = $dlUrl; Version = $version }
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
        $info = Get-PowerBIDownloadInfo -Quiet
        Write-Output $info.Version
        exit 0
    }
    catch {
        Write-Error "Failed to retrieve Power BI Desktop version: $($_.Exception.Message)"
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
    Write-Host "Microsoft Power BI Desktop Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser        : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    Write-Host ("Machine          : {0}"     -f $env:COMPUTERNAME)
    Write-Host "SiteCode         : $SiteCode"
    Write-Host "BaseDownloadRoot : $BaseDownloadRoot"
    Write-Host "NetworkRootPath  : $NetworkRootPath"
    Write-Host "DownloadPageUrl  : $DownloadPageUrl"
    Write-Host ""

    Ensure-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $NetworkRootPath)) {
        throw "Network root path not accessible: $NetworkRootPath"
    }

    Set-Location C: -ErrorAction Stop

    # 1. Get download URL and version from download page
    $info        = Get-PowerBIDownloadInfo
    $version     = $info.Version
    $downloadUrl = $info.Url

    # 2. Download installer
    $localExe = Join-Path $BaseDownloadRoot $ExeFileName
    Write-Host "Downloading $ExeFileName ($version) — 839 MB, please wait..."
    curl.exe -L --fail --silent --show-error -o $localExe $downloadUrl
    if ($LASTEXITCODE -ne 0) { throw "Download failed: $downloadUrl" }
    Write-Host "Downloaded: $localExe"

    # 3. Create versioned content folder
    $contentPath = Join-Path $NetworkRootPath $version
    Ensure-Folder -Path $contentPath

    # 4. Copy installer to network
    $netExe = Join-Path $contentPath $ExeFileName
    if (-not (Test-Path -LiteralPath $netExe)) {
        Write-Host "Copying installer to network..."
        Copy-Item -LiteralPath $localExe -Destination $netExe -Force -ErrorAction Stop
        Write-Host "Copied: $netExe"
    } else {
        Write-Host "Network installer already exists. Skipping copy."
    }

    # 5. Write install.bat
    $installBatPath = Join-Path $contentPath "install.bat"
    $installBat = @"
@echo off
setlocal
start /wait "" "%~dp0$ExeFileName" -quiet ACCEPT_EULA=1
exit /b 0
"@
    Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop

    # 6. Write uninstall.ps1 — registry lookup finds the product code at runtime
    $uninstallPs1Path = Join-Path $contentPath "uninstall.ps1"
    $uninstallPs1 = @'
$app = Get-ChildItem `
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' `
    -ErrorAction SilentlyContinue |
    Get-ItemProperty -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like 'Microsoft Power BI Desktop*' } |
    Sort-Object DisplayVersion -Descending |
    Select-Object -First 1
if ($app) {
    Start-Process msiexec.exe -ArgumentList "/x `"$($app.PSChildName)`" /qn /norestart" -Wait -NoNewWindow
}
'@
    Set-Content -LiteralPath $uninstallPs1Path -Value $uninstallPs1 -Encoding UTF8 -ErrorAction Stop

    Write-Host "install.bat and uninstall.ps1 created."

    # 7. Connect to Configuration Manager
    $appName = "Microsoft Power BI Desktop $version"
    Write-Host "CM Application Name : $appName"
    Write-Host ""

    if (-not (Connect-CMSite -SiteCode $SiteCode)) {
        throw "Cannot proceed without CM connection."
    }

    # 8. Check for existing application
    $existingApp = Get-CMApplication -Name $appName -ErrorAction SilentlyContinue
    if ($existingApp) {
        Write-Warning "Application '$appName' already exists (CI_ID: $($existingApp.CI_ID)). Exiting."
        exit 1
    }

    # 9. Create application
    Write-Host "Creating application '$appName'..." -ForegroundColor Yellow
    $cmApp = New-CMApplication `
        -Name $appName `
        -Publisher $Publisher `
        -SoftwareVersion $version `
        -LocalizedApplicationName $appName `
        -Description $Comment `
        -AutoInstall $true `
        -ErrorAction Stop

    Write-Host "Application CI_ID: $($cmApp.CI_ID)"

    # 10. File version detection on PBIDesktop.exe
    $detectionClause = New-CMDetectionClauseFile `
        -Path "$env:ProgramFiles\Microsoft Power BI Desktop\bin" `
        -FileName "PBIDesktop.exe" `
        -Value `
        -PropertyType Version `
        -ExpressionOperator GreaterEquals `
        -ExpectedValue $version `
        -Is64Bit

    # 11. Add Script Deployment Type
    Write-Host "Adding deployment type '$appName'..."
    Add-CMScriptDeploymentType `
        -ApplicationName $appName `
        -DeploymentTypeName $appName `
        -ContentLocation $contentPath `
        -InstallCommand "install.bat" `
        -UninstallCommand "PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File uninstall.ps1" `
        -InstallationBehaviorType InstallForSystem `
        -LogonRequirementType WhetherOrNotUserLoggedOn `
        -UserInteractionMode Hidden `
        -EstimatedRuntimeMins $EstimatedRuntimeMins `
        -MaximumRuntimeMins $MaximumRuntimeMins `
        -AddDetectionClause @($detectionClause) `
        -ContentFallback $true `
        -SlowNetworkDeploymentMode Download `
        -RebootBehavior NoAction `
        -ErrorAction Stop | Out-Null

    # 12. Revision history cleanup
    Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

    Write-Host "Microsoft Power BI Desktop $version packaged successfully." -ForegroundColor Green
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}
finally {
    Set-Location $originalLocation -ErrorAction SilentlyContinue
    Write-Host "Restored initial location to: ${originalLocation}"
}
