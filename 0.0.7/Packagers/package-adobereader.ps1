<#
Vendor: Adobe Inc.
App: Adobe Acrobat (Reader) DC (x64)
CMName: Adobe Acrobat (Reader) DC

.SYNOPSIS
    Packages the latest Adobe Acrobat (Reader) DC (x64) for MECM.

.DESCRIPTION
    Parses Adobe's official release notes page to determine the current Acrobat DC
    version, constructs the enterprise installer URL, downloads the x64 MUI EXE,
    stages content to a versioned network folder, and creates an MECM Application
    with file version-based detection.

    Install:   AcroRdrDCx64{version}_MUI.exe /sAll /rs /rps /msi /qn /norestart
    Uninstall: PowerShell registry lookup (uninstall.ps1)
    Detection: Acrobat.exe file version >= packaged version

    GetLatestVersionOnly fetches only the Adobe release notes page (small HTML)
    to read the current version — no installer download is performed.

    Adobe Acrobat version notation:
      Release notes use format NN.NNN.NNNNN (e.g., 25.001.21223)
      Download URL uses the same parts concatenated (e.g., 2500121223)

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM"). The PSDrive is assumed to already exist.

.PARAMETER Comment
    Work order or comment string applied to the MECM application description.

.PARAMETER FileServerPath
    UNC root of the SCCM content share (e.g., "\\fileserver\sccm$").

.PARAMETER GetLatestVersionOnly
    Parses Adobe's release notes page for the current version, outputs the version
    string, and exits. No download or MECM changes are made.

.NOTES
    Requires:
      - PowerShell 5.1
      - ConfigMgr Admin Console installed (for ConfigurationManager.psd1)
      - RBAC rights to create Applications and Deployment Types
      - Local administrator
      - Write access to FileServerPath

    Acrobat DC installs to: %ProgramFiles%\Adobe\Acrobat DC\Acrobat\
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
$AdobeReleaseNotesUrl = "https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC/index.html"
$AdobeDownloadBase    = "https://ardownload3.adobe.com/pub/adobe/acrobat/win/AcrobatDC"
$BaseDownloadRoot     = Join-Path $env:USERPROFILE "Downloads\_AutoPackager"
$NetworkRootPath      = Join-Path $FileServerPath "Applications\Adobe\Acrobat Reader DC"
$Publisher            = "Adobe Inc."
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

function Get-AdobeAcrobatVersion {
    # Lightweight: parses Adobe's release notes page — no installer download.
    # Version format on page: NN.NNN.NNNNN  (e.g., 25.001.21223)
    param([switch]$Quiet)
    if (-not $Quiet) { Write-Host "Fetching Adobe Acrobat release notes: $AdobeReleaseNotesUrl" }
    $html = (curl.exe -L --fail --silent --show-error $AdobeReleaseNotesUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Adobe release notes: $AdobeReleaseNotesUrl" }
    # First match is the most recent release listed on the page
    $verMatch = [regex]::Match($html, '\b(\d{2}\.\d{3}\.\d{5})\b')
    if (-not $verMatch.Success) { throw "Could not parse Acrobat DC version from release notes page." }
    $version = $verMatch.Groups[1].Value
    if (-not $Quiet) { Write-Host "Latest Acrobat DC version: $version" }
    return $version
}

function Get-AdobeInstallerInfo {
    # Converts the dotted version to the URL format and constructs the download URL.
    # Input:  "25.001.21223"
    # Output: @{ UrlVersion = "2500121223"; FileName = "AcroRdrDCx642500121223_MUI.exe"; Url = "https://..." }
    param([Parameter(Mandatory)][string]$Version)
    $parts      = $Version -split '\.'           # ["25", "001", "21223"]
    $urlVersion = "$($parts[0])$($parts[1])$($parts[2])"  # "2500121223"
    $fileName   = "AcroRdrDCx64${urlVersion}_MUI.exe"
    $url        = "$AdobeDownloadBase/$urlVersion/$fileName"
    return @{ UrlVersion = $urlVersion; FileName = $fileName; Url = $url }
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
        $version = Get-AdobeAcrobatVersion -Quiet
        Write-Output $version
        exit 0
    }
    catch {
        Write-Error "Failed to retrieve Acrobat DC version: $($_.Exception.Message)"
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
    Write-Host "Adobe Acrobat (Reader) DC Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser             : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    Write-Host ("Machine               : {0}"     -f $env:COMPUTERNAME)
    Write-Host "SiteCode              : $SiteCode"
    Write-Host "BaseDownloadRoot      : $BaseDownloadRoot"
    Write-Host "NetworkRootPath       : $NetworkRootPath"
    Write-Host "AdobeReleaseNotesUrl  : $AdobeReleaseNotesUrl"
    Write-Host ""

    Ensure-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $NetworkRootPath)) {
        throw "Network root path not accessible: $NetworkRootPath"
    }

    Set-Location C: -ErrorAction Stop

    # 1. Get current version from Adobe release notes
    $version     = Get-AdobeAcrobatVersion
    $dlInfo      = Get-AdobeInstallerInfo -Version $version
    $fileName    = $dlInfo.FileName
    $downloadUrl = $dlInfo.Url

    Write-Host "Version       : $version"
    Write-Host "URL version   : $($dlInfo.UrlVersion)"
    Write-Host "File name     : $fileName"
    Write-Host "Download URL  : $downloadUrl"
    Write-Host ""

    # 2. Download installer
    $localExe = Join-Path $BaseDownloadRoot $fileName
    if (-not (Test-Path -LiteralPath $localExe)) {
        Write-Host "Downloading $fileName ..."
        curl.exe -L --fail --silent --show-error -o $localExe $downloadUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $downloadUrl" }
        Write-Host "Downloaded: $localExe"
    } else {
        Write-Host "Installer already cached locally: $localExe"
    }

    # 3. Read file version from downloaded EXE for MECM detection
    $exeFileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($localExe).FileVersion
    if ([string]::IsNullOrWhiteSpace($exeFileVersion)) {
        Write-Warning "Could not read FileVersion from EXE; using release notes version for detection."
        $exeFileVersion = $version
    }
    Write-Host "EXE FileVersion (used for detection) : $exeFileVersion"

    # 4. Create versioned content folder
    $contentPath = Join-Path $NetworkRootPath $version
    Ensure-Folder -Path $contentPath

    # 5. Copy installer to network
    $netExe = Join-Path $contentPath $fileName
    if (-not (Test-Path -LiteralPath $netExe)) {
        Write-Host "Copying installer to network..."
        Copy-Item -LiteralPath $localExe -Destination $netExe -Force -ErrorAction Stop
        Write-Host "Copied: $netExe"
    } else {
        Write-Host "Network installer already exists. Skipping copy."
    }

    # 6. Write install.bat
    $installBatPath = Join-Path $contentPath "install.bat"
    $installBat = @"
@echo off
setlocal
start /wait "" "%~dp0$fileName" /sAll /rs /rps /msi /qn /norestart
exit /b 0
"@
    Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop

    # 7. Write uninstall.ps1 — registry lookup finds the product code at runtime
    $uninstallPs1Path = Join-Path $contentPath "uninstall.ps1"
    $uninstallPs1 = @'
$app = Get-ChildItem `
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' `
    -ErrorAction SilentlyContinue |
    Get-ItemProperty -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like 'Adobe Acrobat*' -and $_.UninstallString -match 'msiexec' } |
    Sort-Object DisplayVersion -Descending |
    Select-Object -First 1
if ($app) {
    Start-Process msiexec.exe -ArgumentList "/x `"$($app.PSChildName)`" /qn /norestart" -Wait -NoNewWindow
}
'@
    Set-Content -LiteralPath $uninstallPs1Path -Value $uninstallPs1 -Encoding UTF8 -ErrorAction Stop

    Write-Host "install.bat and uninstall.ps1 created."

    # 8. Connect to Configuration Manager
    $appName = "Adobe Acrobat (Reader) DC $version"
    Write-Host "CM Application Name : $appName"
    Write-Host ""

    if (-not (Connect-CMSite -SiteCode $SiteCode)) {
        throw "Cannot proceed without CM connection."
    }

    # 9. Check for existing application
    $existingApp = Get-CMApplication -Name $appName -ErrorAction SilentlyContinue
    if ($existingApp) {
        Write-Warning "Application '$appName' already exists (CI_ID: $($existingApp.CI_ID)). Exiting."
        exit 1
    }

    # 10. Create application
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

    # 11. File version detection on Acrobat.exe
    $detectionClause = New-CMDetectionClauseFile `
        -Path "$env:ProgramFiles\Adobe\Acrobat DC\Acrobat" `
        -FileName "Acrobat.exe" `
        -Value `
        -PropertyType Version `
        -ExpressionOperator GreaterEquals `
        -ExpectedValue $exeFileVersion `
        -Is64Bit

    # 12. Add Script Deployment Type
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

    # 13. Revision history cleanup
    Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

    Write-Host "Adobe Acrobat (Reader) DC $version packaged successfully." -ForegroundColor Green
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}
finally {
    Set-Location $originalLocation -ErrorAction SilentlyContinue
    Write-Host "Restored initial location to: ${originalLocation}"
}
