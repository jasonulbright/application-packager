<#
Vendor: The Git Development Community
App: Git for Windows (x64)
CMName: Git for Windows

.SYNOPSIS
    Packages the latest Git for Windows (x64) for MECM.

.DESCRIPTION
    Queries the GitHub releases API for the latest Git for Windows release, downloads
    the 64-bit EXE installer, stages content to a versioned network folder, and creates
    an MECM Application with a PowerShell script detection method.

    Install:   Git-{version}-64-bit.exe /VERYSILENT /NORESTART /NOCANCEL /SP-
    Uninstall: %ProgramFiles%\Git\unins000.exe /VERYSILENT /NORESTART
    Detection: PowerShell script checking HKLM:\SOFTWARE\GitForWindows CurrentVersion >= packaged version

    GetLatestVersionOnly queries only the GitHub releases API (small JSON) and exits
    without downloading the installer.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM"). The PSDrive is assumed to already exist.

.PARAMETER Comment
    Work order or comment string applied to the MECM application description.

.PARAMETER FileServerPath
    UNC root of the SCCM content share (e.g., "\\fileserver\sccm$").

.PARAMETER GetLatestVersionOnly
    Queries the GitHub API for the latest Git for Windows version, outputs the version
    string, and exits. No download or MECM changes are made.

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
$GitHubApiUrl         = "https://api.github.com/repos/git-for-windows/git/releases/latest"
$BaseDownloadRoot     = Join-Path $env:USERPROFILE "Downloads\_AutoPackager"
$GitNetworkRoot       = Join-Path $FileServerPath "Applications\Git for Windows"
$Publisher            = "The Git Development Community"
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

function Get-LatestGitRelease {
    param([switch]$Quiet)
    if (-not $Quiet) { Write-Host "Querying GitHub releases API: $GitHubApiUrl" }
    $json = (curl.exe -L --fail --silent --show-error -A "PowerShell" $GitHubApiUrl) -join ''
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch Git release info: $GitHubApiUrl" }
    $release = ConvertFrom-Json $json

    # tag_name format: "v2.53.0.windows.1" — strip prefix/suffix for clean semver
    $tagName = $release.tag_name
    $version = $tagName -replace '^v', '' -replace '\.windows\.\d+$', ''  # e.g., "2.53.0"

    if ([string]::IsNullOrWhiteSpace($version)) { throw "Could not parse version from tag: $tagName" }

    # Find 64-bit installer EXE asset
    $asset = $release.assets |
        Where-Object { $_.name -match '^Git-[\d.]+-64-bit\.exe$' } |
        Select-Object -First 1

    if (-not $asset) { throw "Could not locate 64-bit EXE installer asset in release '$tagName'." }

    if (-not $Quiet) {
        Write-Host "Latest Git for Windows release : $tagName"
        Write-Host "Clean version                  : $version"
        Write-Host "Installer asset                : $($asset.name)"
        Write-Host "Download URL                   : $($asset.browser_download_url)"
    }

    return @{
        Version     = $version
        TagName     = $tagName
        FileName    = $asset.name
        DownloadUrl = $asset.browser_download_url
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
        $release = Get-LatestGitRelease -Quiet
        Write-Output $release.Version
        exit 0
    }
    catch {
        Write-Error "Failed to retrieve Git for Windows version: $($_.Exception.Message)"
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
    Write-Host "Git for Windows Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser        : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    Write-Host ("Machine          : {0}"     -f $env:COMPUTERNAME)
    Write-Host "SiteCode         : $SiteCode"
    Write-Host "BaseDownloadRoot : $BaseDownloadRoot"
    Write-Host "GitNetworkRoot   : $GitNetworkRoot"
    Write-Host "GitHubApiUrl     : $GitHubApiUrl"
    Write-Host ""

    Ensure-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $GitNetworkRoot)) {
        throw "Network root path not accessible: $GitNetworkRoot"
    }

    Set-Location C: -ErrorAction Stop

    # 1. Get latest release info
    $release     = Get-LatestGitRelease
    $version     = $release.Version
    $fileName    = $release.FileName
    $downloadUrl = $release.DownloadUrl

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

    # 3. Create versioned content folder
    $contentPath = Join-Path $GitNetworkRoot $version
    Ensure-Folder -Path $contentPath

    # 4. Copy installer to network
    $netExe = Join-Path $contentPath $fileName
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
start /wait "" "%~dp0$fileName" /VERYSILENT /NORESTART /NOCANCEL /SP-
exit /b 0
"@
    Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop

    # 6. Write uninstall.bat
    #    Git for Windows creates %ProgramFiles%\Git\unins000.exe during installation.
    $uninstallBatPath = Join-Path $contentPath "uninstall.bat"
    $uninstallBat = @"
@echo off
setlocal
start /wait "" "%ProgramFiles%\Git\unins000.exe" /VERYSILENT /NORESTART
exit /b 0
"@
    Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop

    Write-Host "install.bat and uninstall.bat created."

    # 7. Build PowerShell detection script
    #    Git for Windows writes HKLM:\SOFTWARE\GitForWindows with CurrentVersion value.
    #    The version may include ".windows.N" suffix; strip it before comparison.
    $DetectionScript = @"
`$reg = Get-ItemProperty 'HKLM:\SOFTWARE\GitForWindows' -ErrorAction SilentlyContinue
if (`$reg -and `$reg.CurrentVersion) {
    `$v = (`$reg.CurrentVersion -replace '\.windows\.\d+$', '').Trim()
    try {
        if ([version]`$v -ge [version]"$version") {
            Write-Output "Installed: `$(`$reg.CurrentVersion)"
        }
    }
    catch { }
}
"@

    # Write detection.ps1 to content folder for reference
    Set-Content -LiteralPath (Join-Path $contentPath "detection.ps1") -Value $DetectionScript -Encoding UTF8 -ErrorAction Stop
    Write-Host "Wrote detection.ps1 (reference copy)."

    # 8. Connect to Configuration Manager
    $appName = "Git for Windows $version"
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

    # 11. Add Script Deployment Type with PowerShell detection
    Write-Host "Adding deployment type '$appName' with PowerShell detection..."
    Add-CMScriptDeploymentType `
        -ApplicationName $appName `
        -DeploymentTypeName $appName `
        -InstallCommand "install.bat" `
        -UninstallCommand "uninstall.bat" `
        -ContentLocation $contentPath `
        -ScriptLanguage PowerShell `
        -ScriptText $DetectionScript `
        -InstallationBehaviorType InstallForSystem `
        -LogonRequirementType WhetherOrNotUserLoggedOn `
        -UserInteractionMode Hidden `
        -MaximumRuntimeMins $MaximumRuntimeMins `
        -EstimatedRuntimeMins $EstimatedRuntimeMins `
        -ContentFallback `
        -SlowNetworkDeploymentMode Download `
        -RebootBehavior NoAction `
        -ErrorAction Stop | Out-Null

    # 12. Revision history cleanup
    Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

    Write-Host "Git for Windows $version packaged successfully." -ForegroundColor Green
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}
finally {
    Set-Location $originalLocation -ErrorAction SilentlyContinue
    Write-Host "Restored initial location to: ${originalLocation}"
}
