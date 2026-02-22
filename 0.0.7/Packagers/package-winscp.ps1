<#
Vendor: Martin Prikryl
App: WinSCP
CMName: WinSCP

.SYNOPSIS
    Automates downloading the latest WinSCP installer (x64) and creating an MECM application.

.DESCRIPTION
    Downloads WinSCP from winscp.net, stores content on a network share, creates install/uninstall batch files,
    temporarily installs locally to extract uninstall registry metadata, and creates an MECM application with
    registry value detection (DisplayVersion). Configures deployment type content settings:
      - Allow clients to use fallback source location for content
      - Download content from distribution point and run locally

.NOTES
    - Run with admin privileges for registry access and local install/uninstall.
    - Requires the Configuration Manager console and PowerShell module.
#>

param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)


# -----------------------
# Configuration
# -----------------------

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager"
$NetworkRootPath  = Join-Path $FileServerPath "Applications\WinSCP"

$PublisherOverride = $null
$ForcedVersion     = $null

$EnableTrace = $false

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# -----------------------
# Logging helpers
# -----------------------
function Write-Section {
    param([string]$Text)
    Write-Host ""
    Write-Host "============================================================"
    Write-Host $Text
    Write-Host "============================================================"
}

function Write-KV {
    param([string]$K, [string]$V)
    Write-Host ("{0,-28}: {1}" -f $K, $V)
}

function Assert-Ok {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

# -----------------------
# Admin check
# -----------------------
function Test-IsAdmin {
    try {
        $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Warning "Failed to check admin privileges: $($_.Exception.Message)"
        return $false
    }
}

# -----------------------
# CM site connection (mirrors your Tableau script pattern)
# -----------------------
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


# -----------------------
# WinSCP version + download (CDN secure link)
# -----------------------
function Get-LatestWinSCPVersion {
    param([switch]$Quiet)
    $url = "https://winscp.net/eng/downloads.php"
    if (-not $Quiet) { Write-Host "Fetching WinSCP version from: $url" }

    $html = (curl.exe -L --max-redirs 10 --fail --silent --show-error $url) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch WinSCP downloads page: $url" }

    if ($html -match 'Download\s+WinSCP\s+([0-9]+\.[0-9]+\.[0-9]+)') {
        return $matches[1]
    }
    if ($html -match 'WinSCP-([0-9]+\.[0-9]+\.[0-9]+)-Setup\.exe') {
        return $matches[1]
    }

    throw "Could not parse latest WinSCP version from downloads page."
}

function Get-WinSCPDirectDownloadUrl {
    param([Parameter(Mandatory)][string]$Version)

    $pageUrl = "https://winscp.net/download/WinSCP-$Version-Setup.exe/download"
    Write-Host "WinSCP per-file download page: $pageUrl"

    $html = (curl.exe -L --max-redirs 10 --fail --silent --show-error $pageUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch WinSCP download page: $pageUrl" }

    $cdn = $null
    if ($html -match '(https?://cdn\.winscp\.net/files/WinSCP-[0-9]+\.[0-9]+\.[0-9]+-Setup\.exe\?secure=[^"]+)') {
        $cdn = $matches[1]
    }

    if (-not $cdn) {
        throw "Could not locate CDN direct download link on page: $pageUrl"
    }

    Write-Host "Resolved CDN direct download URL: $cdn"
    return $cdn
}

function Test-DownloadedInstaller {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path $Path)) { return $false }

    $len = (Get-Item $Path).Length
    if ($len -lt 1MB) { return $false }

    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $buf = New-Object byte[] 64
        [void]$fs.Read($buf, 0, $buf.Length)
        $head = [System.Text.Encoding]::ASCII.GetString($buf)
        if ($head -match '<!DOCTYPE|<html|<HTML') { return $false }
    }
    finally { $fs.Dispose() }

    return $true
}

# -----------------------
# Local install + uninstall key discovery
# -----------------------
function Install-WinSCPForDiscovery {
    param([Parameter(Mandatory)][string]$InstallerPath)

    $args = "/VERYSILENT /NORESTART /ALLUSERS"

    Write-Host "Installing WinSCP locally for discovery:"
    Write-KV "Installer" $InstallerPath
    Write-KV "Args" $args

    $p = Start-Process -FilePath $InstallerPath -ArgumentList $args -Wait -PassThru -ErrorAction Stop
    Write-KV "ExitCode" $p.ExitCode

    if ($p.ExitCode -ne 0) {
        throw "Installer returned non-zero exit code: $($p.ExitCode)"
    }
}

function Find-WinSCPUninstallEntry {
    param([Parameter(Mandatory)][string]$ExpectedVersion)

    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $candidates = @()

    foreach ($p in $paths) {
        Write-Host "Scanning uninstall keys: $p"
        $items = Get-ItemProperty -Path $p -ErrorAction SilentlyContinue
        foreach ($it in $items) {
            $dn = $it.DisplayName
            $dv = $it.DisplayVersion
            if ([string]::IsNullOrWhiteSpace($dn)) { continue }
            if ($dn -match '^WinSCP') {
                $obj = [pscustomobject]@{
                    KeyName              = ($it.PSPath -replace '^Microsoft\.PowerShell\.Core\\Registry::HKEY_LOCAL_MACHINE\\', 'HKLM:\')
                    DisplayName          = $dn
                    DisplayVersion       = $dv
                    Publisher            = $it.Publisher
                    InstallLocation      = $it.InstallLocation
                    UninstallString      = $it.UninstallString
                    QuietUninstallString = $it.QuietUninstallString
                }
                $candidates += $obj
            }
        }
    }

    Write-Host ""
    Write-Host "WinSCP uninstall candidates found: $($candidates.Count)"
    $candidates | ForEach-Object {
        Write-Host "  - $($_.DisplayName) | Version=$($_.DisplayVersion) | Publisher=$($_.Publisher)"
        Write-Host "    Key=$($_.KeyName)"
    }

    $match = $candidates | Where-Object { $_.DisplayVersion -eq $ExpectedVersion } | Select-Object -First 1
    if ($match) { return $match }

    return ($candidates | Select-Object -First 1)
}

function Uninstall-WinSCPFromDiscovery {
    param(
        [Parameter()][string]$QuietUninstallString,
        [Parameter()][string]$UninstallString
    )

    $cmd = $QuietUninstallString
    if ([string]::IsNullOrWhiteSpace($cmd)) { $cmd = $UninstallString }

    if ([string]::IsNullOrWhiteSpace($cmd)) {
        Write-Warning "No uninstall command discovered; leaving WinSCP installed on packaging machine."
        return
    }

    Write-Host "Attempting to uninstall WinSCP used for discovery:"
    Write-KV "Command" $cmd

    $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd" -Wait -PassThru -ErrorAction Stop
    Write-KV "Uninstall ExitCode" $p.ExitCode
}

# -----------------------
# Batch wrappers
# -----------------------
function Create-BatchFiles {
    param(
        [Parameter(Mandatory)][string]$ContentFolder,
        [Parameter(Mandatory)][string]$InstallerFileName
    )

    $installBat = @"
@echo off
setlocal
echo Installing WinSCP...
start /wait "" "%~dp0$InstallerFileName" /VERYSILENT /NORESTART /ALLUSERS
exit /b %ERRORLEVEL%
"@

    $uninstallBat = @"
@echo off
setlocal

set "U1=%ProgramFiles%\WinSCP\unins000.exe"
set "U2=%ProgramFiles(x86)%\WinSCP\unins000.exe"

if exist "%U1%" (
  start /wait "" "%U1%" /VERYSILENT /NORESTART
  exit /b %ERRORLEVEL%
)

if exist "%U2%" (
  start /wait "" "%U2%" /VERYSILENT /NORESTART
  exit /b %ERRORLEVEL%
)

echo WinSCP uninstall executable not found.
exit /b 0
"@

    Set-Content -Path (Join-Path $ContentFolder "install.bat")   -Value $installBat   -Encoding ASCII
    Set-Content -Path (Join-Path $ContentFolder "uninstall.bat") -Value $uninstallBat -Encoding ASCII

    Write-Host "Created batch wrappers:"
    Write-KV "install.bat"   (Join-Path $ContentFolder "install.bat")
    Write-KV "uninstall.bat" (Join-Path $ContentFolder "uninstall.bat")
}

# -----------------------
# MECM creation
# -----------------------
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

function New-WinSCPMecmApp {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$Publisher,
        [Parameter(Mandatory)][string]$ContentLocationUNC,
        [Parameter(Mandatory)][string]$UninstallRegKeyRelative,
        [Parameter(Mandatory)][string]$DeploymentTypeName
    )

    Write-Section "MECM: Creating Application + Deployment Type"

    $originalLocation = Get-Location
    Write-Host "Current location before MECM work: ${originalLocation}"

    try {

        Connect-CMSite -SiteCode $SiteCode | Out-Null

        $existing = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Warning "Application already exists: $AppName"
            return
        }

        Write-Host "Creating CM Application..."
        $cmApp = New-CMApplication -Name $AppName -Publisher $Publisher -SoftwareVersion $Version -Description $Comment -ErrorAction Stop

        $detectionClause = New-CMDetectionClauseRegistryKeyValue `
            -Hive LocalMachine `
            -KeyName $UninstallRegKeyRelative `
            -ValueName "DisplayVersion" `
            -PropertyType String `
            -Value `
            -ExpectedValue $Version `
            -ExpressionOperator IsEquals `
            -Is64Bit

        Write-Host "Detection clause details:"
        Write-KV "Registry Key"   $UninstallRegKeyRelative
        Write-KV "ValueName"      "DisplayVersion"
        Write-KV "ExpectedValue"  $Version

        $params = @{
            ApplicationName          = $AppName
            DeploymentTypeName       = $DeploymentTypeName
            InstallCommand           = "install.bat"
            ContentLocation          = $ContentLocationUNC
            UninstallCommand         = "uninstall.bat"
            InstallationBehaviorType = "InstallForSystem"
            LogonRequirementType     = "WhetherOrNotUserLoggedOn"
            MaximumRuntimeMins       = 20
            EstimatedRuntimeMins     = 10
            AddDetectionClause       = @($detectionClause)
            ContentFallback          = $true
            ErrorAction              = "Stop"
        }

        Write-Host "Adding Script Deployment Type..."
        Add-CMScriptDeploymentType @params
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Host "Configuring deployment type content options..."
        Set-CMScriptDeploymentType -ApplicationName $AppName -DeploymentTypeName $DeploymentTypeName -ContentFallback $true -SlowNetworkDeploymentMode Download -ErrorAction Stop

        Write-Host "Created MECM application: ${AppName}"
    }
    catch {
        Write-Error "Failed to create MECM application: $($_.Exception.Message)"
        throw
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
        Write-Host "Restored location to: ${originalLocation}"
    }
}

# -----------------------
# Main
# -----------------------
try {
    Write-Section "WinSCP Auto-Packager starting"

    $originalLocation = Get-Location
    Write-Host "Current location at script start: ${originalLocation}"

    if ($GetLatestVersionOnly) {
        $v = $ForcedVersion
        if ([string]::IsNullOrWhiteSpace($v)) {
            $v = Get-LatestWinSCPVersion -Quiet
        }
        Write-Output $v
        return
    }


    if (-not (Test-IsAdmin)) {
        Write-Error "This script must be run with admin privileges. Please run PowerShell as Administrator."
        exit 1
    }

    Set-Location $PSScriptRoot -ErrorAction Stop
    Write-Host "Set initial location to script directory: ${PSScriptRoot}"

    Write-KV "RunAsUser"        $env:USERNAME
    Write-KV "Machine"          $env:COMPUTERNAME
    Write-KV "SiteCode"         $SiteCode
    Write-KV "NetworkRootPath"  $NetworkRootPath
    Write-KV "BaseDownloadRoot" $BaseDownloadRoot

    $version = $ForcedVersion
    if ([string]::IsNullOrWhiteSpace($version)) {
        $version = Get-LatestWinSCPVersion
    }
    Assert-Ok ($version -match '^\d+\.\d+\.\d+$') "Parsed version '$version' does not look like x.y.z"

    Write-Section "Resolved version"
    Write-KV "WinSCP Version" $version

    $installerFileName  = "WinSCP-$version-Setup.exe"
    $localVersionFolder = Join-Path $BaseDownloadRoot "WinSCP\$version"
    $localInstallerPath = Join-Path $localVersionFolder $installerFileName

    $networkVersionFolder = Join-Path $NetworkRootPath $version
    $networkInstallerPath = Join-Path $networkVersionFolder $installerFileName

    Write-Section "Prepare folders"
    Write-KV "LocalVersionFolder"   $localVersionFolder
    Write-KV "NetworkVersionFolder" $networkVersionFolder

    if (-not (Test-Path $localVersionFolder))   { New-Item -ItemType Directory -Path $localVersionFolder -Force | Out-Null }
    if (-not (Test-Path $networkVersionFolder)) { New-Item -ItemType Directory -Path $networkVersionFolder -Force | Out-Null }

    Write-Section "Download installer"
    if (-not (Test-Path $localInstallerPath)) {
        $cdnUrl = Get-WinSCPDirectDownloadUrl -Version $version

        Write-Host "Downloading WinSCP..."
        Write-KV "From" $cdnUrl
        Write-KV "To"   $localInstallerPath

        curl.exe --max-redirs 10 --fail --silent --show-error -o $localInstallerPath $cdnUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $cdnUrl" }

        if (-not (Test-DownloadedInstaller -Path $localInstallerPath)) {
            try { Remove-Item -Path $localInstallerPath -Force -ErrorAction SilentlyContinue } catch {}
            throw "Downloaded file did not validate as installer (too small/HTML)."
        }

        Write-Host "Download validated."
        Write-KV "Local bytes" ((Get-Item $localInstallerPath).Length.ToString())
    }
    else {
        Write-Host "Local installer already exists: $localInstallerPath"
        Write-KV "Local bytes" ((Get-Item $localInstallerPath).Length.ToString())
    }

    Write-Section "Copy installer to network"
    if ((Get-Location).Provider.Name -ne "FileSystem") {
        Set-Location -LiteralPath $PSScriptRoot -ErrorAction Stop
        Write-Host "Set location to script directory for file operations: ${PSScriptRoot}"
    }

    if (-not (Test-Path $networkInstallerPath)) {
        Copy-Item -Path $localInstallerPath -Destination $networkInstallerPath -Force -ErrorAction Stop
        Write-Host "Copied installer to network: $networkInstallerPath"
    }
    else {
        Write-Host "Network installer already exists: $networkInstallerPath"
        Write-KV "Network bytes" ((Get-Item $networkInstallerPath).Length.ToString())
    }

    Write-Section "Create batch files"
    Create-BatchFiles -ContentFolder $networkVersionFolder -InstallerFileName $installerFileName

    Write-Section "Local discovery install + registry crawl"
    Install-WinSCPForDiscovery -InstallerPath $localInstallerPath

    $uninstallEntry = Find-WinSCPUninstallEntry -ExpectedVersion $version
    Assert-Ok ($null -ne $uninstallEntry) "Could not find any WinSCP uninstall entry after install."

    Write-Host ""
    Write-Host "Selected uninstall entry:"
    Write-KV "DisplayName"          $uninstallEntry.DisplayName
    Write-KV "DisplayVersion"       $uninstallEntry.DisplayVersion
    Write-KV "Publisher"            $uninstallEntry.Publisher
    Write-KV "InstallLocation"      $uninstallEntry.InstallLocation
    Write-KV "UninstallString"      $uninstallEntry.UninstallString
    Write-KV "QuietUninstallString" $uninstallEntry.QuietUninstallString
    Write-KV "RegistryKey"          $uninstallEntry.KeyName

    $regRelative = ($uninstallEntry.KeyName -replace '^HKLM:\\', '')
    Assert-Ok (-not [string]::IsNullOrWhiteSpace($regRelative)) "Failed to compute registry relative key path."

    $publisher = if ($PublisherOverride) { $PublisherOverride } else { $uninstallEntry.Publisher }
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "WinSCP" }

    $appName = $uninstallEntry.DisplayName
    $dtName  = $uninstallEntry.DisplayName

    Write-Section "MECM naming"
    Write-KV "ApplicationName"      $appName
    Write-KV "DeploymentTypeName"   $dtName
    Write-KV "Publisher"            $publisher
    Write-KV "ContentLocation"      $networkVersionFolder
    Write-KV "DetectionRegKey"      $regRelative

    New-WinSCPMecmApp `
        -AppName $appName `
        -Version $version `
        -Publisher $publisher `
        -ContentLocationUNC $networkVersionFolder `
        -UninstallRegKeyRelative $regRelative `
        -DeploymentTypeName $dtName

    Write-Section "Cleanup: uninstall discovery install"
    Uninstall-WinSCPFromDiscovery -QuietUninstallString $uninstallEntry.QuietUninstallString -UninstallString $uninstallEntry.UninstallString

    Write-Section "DONE"
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    throw
}
finally {
    try {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
        Write-Host "Restored location to: ${originalLocation}"
    } catch {}
}