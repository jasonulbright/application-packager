<#
Vendor: Martin Prikryl
App: WinSCP (x64)
CMName: WinSCP

.SYNOPSIS
    Packages WinSCP (x64) for MECM.

.DESCRIPTION
    Downloads the latest WinSCP installer from winscp.net, stages content to a
    versioned network location, temporarily installs locally to extract registry
    metadata (DisplayName, Publisher, uninstall key), creates an MECM Application
    with registry-based detection (DisplayVersion), then uninstalls from the
    packaging machine.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\WinSCP\WinSCP\<Version>

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available WinSCP version string and exits.

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
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Configuration ---
$VendorFolder = "WinSCP"
$AppFolder    = "WinSCP"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager\WinSCP"

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
        if (-not (Get-Module -Name ConfigurationManager -ErrorAction SilentlyContinue)) {
            $cmModulePath = Join-Path $env:SMS_ADMIN_UI_PATH "..\ConfigurationManager.psd1"
            if (Test-Path -LiteralPath $cmModulePath) {
                Import-Module $cmModulePath -ErrorAction Stop
            }
            else {
                Import-Module ConfigurationManager -ErrorAction Stop
            }
        }

        if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
            throw "Configuration Manager PSDrive '$SiteCode' is not available."
        }

        Set-Location "${SiteCode}:" -ErrorAction Stop
        Write-Host "Connected to CM site: $SiteCode"
        return $true
    }
    catch {
        Write-Error "Failed to connect to CM site: $($_.Exception.Message)"
        return $false
    }
}

function Initialize-Folder {
    param([Parameter(Mandatory)][string]$Path)

    $origLocation = Get-Location
    try {
        Set-Location C: -ErrorAction Stop
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }
    }
    finally {
        Set-Location $origLocation -ErrorAction SilentlyContinue
    }
}

function Test-NetworkShareAccess {
    param([Parameter(Mandatory)][string]$Path)

    $origLocation = Get-Location
    try {
        Set-Location C: -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) {
            Write-Error "Network path does not exist or is inaccessible: $Path"
            return $false
        }

        try {
            $tmp = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
            Set-Content -LiteralPath $tmp -Value "test" -Encoding ASCII -ErrorAction Stop
            Remove-Item -LiteralPath $tmp -ErrorAction Stop
            return $true
        }
        catch {
            Write-Error "Network share is not writable: $Path ($($_.Exception.Message))"
            return $false
        }
    }
    finally {
        Set-Location $origLocation -ErrorAction SilentlyContinue
    }
}

function Get-LatestWinSCPVersion {
    param([switch]$Quiet)

    $url = "https://winscp.net/eng/downloads.php"
    if (-not $Quiet) {
        Write-Host "WinSCP downloads page        : $url"
    }

    try {
        $html = (curl.exe -L --max-redirs 10 --fail --silent --show-error $url) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch WinSCP downloads page: $url" }

        $version = $null
        if ($html -match 'Download\s+WinSCP\s+([0-9]+\.[0-9]+\.[0-9]+)') {
            $version = $matches[1]
        }
        elseif ($html -match 'WinSCP-([0-9]+\.[0-9]+\.[0-9]+)-Setup\.exe') {
            $version = $matches[1]
        }

        if (-not $version) {
            throw "Could not parse latest WinSCP version from downloads page."
        }

        if (-not $Quiet) {
            Write-Host "Latest WinSCP version        : $version"
        }
        return $version
    }
    catch {
        Write-Error "Failed to get WinSCP version: $($_.Exception.Message)"
        return $null
    }
}

function Resolve-WinSCPDownloadUrl {
    param([Parameter(Mandatory)][string]$Version)

    $pageUrl = "https://winscp.net/download/WinSCP-$Version-Setup.exe/download"
    Write-Host "WinSCP per-file download page: $pageUrl"

    try {
        $html = (curl.exe -L --max-redirs 10 --fail --silent --show-error $pageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch WinSCP download page: $pageUrl" }

        if ($html -match '(https?://cdn\.winscp\.net/files/WinSCP-[0-9]+\.[0-9]+\.[0-9]+-Setup\.exe\?secure=[^"]+)') {
            $cdnUrl = $matches[1]
            Write-Host "Resolved CDN URL             : $cdnUrl"
            return $cdnUrl
        }

        throw "Could not locate CDN direct download link on page: $pageUrl"
    }
    catch {
        Write-Error "Failed to resolve WinSCP download URL: $($_.Exception.Message)"
        return $null
    }
}

function Test-DownloadedInstaller {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $false }

    $len = (Get-Item -LiteralPath $Path).Length
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

function Install-WinSCPForDiscovery {
    param([Parameter(Mandatory)][string]$InstallerPath)

    Write-Host "Installing WinSCP locally for registry discovery..."
    Write-Host "Installer                    : $InstallerPath"

    $p = Start-Process -FilePath $InstallerPath -ArgumentList "/VERYSILENT /NORESTART /ALLUSERS" -Wait -PassThru -ErrorAction Stop
    Write-Host "Install exit code            : $($p.ExitCode)"

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
        $items = Get-ItemProperty -Path $p -ErrorAction SilentlyContinue
        foreach ($it in $items) {
            $dn = $it.DisplayName
            if ([string]::IsNullOrWhiteSpace($dn)) { continue }
            if ($dn -match '^WinSCP') {
                $candidates += [pscustomobject]@{
                    KeyName              = ($it.PSPath -replace '^Microsoft\.PowerShell\.Core\\Registry::HKEY_LOCAL_MACHINE\\', 'HKLM:\')
                    DisplayName          = $dn
                    DisplayVersion       = $it.DisplayVersion
                    Publisher            = $it.Publisher
                    UninstallString      = $it.UninstallString
                    QuietUninstallString = $it.QuietUninstallString
                }
            }
        }
    }

    Write-Host "WinSCP uninstall candidates  : $($candidates.Count)"

    $match = $candidates | Where-Object { $_.DisplayVersion -eq $ExpectedVersion } | Select-Object -First 1
    if ($match) { return $match }

    return ($candidates | Select-Object -First 1)
}

function Uninstall-WinSCPFromDiscovery {
    param(
        [string]$QuietUninstallString,
        [string]$UninstallString
    )

    $cmd = $QuietUninstallString
    if ([string]::IsNullOrWhiteSpace($cmd)) { $cmd = $UninstallString }

    if ([string]::IsNullOrWhiteSpace($cmd)) {
        Write-Warning "No uninstall command discovered; leaving WinSCP installed on packaging machine."
        return
    }

    Write-Host "Uninstalling discovery install..."
    Write-Host "Uninstall command            : $cmd"

    $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd" -Wait -PassThru -ErrorAction Stop
    Write-Host "Uninstall exit code          : $($p.ExitCode)"
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
        if ($h.PSObject.Properties.Name -contains 'Revision') { $revs += [UInt32]$h.Revision; continue }
        if ($h.PSObject.Properties.Name -contains 'CIVersion') { $revs += [UInt32]$h.CIVersion; continue }
    }

    $revs = $revs | Sort-Object -Unique -Descending
    if ($revs.Count -le $KeepLatest) { return }

    foreach ($rev in ($revs | Select-Object -Skip $KeepLatest)) {
        Remove-CMApplicationRevisionHistory -Id $CI_ID -Revision $rev -Force -ErrorAction Stop
    }
}

function New-MECMWinSCPApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$ContentPath,
        [Parameter(Mandatory)][string]$InstallerFileName,
        [Parameter(Mandatory)][string]$Publisher,
        [Parameter(Mandatory)][string]$UninstallRegKeyRelative
    )

    $orig = Get-Location

    try {
        if (-not (Test-IsAdmin)) { throw "Run PowerShell as Administrator." }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $existing = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Warning "Application already exists: $AppName"
            return
        }

        Write-Host "Creating CM Application      : $AppName"
        $cmApp = New-CMApplication `
            -Name $AppName `
            -Publisher $Publisher `
            -SoftwareVersion $SoftwareVersion `
            -Description $Comment `
            -AutoInstall $true `
            -ErrorAction Stop

        Write-Host "Application CI_ID            : $($cmApp.CI_ID)"

        Set-Location C: -ErrorAction Stop

        $installBatPath   = Join-Path $ContentPath "install.bat"
        $uninstallBatPath = Join-Path $ContentPath "uninstall.bat"

        if (-not (Test-Path -LiteralPath $installBatPath)) {
            $installBat = @"
@echo off
setlocal
start /wait "" "%~dp0$InstallerFileName" /VERYSILENT /NORESTART /ALLUSERS
exit /b 0
"@
            Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
setlocal

set "U1=%ProgramFiles%\WinSCP\unins000.exe"
set "U2=%ProgramFiles(x86)%\WinSCP\unins000.exe"

if exist "%U1%" (
  start /wait "" "%U1%" /VERYSILENT /NORESTART
  exit /b 0
)

if exist "%U2%" (
  start /wait "" "%U2%" /VERYSILENT /NORESTART
  exit /b 0
)

echo WinSCP uninstall executable not found.
exit /b 0
"@
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $dtName = $AppName

        $clause = New-CMDetectionClauseRegistryKeyValue `
            -Hive LocalMachine `
            -KeyName $UninstallRegKeyRelative `
            -ValueName "DisplayVersion" `
            -PropertyType String `
            -Value `
            -ExpectedValue $SoftwareVersion `
            -ExpressionOperator IsEquals `
            -Is64Bit

        Write-Host "Adding Script Deployment Type: $dtName"
        Add-CMScriptDeploymentType `
            -ApplicationName $AppName `
            -DeploymentTypeName $dtName `
            -ContentLocation $ContentPath `
            -InstallCommand "install.bat" `
            -UninstallCommand "uninstall.bat" `
            -InstallationBehaviorType InstallForSystem `
            -LogonRequirementType WhetherOrNotUserLoggedOn `
            -EstimatedRuntimeMins $EstimatedRuntimeMins `
            -MaximumRuntimeMins $MaximumRuntimeMins `
            -AddDetectionClause @($clause) `
            -ContentFallback `
            -SlowNetworkDeploymentMode Download `
            -ErrorAction Stop | Out-Null

        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Host "Created MECM application     : $AppName"
    }
    finally {
        Set-Location $orig -ErrorAction SilentlyContinue
    }
}

function Get-WinSCPNetworkAppRoot {
    param([Parameter(Mandatory)][string]$FileServerPath)

    $appsRoot   = Join-Path $FileServerPath "Applications"
    $vendorPath = Join-Path $appsRoot $VendorFolder
    $appPath    = Join-Path $vendorPath $AppFolder

    Initialize-Folder -Path $appsRoot
    Initialize-Folder -Path $vendorPath
    Initialize-Folder -Path $appPath

    return $appPath
}

# --- Latest-only mode ---
if ($GetLatestVersionOnly) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $v = Get-LatestWinSCPVersion -Quiet
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

    Write-Host ""
    Write-Host ("=" * 60)
    Write-Host "WinSCP (x64) Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Host ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Host "Start location               : $startLocation"
    Write-Host "SiteCode                     : $SiteCode"
    Write-Host "FileServerPath               : $FileServerPath"
    Write-Host "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-WinSCPNetworkAppRoot -FileServerPath $FileServerPath

    $version = Get-LatestWinSCPVersion
    if (-not $version) {
        throw "Could not resolve WinSCP version."
    }

    if ($version -notmatch '^\d+\.\d+\.\d+$') {
        throw "Parsed version '$version' does not look like x.y.z"
    }

    $installerFileName = "WinSCP-$version-Setup.exe"
    $localExe          = Join-Path $BaseDownloadRoot $installerFileName
    $contentPath       = Join-Path $networkAppRoot $version

    Initialize-Folder -Path $contentPath

    $netExe = Join-Path $contentPath $installerFileName

    Write-Host "Version                      : $version"
    Write-Host "Installer filename           : $installerFileName"
    Write-Host "Local installer              : $localExe"
    Write-Host "ContentPath                  : $contentPath"
    Write-Host "Network installer            : $netExe"
    Write-Host ""

    # Download (two-step: resolve CDN URL, then download)
    if (-not (Test-Path -LiteralPath $localExe)) {
        $cdnUrl = Resolve-WinSCPDownloadUrl -Version $version
        if (-not $cdnUrl) {
            throw "Could not resolve WinSCP CDN download URL."
        }

        Write-Host "Downloading installer..."
        curl.exe --max-redirs 10 --fail --silent --show-error -o $localExe $cdnUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $cdnUrl" }

        if (-not (Test-DownloadedInstaller -Path $localExe)) {
            try { Remove-Item -LiteralPath $localExe -Force -ErrorAction SilentlyContinue } catch {}
            throw "Downloaded file did not validate as installer (too small or HTML content)."
        }

        Write-Host "Download validated             ($((Get-Item -LiteralPath $localExe).Length) bytes)"
    }
    else {
        Write-Host "Local installer exists. Skipping download."
    }

    if (-not (Test-Path -LiteralPath $netExe)) {
        Write-Host "Copying installer to network..."
        Copy-Item -LiteralPath $localExe -Destination $netExe -Force -ErrorAction Stop
    }
    else {
        Write-Host "Network installer exists. Skipping copy."
    }

    # Local install for registry discovery
    Write-Host ""
    Install-WinSCPForDiscovery -InstallerPath $localExe

    $uninstallEntry = Find-WinSCPUninstallEntry -ExpectedVersion $version
    if ($null -eq $uninstallEntry) {
        throw "Could not find any WinSCP uninstall entry after install."
    }

    Write-Host "Registry DisplayName         : $($uninstallEntry.DisplayName)"
    Write-Host "Registry DisplayVersion      : $($uninstallEntry.DisplayVersion)"
    Write-Host "Registry Publisher           : $($uninstallEntry.Publisher)"
    Write-Host "Registry Key                 : $($uninstallEntry.KeyName)"

    $regRelative = ($uninstallEntry.KeyName -replace '^HKLM:\\', '')
    if ([string]::IsNullOrWhiteSpace($regRelative)) {
        throw "Failed to compute registry relative key path."
    }

    $publisher = $uninstallEntry.Publisher
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Martin Prikryl" }

    $appName = $uninstallEntry.DisplayName

    Write-Host ""
    Write-Host "CM Application Name          : $appName"
    Write-Host "CM SoftwareVersion           : $version"
    Write-Host "Detection RegKey             : $regRelative"
    Write-Host ""

    New-MECMWinSCPApplication `
        -AppName $appName `
        -SoftwareVersion $version `
        -ContentPath $contentPath `
        -InstallerFileName $installerFileName `
        -Publisher $publisher `
        -UninstallRegKeyRelative $regRelative

    # Cleanup: uninstall the discovery install
    Write-Host ""
    Uninstall-WinSCPFromDiscovery `
        -QuietUninstallString $uninstallEntry.QuietUninstallString `
        -UninstallString $uninstallEntry.UninstallString

    Write-Host ""
    Write-Host "Script execution complete."
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
