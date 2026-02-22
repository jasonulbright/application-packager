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
    [string]$LogPath,
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

. "$PSScriptRoot\AppPackagerCommon.ps1"
Initialize-Logging -LogPath $LogPath

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
        Write-Log "Admin check failed: $($_.Exception.Message)" -Level WARN
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
        Write-Log "Connected to CM site: $SiteCode"
        return $true
    }
    catch {
        Write-Log "Failed to connect to CM site: $($_.Exception.Message)" -Level ERROR
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
            Write-Log "Network path does not exist or is inaccessible: $Path" -Level ERROR
            return $false
        }

        try {
            $tmp = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
            Set-Content -LiteralPath $tmp -Value "test" -Encoding ASCII -ErrorAction Stop
            Remove-Item -LiteralPath $tmp -ErrorAction Stop
            return $true
        }
        catch {
            Write-Log "Network share is not writable: $Path ($($_.Exception.Message))" -Level ERROR
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
    Write-Log "WinSCP downloads page        : $url" -Quiet:$Quiet

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

        Write-Log "Latest WinSCP version        : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get WinSCP version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Resolve-WinSCPDownloadUrl {
    param([Parameter(Mandatory)][string]$Version)

    $pageUrl = "https://winscp.net/download/WinSCP-$Version-Setup.exe/download"
    Write-Log "WinSCP per-file download page: $pageUrl"

    try {
        $html = (curl.exe -L --max-redirs 10 --fail --silent --show-error $pageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch WinSCP download page: $pageUrl" }

        if ($html -match '(https?://cdn\.winscp\.net/files/WinSCP-[0-9]+\.[0-9]+\.[0-9]+-Setup\.exe\?secure=[^"]+)') {
            $cdnUrl = $matches[1]
            Write-Log "Resolved CDN URL             : $cdnUrl"
            return $cdnUrl
        }

        throw "Could not locate CDN direct download link on page: $pageUrl"
    }
    catch {
        Write-Log "Failed to resolve WinSCP download URL: $($_.Exception.Message)" -Level ERROR
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

    Write-Log "Installing WinSCP locally for registry discovery..."
    Write-Log "Installer                    : $InstallerPath"

    $p = Start-Process -FilePath $InstallerPath -ArgumentList "/VERYSILENT /NORESTART /ALLUSERS" -Wait -PassThru -ErrorAction Stop
    Write-Log "Install exit code            : $($p.ExitCode)"

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

    Write-Log "WinSCP uninstall candidates  : $($candidates.Count)"

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
        Write-Log "No uninstall command discovered; leaving WinSCP installed on packaging machine." -Level WARN
        return
    }

    Write-Log "Uninstalling discovery install..."
    Write-Log "Uninstall command            : $cmd"

    $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd" -Wait -PassThru -ErrorAction Stop
    Write-Log "Uninstall exit code          : $($p.ExitCode)"
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
            Write-Log "Application already exists: $AppName" -Level WARN
            return
        }

        Write-Log "Creating CM Application      : $AppName"
        $cmApp = New-CMApplication `
            -Name $AppName `
            -Publisher $Publisher `
            -SoftwareVersion $SoftwareVersion `
            -Description $Comment `
            -AutoInstall $true `
            -ErrorAction Stop

        Write-Log "Application CI_ID            : $($cmApp.CI_ID)"

        Set-Location C: -ErrorAction Stop

        $installBatPath   = Join-Path $ContentPath "install.bat"
        $installPs1Path   = Join-Path $ContentPath "install.ps1"
        $uninstallBatPath = Join-Path $ContentPath "uninstall.bat"
        $uninstallPs1Path = Join-Path $ContentPath "uninstall.ps1"

        if (-not (Test-Path -LiteralPath $installBatPath)) {
            $installBat = @"
@echo off
PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0install.ps1"
exit /b %ERRORLEVEL%
"@
            Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $installPs1Path)) {
            $installPs1 = @"
`$proc = Start-Process "`$PSScriptRoot\$InstallerFileName" -ArgumentList "/VERYSILENT /NORESTART /ALLUSERS" -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@
            Set-Content -LiteralPath $installPs1Path -Value $installPs1 -Encoding UTF8 -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
exit /b %ERRORLEVEL%
"@
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallPs1Path)) {
            $uninstallPs1 = @"
`$uninstaller = "`$env:ProgramFiles\WinSCP\unins000.exe"
if (-not (Test-Path -LiteralPath `$uninstaller)) {
    `$uninstaller = "`${env:ProgramFiles(x86)}\WinSCP\unins000.exe"
}
if (Test-Path -LiteralPath `$uninstaller) {
    `$proc = Start-Process `$uninstaller -ArgumentList "/VERYSILENT /NORESTART" -Wait -PassThru -NoNewWindow
    exit `$proc.ExitCode
}
Write-Warning "WinSCP uninstall executable not found."
exit 0
"@
            Set-Content -LiteralPath $uninstallPs1Path -Value $uninstallPs1 -Encoding UTF8 -ErrorAction Stop
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

        Write-Log "Adding Script Deployment Type: $dtName"
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

        Write-Log "Created MECM application     : $AppName"
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

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "WinSCP (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
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

    Write-Log "Version                      : $version"
    Write-Log "Installer filename           : $installerFileName"
    Write-Log "Local installer              : $localExe"
    Write-Log "ContentPath                  : $contentPath"
    Write-Log "Network installer            : $netExe"
    Write-Log ""

    # Download (two-step: resolve CDN URL, then download)
    if (-not (Test-Path -LiteralPath $localExe)) {
        $cdnUrl = Resolve-WinSCPDownloadUrl -Version $version
        if (-not $cdnUrl) {
            throw "Could not resolve WinSCP CDN download URL."
        }

        Write-Log "Downloading installer..."
        Invoke-DownloadWithRetry -Url $cdnUrl -OutFile $localExe -ExtraCurlArgs @('--max-redirs', '10')

        if (-not (Test-DownloadedInstaller -Path $localExe)) {
            try { Remove-Item -LiteralPath $localExe -Force -ErrorAction SilentlyContinue } catch {}
            throw "Downloaded file did not validate as installer (too small or HTML content)."
        }

        Write-Log "Download validated             ($((Get-Item -LiteralPath $localExe).Length) bytes)"
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    if (-not (Test-Path -LiteralPath $netExe)) {
        Write-Log "Copying installer to network..."
        Copy-Item -LiteralPath $localExe -Destination $netExe -Force -ErrorAction Stop
    }
    else {
        Write-Log "Network installer exists. Skipping copy."
    }

    # Local install for registry discovery
    Write-Log ""
    Install-WinSCPForDiscovery -InstallerPath $localExe

    $uninstallEntry = Find-WinSCPUninstallEntry -ExpectedVersion $version
    if ($null -eq $uninstallEntry) {
        throw "Could not find any WinSCP uninstall entry after install."
    }

    Write-Log "Registry DisplayName         : $($uninstallEntry.DisplayName)"
    Write-Log "Registry DisplayVersion      : $($uninstallEntry.DisplayVersion)"
    Write-Log "Registry Publisher           : $($uninstallEntry.Publisher)"
    Write-Log "Registry Key                 : $($uninstallEntry.KeyName)"

    $regRelative = ($uninstallEntry.KeyName -replace '^HKLM:\\', '')
    if ([string]::IsNullOrWhiteSpace($regRelative)) {
        throw "Failed to compute registry relative key path."
    }

    $publisher = $uninstallEntry.Publisher
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Martin Prikryl" }

    $appName = $uninstallEntry.DisplayName

    Write-Log ""
    Write-Log "CM Application Name          : $appName"
    Write-Log "CM SoftwareVersion           : $version"
    Write-Log "Detection RegKey             : $regRelative"
    Write-Log ""

    New-MECMWinSCPApplication `
        -AppName $appName `
        -SoftwareVersion $version `
        -ContentPath $contentPath `
        -InstallerFileName $installerFileName `
        -Publisher $publisher `
        -UninstallRegKeyRelative $regRelative

    # Cleanup: uninstall the discovery install
    Write-Log ""
    Uninstall-WinSCPFromDiscovery `
        -QuietUninstallString $uninstallEntry.QuietUninstallString `
        -UninstallString $uninstallEntry.UninstallString

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
