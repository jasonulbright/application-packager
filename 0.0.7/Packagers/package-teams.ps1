<#
Vendor: Microsoft
App: Microsoft Teams Enterprise (x64)
CMName: Microsoft Teams Enterprise

.SYNOPSIS
    Packages Microsoft Teams Enterprise (system-wide) MSIX for MECM.

.DESCRIPTION
    Downloads teamsbootstrapper.exe and the Teams Enterprise MSIX (x64) from
    Microsoft, stages both to a versioned network location, and creates an MECM
    Application with a PowerShell detection script.
    Detection uses Get-AppxPackage -AllUsers to check the provisioned MSTeams
    package version >= packaged version.

    Install:   teamsbootstrapper.exe -p -o "MSTeams-x64.msix"
    Uninstall: teamsbootstrapper.exe -x

    NOTE: The bootstrapper and MSIX are always re-downloaded to ensure the
    latest version is used. Version is read from AppxManifest.xml inside the
    MSIX.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\Teams\<Version>

.PARAMETER GetLatestVersionOnly
    Downloads the MSIX to a local staging folder, extracts the version from
    AppxManifest.xml, outputs the version string, and exits.

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
$BootstrapperUrl      = "https://go.microsoft.com/fwlink/?linkid=2243204"
$MsixUrl              = "https://go.microsoft.com/fwlink/?linkid=2196106"
$BootstrapperFileName = "teamsbootstrapper.exe"
$MsixFileName         = "MSTeams-x64.msix"
$AppXPackageName      = "MSTeams"

$VendorFolder = "Microsoft"
$AppFolder    = "Teams"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager\Teams"

$EstimatedRuntimeMins = 10
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

function Get-MsixVersion {
    param([Parameter(Mandatory)][string]$MsixPath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::OpenRead($MsixPath)
    try {
        $manifestEntry = $zip.Entries |
            Where-Object { $_.FullName -eq "AppxManifest.xml" } |
            Select-Object -First 1

        if (-not $manifestEntry) {
            throw "AppxManifest.xml not found inside MSIX: $MsixPath"
        }

        $reader = [System.IO.StreamReader]::new($manifestEntry.Open())
        try {
            $xmlContent = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $zip.Dispose()
    }

    $xml     = [xml]$xmlContent
    $version = $xml.Package.Identity.Version
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Version attribute missing or empty in AppxManifest.xml."
    }
    return $version
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

function New-MECMTeamsApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$ContentPath,
        [Parameter(Mandatory)][string]$Publisher
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
"%~dp0$BootstrapperFileName" -p -o "%~dp0$MsixFileName"
exit /b 0
"@
            Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
setlocal
"%~dp0$BootstrapperFileName" -x
exit /b 0
"@
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        # Write detection.ps1 to content folder for reference / manual testing
        $detectionPs1Path = Join-Path $ContentPath "detection.ps1"
        $detectionScript = @"
`$pkg = Get-AppxPackage -Name "$AppXPackageName" -AllUsers |
    Sort-Object { [version]`$_.Version } -Descending |
    Select-Object -First 1
if (`$pkg -and [version]`$pkg.Version -ge [version]"$SoftwareVersion") {
    Write-Output "Installed: `$(`$pkg.Version)"
}
"@
        if (-not (Test-Path -LiteralPath $detectionPs1Path)) {
            Set-Content -LiteralPath $detectionPs1Path -Value $detectionScript -Encoding UTF8 -ErrorAction Stop
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $dtName = $AppName

        Write-Host "Adding Script Deployment Type: $dtName"
        Add-CMScriptDeploymentType `
            -ApplicationName $AppName `
            -DeploymentTypeName $dtName `
            -ContentLocation $ContentPath `
            -InstallCommand "install.bat" `
            -UninstallCommand "uninstall.bat" `
            -ScriptLanguage PowerShell `
            -ScriptText $detectionScript `
            -InstallationBehaviorType InstallForSystem `
            -LogonRequirementType WhetherOrNotUserLoggedOn `
            -EstimatedRuntimeMins $EstimatedRuntimeMins `
            -MaximumRuntimeMins $MaximumRuntimeMins `
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

function Get-TeamsNetworkAppRoot {
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
        Initialize-Folder -Path $BaseDownloadRoot
        $tempMsix = Join-Path $BaseDownloadRoot $MsixFileName
        curl.exe -L --fail --silent --show-error -o $tempMsix $MsixUrl
        if ($LASTEXITCODE -ne 0) { throw "MSIX download failed: $MsixUrl" }
        $version = Get-MsixVersion -MsixPath $tempMsix
        Write-Output $version
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
    Write-Host "Microsoft Teams Enterprise (x64) Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Host ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Host "Start location               : $startLocation"
    Write-Host "SiteCode                     : $SiteCode"
    Write-Host "FileServerPath               : $FileServerPath"
    Write-Host "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Host "BootstrapperUrl              : $BootstrapperUrl"
    Write-Host "MsixUrl                      : $MsixUrl"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-TeamsNetworkAppRoot -FileServerPath $FileServerPath

    # Always re-download bootstrapper and MSIX to ensure latest version
    $localBootstrapper = Join-Path $BaseDownloadRoot $BootstrapperFileName
    $localMsix         = Join-Path $BaseDownloadRoot $MsixFileName

    Write-Host "Downloading bootstrapper..."
    curl.exe -L --fail --silent --show-error -o $localBootstrapper $BootstrapperUrl
    if ($LASTEXITCODE -ne 0) { throw "Bootstrapper download failed: $BootstrapperUrl" }

    Write-Host "Downloading MSIX..."
    curl.exe -L --fail --silent --show-error -o $localMsix $MsixUrl
    if ($LASTEXITCODE -ne 0) { throw "MSIX download failed: $MsixUrl" }

    $version = Get-MsixVersion -MsixPath $localMsix

    $contentPath = Join-Path $networkAppRoot $version

    Initialize-Folder -Path $contentPath

    $netBootstrapper = Join-Path $contentPath $BootstrapperFileName
    $netMsix         = Join-Path $contentPath $MsixFileName

    Write-Host "Version                      : $version"
    Write-Host "Local bootstrapper           : $localBootstrapper"
    Write-Host "Local MSIX                   : $localMsix"
    Write-Host "ContentPath                  : $contentPath"
    Write-Host "Network bootstrapper         : $netBootstrapper"
    Write-Host "Network MSIX                 : $netMsix"
    Write-Host ""

    # Always copy fresh downloads to network
    Write-Host "Copying content to network..."
    Copy-Item -LiteralPath $localBootstrapper -Destination $netBootstrapper -Force -ErrorAction Stop
    Copy-Item -LiteralPath $localMsix         -Destination $netMsix         -Force -ErrorAction Stop

    $appName   = "Microsoft Teams Enterprise - $version"
    $publisher = "Microsoft Corporation"

    Write-Host ""
    Write-Host "CM Application Name          : $appName"
    Write-Host "CM SoftwareVersion           : $version"
    Write-Host ""

    New-MECMTeamsApplication `
        -AppName $appName `
        -SoftwareVersion $version `
        -ContentPath $contentPath `
        -Publisher $publisher

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
