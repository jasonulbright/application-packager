<#
Vendor: Microsoft
App: Microsoft Teams Enterprise
CMName: Microsoft Teams Enterprise

.SYNOPSIS
    Packages the latest Microsoft Teams Enterprise (system-wide) MSIX and creates an MECM Application.

.DESCRIPTION
    Downloads teamsbootstrapper.exe and the Teams Enterprise MSIX (x64) from Microsoft, stages
    both to a versioned network path, and creates an MECM Application with a Script Deployment
    Type. Detection uses a PowerShell script that checks the provisioned MSTeams AppX package
    version via Get-AppxPackage -AllUsers.

    Install:   teamsbootstrapper.exe -p -o "MSTeams-x64.msix"  (offline provision, all users)
    Uninstall: teamsbootstrapper.exe -x                          (deprovision for all users)
    Detection: Get-AppxPackage -Name "MSTeams" -AllUsers, version >= packaged version

    NOTE: The bootstrapper is always re-downloaded to ensure the latest version is used.
    The MSIX is also always re-downloaded to ensure the packaged version matches what is
    currently distributed by Microsoft. Version is read from AppxManifest.xml inside the MSIX.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Work order or comment string applied to the MECM application description.

.PARAMETER FileServerPath
    UNC root of the SCCM content share (e.g., "\\fileserver\sccm$").

.PARAMETER GetLatestVersionOnly
    Downloads the MSIX to a local staging folder, extracts the version from AppxManifest.xml,
    outputs the version string, and exits. No MECM changes are made.

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
$BootstrapperUrl      = "https://go.microsoft.com/fwlink/?linkid=2243204"
$MsixUrl              = "https://go.microsoft.com/fwlink/?linkid=2196106"
$BootstrapperFileName = "teamsbootstrapper.exe"
$MsixFileName         = "MSTeams-x64.msix"
$LocalDlRoot          = Join-Path $env:USERPROFILE "Downloads\TeamsEnterprise"
$TeamsNetworkRoot     = Join-Path $FileServerPath "Applications\Microsoft\Teams"
$Publisher            = "Microsoft Corporation"
$AppXPackageName      = "MSTeams"

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
        Set-Location C: -ErrorAction Stop
        if (-not (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) {
            Write-Error "Network path does not exist or is inaccessible: $Path"
            return $false
        }
        $testFile = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
        Set-Content -LiteralPath $testFile -Value "Test" -Encoding ASCII -ErrorAction Stop
        Remove-Item -LiteralPath $testFile -ErrorAction Stop
        Write-Host "Network share is accessible and writable: $Path"
        return $true
    }
    catch {
        Write-Error "Failed to access network share: $Path ($($_.Exception.Message))"
        return $false
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
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
        if (-not (Test-Path -LiteralPath $LocalDlRoot)) {
            New-Item -Path $LocalDlRoot -ItemType Directory -Force | Out-Null
        }
        $tempMsix = Join-Path $LocalDlRoot $MsixFileName
        Write-Host "Downloading $MsixFileName to extract version..."
        curl.exe -L --fail --silent --show-error -o $tempMsix $MsixUrl
        if ($LASTEXITCODE -ne 0) { throw "MSIX download failed: $MsixUrl" }
        $version = Get-MsixVersion -MsixPath $tempMsix
        Write-Output $version
        exit 0
    }
    catch {
        Write-Error "Failed to retrieve Teams version: $($_.Exception.Message)"
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

    # 1. Prepare local staging folder
    if (-not (Test-Path -LiteralPath $LocalDlRoot)) {
        New-Item -Path $LocalDlRoot -ItemType Directory -Force | Out-Null
    }

    $LocalBootstrapper = Join-Path $LocalDlRoot $BootstrapperFileName
    $LocalMsix         = Join-Path $LocalDlRoot $MsixFileName

    # 2. Download bootstrapper (always latest per Microsoft guidance)
    Write-Host "Downloading $BootstrapperFileName..." -ForegroundColor Cyan
    curl.exe -L --fail --silent --show-error -o $LocalBootstrapper $BootstrapperUrl
    if ($LASTEXITCODE -ne 0) { throw "Bootstrapper download failed: $BootstrapperUrl" }
    Write-Host "Downloaded: $LocalBootstrapper"

    # 3. Download MSIX (always re-download to ensure version accuracy)
    Write-Host "Downloading $MsixFileName..." -ForegroundColor Cyan
    curl.exe -L --fail --silent --show-error -o $LocalMsix $MsixUrl
    if ($LASTEXITCODE -ne 0) { throw "MSIX download failed: $MsixUrl" }
    Write-Host "Downloaded: $LocalMsix"

    # 4. Extract version from MSIX AppxManifest.xml
    Write-Host "Extracting version from MSIX..."
    $TeamsVersion = Get-MsixVersion -MsixPath $LocalMsix
    Write-Host "Teams version: $TeamsVersion" -ForegroundColor Green

    # 5. Validate network share access
    if (-not (Test-NetworkShareAccess -Path $TeamsNetworkRoot)) { exit 1 }

    Set-Location C: -ErrorAction Stop

    # 6. Create versioned content directory on network share
    $UncRoot = Join-Path $TeamsNetworkRoot $TeamsVersion
    if (-not (Test-Path -LiteralPath $UncRoot)) {
        Write-Host "Creating network directory: $UncRoot"
        New-Item -Path $UncRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    # 7. Copy content to network share
    Write-Host "Copying content to $UncRoot ..."
    Copy-Item -LiteralPath $LocalBootstrapper -Destination (Join-Path $UncRoot $BootstrapperFileName) -Force -ErrorAction Stop
    Copy-Item -LiteralPath $LocalMsix         -Destination (Join-Path $UncRoot $MsixFileName)         -Force -ErrorAction Stop
    Write-Host "Content staged."

    # 8. Write install.bat and uninstall.bat
    $InstallBat   = Join-Path $UncRoot "install.bat"
    $UninstallBat = Join-Path $UncRoot "uninstall.bat"

    $InstallBatContent = @"
@ECHO OFF
"%~dp0$BootstrapperFileName" -p -o "%~dp0$MsixFileName"
"@

    $UninstallBatContent = @"
@ECHO OFF
"%~dp0$BootstrapperFileName" -x
"@

    Set-Content -LiteralPath $InstallBat   -Value $InstallBatContent   -Encoding ASCII -ErrorAction Stop
    Set-Content -LiteralPath $UninstallBat -Value $UninstallBatContent -Encoding ASCII -ErrorAction Stop
    Write-Host "Created install.bat and uninstall.bat."

    # 9. Build PowerShell detection script
    #    MECM runs detection as SYSTEM. Get-AppxPackage -AllUsers works in SYSTEM context.
    #    The script outputs any non-empty string when detected; empty output = not detected.
    $DetectionScript = @"
`$pkg = Get-AppxPackage -Name "$AppXPackageName" -AllUsers |
    Sort-Object { [version]`$_.Version } -Descending |
    Select-Object -First 1
if (`$pkg -and [version]`$pkg.Version -ge [version]"$TeamsVersion") {
    Write-Output "Installed: `$(`$pkg.Version)"
}
"@

    # Write detection.ps1 to content folder for reference / manual testing
    $DetectionPs1 = Join-Path $UncRoot "detection.ps1"
    Set-Content -LiteralPath $DetectionPs1 -Value $DetectionScript -Encoding UTF8 -ErrorAction Stop
    Write-Host "Wrote detection.ps1 (reference copy)."

    # 10. Connect to Configuration Manager
    if (-not (Connect-CMSite -SiteCode $SiteCode)) {
        Write-Error "Cannot proceed without CM connection."
        exit 1
    }

    # 11. Application and deployment type names
    $AppName = "Microsoft Teams Enterprise - $TeamsVersion"
    $DTName  = "Microsoft Teams Enterprise - $TeamsVersion"

    # 12. Check for existing application
    $ExistingApp = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
    if ($ExistingApp) {
        Write-Warning "Application '$AppName' already exists (CI_ID: $($ExistingApp.CI_ID)). Exiting."
        exit 1
    }

    # 13. Create application
    Write-Host "Creating application '$AppName'..." -ForegroundColor Yellow
    $App = New-CMApplication `
        -Name $AppName `
        -Publisher $Publisher `
        -SoftwareVersion $TeamsVersion `
        -LocalizedApplicationName $AppName `
        -Description $Comment `
        -AutoInstall $true `
        -ErrorAction Stop

    Write-Host "Application CI_ID: $($App.CI_ID)"

    # 14. Add Script Deployment Type with PowerShell detection script
    #     ScriptLanguage + ScriptText selects the "custom script" detection method.
    #     No -AddDetectionClause is used; these are mutually exclusive parameter sets.
    Write-Host "Adding deployment type '$DTName' with PowerShell detection..."
    Add-CMScriptDeploymentType `
        -ApplicationName $AppName `
        -DeploymentTypeName $DTName `
        -InstallCommand "install.bat" `
        -UninstallCommand "uninstall.bat" `
        -ContentLocation $UncRoot `
        -ScriptLanguage PowerShell `
        -ScriptText $DetectionScript `
        -InstallationBehaviorType InstallForSystem `
        -LogonRequirementType WhetherOrNotUserLoggedOn `
        -UserInteractionMode Hidden `
        -MaximumRuntimeMins 30 `
        -EstimatedRuntimeMins 10 `
        -ContentFallback `
        -SlowNetworkDeploymentMode Download `
        -RebootBehavior NoAction `
        -ErrorAction Stop | Out-Null

    # 15. Revision history cleanup
    Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$App.CI_ID) -KeepLatest 1

    Write-Host "Microsoft Teams Enterprise $TeamsVersion packaged successfully." -ForegroundColor Green
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}
finally {
    Set-Location $originalLocation -ErrorAction SilentlyContinue
    Write-Host "Restored initial location to: ${originalLocation}"
}
