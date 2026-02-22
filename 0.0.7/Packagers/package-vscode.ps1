<#
Vendor: Microsoft Corporation
App: VS Code
CMName: VS Code

.SYNOPSIS
    Automates downloading the latest VS Code installer and creating a MECM application.
.DESCRIPTION
    Creates a single MECM application for Visual Studio Code (x64), using version-based detection
    and batch file installation.
    Application name matches Programs and Features (e.g., "VS Code 1.104.2").
    Uses static metadata (Publisher: "Microsoft", SoftwareVersion: download version).
    MECM settings: 20-minute max runtime, 10-minute estimated runtime, system installation, no user logon requirement.
    Checks for existing files locally and on network share to skip redundant downloads/copies.
    Uses Add-CMScriptDeploymentType with version-based detection for Code.exe (>= 1.104.0) to resolve Tenable finding ID 265431.
    Sets content tab: Enables fallback source locations and sets slow network mode to "Download content from distribution point and run locally".
.NOTES
    - Run with admin privileges for MECM operations.
    - Customize $SiteCode for your MECM environment (e.g., PRD).
    - Requires PowerShell 5.1 and Configuration Manager module.
#>

param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

# --- Configuration Variables ---
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads"
$VSCodeRootNetworkPath = Join-Path $FileServerPath "Applications\Microsoft\Visual Studio Code"
$StaticDownloadUrl = "https://update.code.visualstudio.com/latest/win32-x64/stable"
$VersionApiUrl = "https://update.code.visualstudio.com/api/releases/stable"

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

function Get-LatestVSCodeVersion {
    param([switch]$Quiet)
    try {
        $json = (curl.exe -L --fail --silent --show-error $VersionApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch version info: $VersionApiUrl" }
        $versions = ConvertFrom-Json $json
        if (-not $versions -or $versions.Count -eq 0) {
            Write-Error "No versions found in API response."
            exit 1
        }
        $version = $versions[0]
        if (-not $Quiet) {
            Write-Host "Found latest VS Code version: $version"
        }
        return $version
    }
    catch {
        Write-Error "Failed to get VS Code version from API: $($_.Exception.Message)"
        exit 1
    }
}

function Test-NetworkShareAccess {
    param ([string]$Path)
    $originalLocation = Get-Location
    Write-Host "Current location before network share validation: ${originalLocation}"
    try {
        Set-Location $PSScriptRoot -ErrorAction Stop
        Write-Host "Set location to script directory for network share validation: ${PSScriptRoot}"
        if (-not (Test-Path $Path)) {
            Write-Error "Network path '${Path}' does not exist or is inaccessible."
            exit 1
        }

        $testFile = Join-Path $Path "_write_test_$(Get-Random).txt"
        Set-Content -Path $testFile -Value "Test" -Encoding ASCII
        Remove-Item $testFile -Force
        Write-Host "Network share '${Path}' is accessible and writable."
    }
    catch {
        Write-Error "Failed to access network share '${Path}': $($_.Exception.Message)"
        exit 1
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
        Write-Host "Restored location to: ${originalLocation}"
    }
}

function Create-BatchFiles {
    param ([string]$NetworkPath, [string]$Version, [string]$FileName)
    $originalLocation = Get-Location
    Write-Host "Current location before batch file creation: ${originalLocation}"
    try {
        Set-Location $PSScriptRoot -ErrorAction Stop
        Write-Host "Set location to script directory for batch file creation: ${PSScriptRoot}"
        Write-Host "Verifying Set-Content cmdlet: $(Get-Command Set-Content | Select-Object -ExpandProperty Source)"
        $InstallBatContent = @"
start /wait "" "%~dp0${FileName}" /VERYSILENT /NORESTART /FORCECLOSEAPPLICATIONS /MERGETASKS=!runcode
"@
        $UninstallBatContent = @"
"C:\Program Files\Microsoft VS Code\unins000.exe" /SILENT
"@
        $InstallBatPath = Join-Path $NetworkPath "install.bat"
        $UninstallBatPath = Join-Path $NetworkPath "uninstall.bat"
        Write-Host "Writing install.bat to ${InstallBatPath}"
        Set-Content -Path $InstallBatPath -Value $InstallBatContent -Encoding ASCII -ErrorAction Stop
        Write-Host "Writing uninstall.bat to ${UninstallBatPath}"
        Set-Content -Path $UninstallBatPath -Value $UninstallBatContent -Encoding ASCII -ErrorAction Stop
        Write-Host "Created install.bat and uninstall.bat at ${NetworkPath}"
    }
    catch {
        Write-Error "Failed to create batch files in '${NetworkPath}': $($_.Exception.Message)"
        throw
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

        Write-Host "Creating batch files for ${AppName}"
        Create-BatchFiles -NetworkPath $NetworkPath -Version $Version -FileName $FileName

        Write-Host "Creating version-based detection clause for Code.exe (>= 1.104.0)"
        $detectionClauses = @()
        $detectionClause = New-CMDetectionClauseFile -Path "$env:ProgramFiles\Microsoft VS Code" -FileName "Code.exe" -Value -PropertyType Version -ExpressionOperator GreaterEquals -ExpectedValue "1.104.0" -ErrorAction Stop
        $detectionClauses += $detectionClause

        Write-Host "Calling Add-CMScriptDeploymentType with parameters:"
        Write-Host "  ApplicationName: ${AppName}"
        Write-Host "  Detection clauses count: $($detectionClauses.Count)"

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
            ContentFallback = $true  # Enables fallback source locations (checks the box)
            SlowNetworkDeploymentMode = "Download"  # Download content from distribution point and run locally
            ErrorAction = "Stop"
        }

        Add-CMScriptDeploymentType @params
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$app.CI_ID) -KeepLatest 1
        Write-Host "Created MECM application: ${AppName} with version-based detection (>= 1.104.0) and content tab settings configured"
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

# --- Main Script ---
try {
    if (-not (Test-IsAdmin)) {
        Write-Error "This script must be run with admin privileges. Please run PowerShell as Administrator."
        exit 1
    }

    Set-Location $PSScriptRoot -ErrorAction Stop
    Write-Host "Set initial location to script directory: ${PSScriptRoot}"
    Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)"
    Write-Host "Verifying Set-Content cmdlet: $(Get-Command Set-Content | Select-Object -ExpandProperty Source)"

    if ($GetLatestVersionOnly) {
        $v = Get-LatestVSCodeVersion -Quiet
        Write-Output $v
        return
    }

    $Version = Get-LatestVSCodeVersion
    if (-not $Version) { exit 1 }

    $FileName = "VSCodeSetup-x64-${Version}.exe"
    $DownloadFolderName = "VSCode_${Version}_installers"
    $DownloadPath = Join-Path $BaseDownloadRoot $DownloadFolderName
    $NetworkPath = Join-Path $VSCodeRootNetworkPath $Version
    $AppName = "VS Code $Version"
    $Publisher = "Microsoft"

    Write-Host "DownloadPath: ${DownloadPath}"
    Write-Host "NetworkPath: ${NetworkPath}"

    Test-NetworkShareAccess -Path $VSCodeRootNetworkPath

    if (-not (Test-Path $DownloadPath)) {
        New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
        Write-Host "Created local download directory: ${DownloadPath}"
    }

    if (-not (Test-Path $NetworkPath)) {
        New-Item -ItemType Directory -Path $NetworkPath -Force | Out-Null
        Write-Host "Created network directory: ${NetworkPath}"
    }

    $OutputPath = Join-Path $DownloadPath $FileName
    $NetworkFilePath = Join-Path $NetworkPath $FileName
    Write-Host "OutputPath: ${OutputPath}"
    Write-Host "NetworkFilePath: ${NetworkFilePath}"

    if (-not (Test-Path $OutputPath)) {
        Write-Host "Downloading VS Code installer from ${StaticDownloadUrl}"
        curl.exe -L --fail --silent --show-error -o $OutputPath $StaticDownloadUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $StaticDownloadUrl" }
        Write-Host "Downloaded installer to ${OutputPath}"
    } else {
        Write-Host "Installer already exists locally: ${OutputPath}. Skipping download."
    }

    if (-not (Test-Path $NetworkFilePath)) {
        Write-Host "Copying installer to network path..."
        Copy-Item -Path $OutputPath -Destination $NetworkPath -Force -ErrorAction Stop
        Write-Host "Copied installer to ${NetworkFilePath}"
    } else {
        Write-Host "Installer already exists on network: ${NetworkFilePath}. Skipping copy."
    }

    Write-Host "Creating batch install/uninstall wrappers..."
    Create-BatchFiles -NetworkPath $NetworkPath -Version $Version -FileName $FileName

    Write-Host "Creating MECM application..."
    New-MECMApplication -AppName $AppName -Version $Version -NetworkPath $NetworkPath -FileName $FileName -Publisher $Publisher

    Write-Host "Script completed successfully for ${AppName}"
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}
