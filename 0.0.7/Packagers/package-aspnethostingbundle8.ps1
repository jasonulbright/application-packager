<#
Vendor: Microsoft
App: ASP.NET 8 Server Hosting Bundle (x64)
CMName: Microsoft .NET 8

.SYNOPSIS
    Automates downloading the .NET 8 ASP.NET Core Windows Server Hosting Bundle installer and creating an MECM application with registry-based detection.
.DESCRIPTION
    Creates one MECM application:
    - ASP.NET Core Hosting Bundle with registry-based detection.
    Application names match Programs and Features (e.g., "Microsoft .NET 8.0.x - Windows Server Hosting").
    Uses static metadata (Publisher: "Microsoft Corporation", SoftwareVersion: download version).
    MECM content settings are set on the deployment type.
.PARAMETER SiteCode
    ConfigMgr site code for the CM PSDrive (e.g., "MCM").
.PARAMETER Comment
    Work order or comment string applied to the MECM application description.
.PARAMETER FileServerPath
    UNC root of the SCCM content share (e.g., "\\fileserver\sccm$").
.PARAMETER GetLatestVersionOnly
    Outputs only the latest version string and exits.
.NOTES
    Requires:
      - PowerShell 5.1
      - ConfigMgr Admin Console installed (for ConfigurationManager.psd1)
      - RBAC rights to create Applications and Deployment Types
.KNOWN ISSUES
    Deployment type content download settings are set in Add-CMScriptDeploymentType parameters.
    Registry detection uses "SOFTWARE\WOW6432Node\Microsoft\ASP.NET Core\Shared Framework\v8.0\<Version>" key existence.
#>

param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Stop"

# --- Configuration Variables ---
$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads"

$DesktopRuntimeRootNetworkPath = Join-Path $FileServerPath "Applications\Microsoft\.NET Core"
$AspNetHostingBundleRootNetworkPath = Join-Path $FileServerPath "Applications\Microsoft\.NET Core"

$DotnetReleasesJsonUrl = "https://builds.dotnet.microsoft.com/dotnet/release-metadata/releases-index.json"
$DownloadBaseUrl = "https://dotnetcli.azureedge.net/dotnet/"

$TargetProducts = @(
    @{
        Name = "Microsoft .NET {0} - Windows Server Hosting"
        FileNamePatterns = @("dotnet-hosting-{0}-win.exe")
        UrlSegment = "aspnetcore/Runtime"
        RootNetworkPath = $AspNetHostingBundleRootNetworkPath
        ProgramsAndFeaturesName = "Microsoft .NET {0} - Windows Server Hosting"
        RegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        DetectionType = "Registry"
        Publisher = "Microsoft Corporation"
    }
)

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

function Test-NetworkShareAccess {
    param([Parameter(Mandatory)][string]$Path)

    $originalLocation = Get-Location
    try {
        if (-not (Test-Path -LiteralPath $Path -ErrorAction Stop)) {
            Write-Error "Network path '$Path' does not exist or is inaccessible."
            return $false
        }

        $testFile = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
        Set-Content -LiteralPath $testFile -Value "Test" -Encoding ASCII -ErrorAction Stop
        Remove-Item -LiteralPath $testFile -ErrorAction Stop

        return $true
    }
    catch {
        Write-Error "Failed to access network share '$Path': $($_.Exception.Message)"
        return $false
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
    }
}

function Get-LatestDotnet8RuntimeVersion {
    param([switch]$Quiet)

    if (-not $Quiet) {
        Write-Host "Fetching .NET release information from: ${DotnetReleasesJsonUrl}"
    }

    try {
        $JsonContent = (curl.exe -L --fail --silent --show-error $DotnetReleasesJsonUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch .NET release info: $DotnetReleasesJsonUrl" }
        $Releases = ConvertFrom-Json $JsonContent
        $Dotnet8ReleaseIndex = $Releases.'releases-index' | Where-Object { $_.'channel-version' -eq '8.0' -and $_.'release-type' -eq 'lts' } | Select-Object -First 1
        if ($Dotnet8ReleaseIndex -and $Dotnet8ReleaseIndex.'latest-runtime') {
            $version = $Dotnet8ReleaseIndex.'latest-runtime'
            if (-not $Quiet) {
                Write-Host "Found latest .NET 8.0 Runtime version: ${version}"
            }
            return $version
        }

        Write-Error "Could not find .NET 8.0 LTS release channel or latest runtime."
        exit 1
    }
    catch {
        Write-Error "Failed to get .NET 8.0 version: $($_.Exception.Message)"
        exit 1
    }
}

function Get-NextVersion {
    param([Parameter(Mandatory)][string]$CurrentVersion)

    $parts = $CurrentVersion -split '\.'
    if ($parts.Count -ne 3) { return $null }

    $major = $parts[0]
    $minor = $parts[1]
    $patch = [int]$parts[2] + 1
    return ("{0}.{1}.{2}" -f $major, $minor, $patch)
}

function Create-BatchFiles {
    param ([string]$NetworkPath, [string]$Version, [string]$ProductName)
    $originalLocation = Get-Location
    Write-Host "Current location before batch file creation: ${originalLocation}"
    try {
        Set-Location $PSScriptRoot -ErrorAction Stop
        Write-Host "Set location to script directory for batch file creation: ${PSScriptRoot}"

        $InstallBatContent = @"
start /wait "" "%~dp0dotnet-hosting-${Version}-win.exe" /install /quiet /norestart
"@
        $UninstallBatContent = @"
start /wait "" "%~dp0dotnet-hosting-${Version}-win.exe" /uninstall /quiet /norestart
"@

        $InstallBatPath = Join-Path $NetworkPath "install.bat"
        $UninstallBatPath = Join-Path $NetworkPath "uninstall.bat"
        Set-Content -Path $InstallBatPath -Value $InstallBatContent -Encoding ASCII -ErrorAction Stop
        Set-Content -Path $UninstallBatPath -Value $UninstallBatContent -Encoding ASCII -ErrorAction Stop
        Write-Host "Created install.bat and uninstall.bat at ${NetworkPath}"
    }
    catch {
        Write-Error "Failed to create batch files in '${NetworkPath}': $($_.Exception.Message)"
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
        [string]$NextVersion,
        [string]$Publisher,
        [string]$ProductVersion,
        [string]$DetectionType
    )
    $originalLocation = Get-Location
    Write-Host "Current location before MECM application creation: ${originalLocation}"
    Write-Host "Application Name: ${AppName}"
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
        -SoftwareVersion $ProductVersion `
        -LocalizedApplicationName $AppName `
        -Description $Comment `
        -ErrorAction Stop
}

        Create-BatchFiles -NetworkPath $NetworkPath -Version $Version -ProductName $AppName

        $detectionClauses = @()

        # ASP.NET Core Hosting Bundle - registry detection
        $registryKeyPath = "SOFTWARE\WOW6432Node\Microsoft\ASP.NET Core\Shared Framework\v8.0\${Version}"
        $registryClause = New-CMDetectionClauseRegistryKey -Hive LocalMachine -KeyName $registryKeyPath -Existence
        $detectionClauses += $registryClause

Write-Host "Calling Add-CMScriptDeploymentType with parameters:"
Write-Host "  ApplicationName: ${AppName}"
Write-Host "  DeploymentTypeName: ${AppName} Script DT"
Write-Host "  ContentLocation: ${NetworkPath}"
Write-Host "  InstallCommand: install.bat"
Write-Host "  UninstallCommand: uninstall.bat"
Write-Host "  InstallationBehaviorType: InstallForSystem"
Write-Host "  LogonRequirementType: WhetherOrNotUserLoggedOn"
Write-Host "  MaximumRuntimeMins: 30"
Write-Host "  EstimatedRuntimeMins: 10"
Write-Host "  Detection clauses count: $($detectionClauses.Count)"

$params = @{
    ApplicationName = $AppName
    DeploymentTypeName = "${AppName} Script DT"
    InstallCommand = "install.bat"
    ContentLocation = $NetworkPath
    UninstallCommand = "uninstall.bat"
    InstallationBehaviorType = "InstallForSystem"
    LogonRequirementType = "WhetherOrNotUserLoggedOn"
    MaximumRuntimeMins = 30
    EstimatedRuntimeMins = 10
    ContentFallback = $true
    SlowNetworkDeploymentMode = "Download"
    AddDetectionClause = $detectionClauses
    ErrorAction = "Stop"
}

Add-CMScriptDeploymentType @params | Out-Null
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$app.CI_ID) -KeepLatest 1

        Write-Host "Created MECM application: ${AppName} with ${DetectionType}-based detection"
    }
    catch {
        Write-Error "Failed to create MECM application: $($_.Exception.Message)"
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

    $LatestVersion = Get-LatestDotnet8RuntimeVersion
    if (-not $LatestVersion) { exit 1 }

    if ($GetLatestVersionOnly) {
        Write-Output $LatestVersion
        exit 0
    }

    $NextVersion = Get-NextVersion -CurrentVersion $LatestVersion

    $DownloadFolderName = "dotnet${LatestVersion}_installers"
    $DownloadPath = Join-Path $BaseDownloadRoot $DownloadFolderName
    if (-not (Test-Path $DownloadPath)) {
        Write-Host "Creating download directory: ${DownloadPath}"
        New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
    }

    foreach ($Product in $TargetProducts) {
        $AppName = $Product.Name -f $LatestVersion
        $NetworkPath = Join-Path $Product.RootNetworkPath $LatestVersion
        $ProgramsAndFeaturesName = $Product.ProgramsAndFeaturesName -f $LatestVersion
        $DetectionType = $Product.DetectionType
        $RegistryKey = $Product.RegistryKey
        $RegistryKey64 = if ($Product.RegistryKey64) { $Product.RegistryKey64 } else { $RegistryKey }
        $Publisher = $Product.Publisher

        if (-not (Test-NetworkShareAccess -Path $Product.RootNetworkPath)) {
            Write-Error "Network share '${($Product.RootNetworkPath)}' is inaccessible. Skipping '${AppName}'."
            continue
        }

        if (-not (Test-Path $NetworkPath)) {
            Write-Host "Creating network directory: ${NetworkPath}"
            New-Item -ItemType Directory -Path $NetworkPath -Force -ErrorAction Stop | Out-Null
        }

        foreach ($pattern in $Product.FileNamePatterns) {
            $fileName = $pattern -f $LatestVersion
            $downloadUrl = "{0}{1}/{2}/{3}" -f $DownloadBaseUrl, $Product.UrlSegment, $LatestVersion, $fileName

            $localFile = Join-Path $DownloadPath $fileName
            $networkFile = Join-Path $NetworkPath $fileName

            if (-not (Test-Path -LiteralPath $localFile)) {
                Write-Host "Downloading: ${downloadUrl}"
                try {
                    curl.exe -L --fail --silent --show-error -o $localFile $downloadUrl
                    if ($LASTEXITCODE -ne 0) { throw "Download failed: $downloadUrl" }
                    Write-Host "Downloaded: ${localFile}"
                }
                catch {
                    Write-Error "Failed to download ${fileName}: $($_.Exception.Message)"
                    continue
                }
            }
            else {
                Write-Host "${fileName} already exists locally: ${localFile}, skipping download."
            }

            if (-not (Test-Path -LiteralPath $networkFile)) {
                Write-Host "Copying ${fileName} to network share..."
                Copy-Item -LiteralPath $localFile -Destination $NetworkPath -Force -ErrorAction Stop
                Write-Host "Copied ${fileName} to ${NetworkPath}"
            }
            else {
                Write-Host "${fileName} already exists at network share: ${networkFile}, skipping copy."
            }
        }

        New-MECMApplication `
            -AppName $AppName `
            -Version $LatestVersion `
            -NetworkPath $NetworkPath `
            -NextVersion $NextVersion `
            -Publisher $Publisher `
            -ProductVersion $LatestVersion `
            -DetectionType $DetectionType
    }

    Write-Host ""
    Write-Host "Script execution complete."
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    exit 1
}
