<#
Vendor: Microsoft
App: .NET 8 Desktop Runtime (x86+x64)
CMName: Microsoft Windows Desktop Runtime - 8

.SYNOPSIS
    Automates downloading .NET 8.0 Windows Desktop Runtime installers (x86 and x64) and creating an MECM application.
.DESCRIPTION
    Creates one MECM application:
    - Microsoft Windows Desktop Runtime (x86 and x64 combined) with one deployment type, file-based detection (x86 AND x64), and batch file installation.
    Application names match Programs and Features naming for consistency.
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
$DownloadBaseUrl = "https://dotnetcli.azureedge.net/dotnet/"

$TargetProducts = @(
    @{
        Name = "Microsoft Windows Desktop Runtime - {0}"
        FileNamePatterns = @("windowsdesktop-runtime-{0}-win-x64.exe", "windowsdesktop-runtime-{0}-win-x86.exe")
        UrlSegment = "WindowsDesktop"
        RootNetworkPath = $DesktopRuntimeRootNetworkPath
        ProgramsAndFeaturesName = "Microsoft Windows Desktop Runtime - {0} (x86)"
        RegistryKey = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        RegistryKey64 = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        DetectionType = "Folder"
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
    param(
        [Parameter(Mandatory)][string]$UrlSegment
    )

    $versionIndexUrl = "{0}{1}/Runtime/" -f $DownloadBaseUrl, $UrlSegment
    Write-Host "Fetching .NET 8 version list from: ${versionIndexUrl}"

    $pageHtml = (curl.exe -L --fail --silent --show-error $versionIndexUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch .NET 8 version list: $versionIndexUrl" }
    $versions = [regex]::Matches($pageHtml, 'href="(8\.0\.\d+/)"') |
        ForEach-Object { $_.Groups[1].Value.TrimEnd('/') } |
        Select-Object -Unique

    if (-not $versions) {
        throw "Could not find any .NET 8.0 versions under ${versionIndexUrl}"
    }

    $latest = $versions | Sort-Object { [version]$_ } -Descending | Select-Object -First 1
    if (-not $latest) {
        throw "Failed to determine latest .NET 8.0 version."
    }

    Write-Host "Latest .NET 8 runtime version: ${latest}"
    return $latest
}

function Get-NextDotnet8PatchVersion {
    param(
        [Parameter(Mandatory)][string]$CurrentVersion,
        [Parameter(Mandatory)][string[]]$AllVersions
    )

    $sorted = $AllVersions | Sort-Object { [version]$_ }
    $idx = [Array]::IndexOf($sorted, $CurrentVersion)
    if ($idx -ge 0 -and $idx -lt ($sorted.Count - 1)) {
        return $sorted[$idx + 1]
    }
    return $null
}

function Create-BatchFiles {
    param ([string]$NetworkPath, [string]$Version, [string]$ProductName)

    $originalLocation = Get-Location
    Write-Host "Current location before batch file creation: ${originalLocation}"
    try {
        Set-Location $PSScriptRoot -ErrorAction Stop
        Write-Host "Set location to script directory for batch file creation: ${PSScriptRoot}"

        $InstallBatContent = @"
start /wait "" "%~dp0windowsdesktop-runtime-${Version}-win-x64.exe" /install /quiet /norestart
start /wait "" "%~dp0windowsdesktop-runtime-${Version}-win-x86.exe" /install /quiet /norestart
"@

        $UninstallBatContent = @"
start /wait "" "%~dp0windowsdesktop-runtime-${Version}-win-x64.exe" /uninstall /quiet /norestart
start /wait "" "%~dp0windowsdesktop-runtime-${Version}-win-x86.exe" /uninstall /quiet /norestart
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
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$Publisher,
        [Parameter(Mandatory)][string]$ProductVersion,
        [Parameter(Mandatory)][string]$NetworkPath,
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$NextVersion,
        [Parameter(Mandatory)][string]$DetectionType
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
            Write-Host "Application '${AppName}' already exists. Skipping creation."
            return
        }

        Write-Host "Creating new MECM application: ${AppName}"
        $cmApp = New-CMApplication `
            -Name $AppName `
            -Publisher $Publisher `
            -SoftwareVersion $ProductVersion `
            -Description $Comment `
            -LocalizedApplicationName $AppName `
            -ErrorAction Stop

        Create-BatchFiles -NetworkPath $NetworkPath -Version $Version -ProductName $AppName

        $detectionClauses = @()

        if ($DetectionType -eq "Folder") {
            $x86Clause = New-CMDetectionClauseFile -Path "$env:ProgramFiles (x86)\dotnet\host\fxr\${Version}" -FileName "hostfxr.dll" -Existence -Is64Bit:$false
            $x64Clause = New-CMDetectionClauseFile -Path "$env:ProgramFiles\dotnet\host\fxr\${Version}" -FileName "hostfxr.dll" -Existence -Is64Bit

            $detectionClauses += $x86Clause, $x64Clause

            if ($NextVersion) {
                $x86NextClause = New-CMDetectionClauseFile -Path "$env:ProgramFiles (x86)\dotnet\host\fxr\${NextVersion}" -FileName "hostfxr.dll" -Existence -Is64Bit:$false
                $x64NextClause = New-CMDetectionClauseFile -Path "$env:ProgramFiles\dotnet\host\fxr\${NextVersion}" -FileName "hostfxr.dll" -Existence -Is64Bit
                $detectionClauses += $x86NextClause, $x64NextClause
            }
        }

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
            ApplicationName          = $AppName
            DeploymentTypeName       = "${AppName} Script DT"
            InstallCommand           = "install.bat"
            ContentLocation          = $NetworkPath
            UninstallCommand         = "uninstall.bat"
            InstallationBehaviorType = "InstallForSystem"
            LogonRequirementType     = "WhetherOrNotUserLoggedOn"
            MaximumRuntimeMins       = 30
            EstimatedRuntimeMins     = 10
            AddDetectionClause       = $detectionClauses
            ErrorAction              = "Stop"
        }

        Add-CMScriptDeploymentType @params | Out-Null
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

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

    foreach ($Product in $TargetProducts) {

        $allVersionsUrl = "{0}{1}/Runtime/" -f $DownloadBaseUrl, $Product.UrlSegment
        $AllVersionsHtml = (curl.exe -L --fail --silent --show-error $allVersionsUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch .NET 8 version list: $allVersionsUrl" }
        $AllVersions = [regex]::Matches($AllVersionsHtml, 'href="(8\.0\.\d+/)"') |
            ForEach-Object { $_.Groups[1].Value.TrimEnd('/') } |
            Select-Object -Unique

        $Version = Get-LatestDotnet8RuntimeVersion -UrlSegment $Product.UrlSegment
        if ($GetLatestVersionOnly) {
            Write-Output $Version
            return
        }

        $NextVersion = Get-NextDotnet8PatchVersion -CurrentVersion $Version -AllVersions $AllVersions

        $AppName = $Product.Name -f $Version
        $ProductVersion = $Version
        $DetectionType = $Product.DetectionType
        $RegistryKey = $Product.RegistryKey
        $RegistryKey64 = if ($Product.RegistryKey64) { $Product.RegistryKey64 } else { $RegistryKey }
        $Publisher = $Product.Publisher

        if (-not (Test-NetworkShareAccess -Path $Product.RootNetworkPath)) {
            Write-Error "Network share '${($Product.RootNetworkPath)}' is inaccessible. Skipping '${AppName}'."
            continue
        }

        $NetworkPath = Join-Path $Product.RootNetworkPath $Version
        if (-not (Test-Path $NetworkPath)) {
            Write-Host "Creating network directory: ${NetworkPath}"
            New-Item -ItemType Directory -Path $NetworkPath -Force -ErrorAction Stop | Out-Null
        }

        $DownloadFolderName = "dotnet8_${($Product.UrlSegment)}_${Version}_installers"
        $DownloadPath = Join-Path $BaseDownloadRoot $DownloadFolderName
        if (-not (Test-Path -LiteralPath $DownloadPath)) {
            New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
        }

        foreach ($pattern in $Product.FileNamePatterns) {
            $fileName = $pattern -f $Version
            $downloadUrl = "{0}{1}/Runtime/{2}/{3}" -f $DownloadBaseUrl, $Product.UrlSegment, $Version, $fileName

            $localFile = Join-Path $DownloadPath $fileName
            $networkFile = Join-Path $NetworkPath $fileName

            if (-not (Test-Path -LiteralPath $localFile)) {
                Write-Host "Downloading: ${downloadUrl}"
                curl.exe -L --fail --silent --show-error -o $localFile $downloadUrl
                if ($LASTEXITCODE -ne 0) { throw "Download failed: $downloadUrl" }
            }
            else {
                Write-Host "Local file exists, skipping download: ${localFile}"
            }

            if (-not (Test-Path -LiteralPath $networkFile)) {
                Copy-Item -LiteralPath $localFile -Destination $networkFile -Force -ErrorAction Stop
                Write-Host "Copied to network: ${networkFile}"
            }
            else {
                Write-Host "Network file exists, skipping copy: ${networkFile}"
            }
        }

        New-MECMApplication `
            -AppName $AppName `
            -Publisher $Publisher `
            -ProductVersion $ProductVersion `
            -NetworkPath $NetworkPath `
            -Version $Version `
            -NextVersion $NextVersion `
            -DetectionType $DetectionType
    }

    Write-Host "Script execution complete."
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    exit 1
}
