<#
Vendor: Microsoft Corporation
App: .NET 8 Desktop Runtime (x86+x64)
CMName: Microsoft Windows Desktop Runtime - 8

.SYNOPSIS
    Packages .NET 8 Windows Desktop Runtime (x86 and x64) for MECM.

.DESCRIPTION
    Downloads the latest .NET 8 Windows Desktop Runtime installers for both
    x86 and x64 from the official Microsoft CDN, stages content to a versioned
    network location, and creates an MECM Application with file-based detection.
    Detection uses hostfxr.dll existence in the version-specific fxr path for
    both architectures (x86 AND x64).

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\.NET Core\<Version>

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available .NET 8 runtime version string and exits.

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
$ReleasesIndexUrl  = "https://builds.dotnet.microsoft.com/dotnet/release-metadata/releases-index.json"
$DownloadUrlBase   = "https://dotnetcli.azureedge.net/dotnet/WindowsDesktop/Runtime"

$VendorFolder = "Microsoft"
$AppFolder    = ".NET Core"

$X64FileNamePattern = "windowsdesktop-runtime-{0}-win-x64.exe"
$X86FileNamePattern = "windowsdesktop-runtime-{0}-win-x86.exe"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager\DotNet8DesktopRuntime"

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

function Get-LatestDotNet8Version {
    param([switch]$Quiet)

    if (-not $Quiet) {
        Write-Host "Releases index URL           : $ReleasesIndexUrl"
    }

    try {
        $json = (curl.exe -L --fail --silent --show-error $ReleasesIndexUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch .NET release info: $ReleasesIndexUrl" }

        $releases = ConvertFrom-Json $json
        $dotnet8Channel = $releases.'releases-index' |
            Where-Object { $_.'channel-version' -eq '8.0' -and $_.'release-type' -eq 'lts' } |
            Select-Object -First 1

        if (-not $dotnet8Channel -or -not $dotnet8Channel.'latest-runtime') {
            throw "Could not find .NET 8.0 LTS release channel or latest runtime."
        }

        $version = $dotnet8Channel.'latest-runtime'

        if (-not $Quiet) {
            Write-Host "Latest .NET 8 runtime version: $version"
        }
        return $version
    }
    catch {
        Write-Error "Failed to get .NET 8 version: $($_.Exception.Message)"
        return $null
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
        if ($h.PSObject.Properties.Name -contains 'Revision') { $revs += [UInt32]$h.Revision; continue }
        if ($h.PSObject.Properties.Name -contains 'CIVersion') { $revs += [UInt32]$h.CIVersion; continue }
    }

    $revs = $revs | Sort-Object -Unique -Descending
    if ($revs.Count -le $KeepLatest) { return }

    foreach ($rev in ($revs | Select-Object -Skip $KeepLatest)) {
        Remove-CMApplicationRevisionHistory -Id $CI_ID -Revision $rev -Force -ErrorAction Stop
    }
}

function New-MECMDotNet8DesktopRuntimeApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$ContentPath,
        [Parameter(Mandatory)][string]$X64FileName,
        [Parameter(Mandatory)][string]$X86FileName,
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
start /wait "" "%~dp0$X64FileName" /install /quiet /norestart
start /wait "" "%~dp0$X86FileName" /install /quiet /norestart
exit /b 0
"@
            Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
setlocal
start /wait "" "%~dp0$X64FileName" /uninstall /quiet /norestart
start /wait "" "%~dp0$X86FileName" /uninstall /quiet /norestart
exit /b 0
"@
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $dtName = $AppName

        # Detection: hostfxr.dll existence in version-specific fxr paths (x86 AND x64)
        $x86Clause = New-CMDetectionClauseFile `
            -Path "$env:ProgramFiles (x86)\dotnet\host\fxr\${SoftwareVersion}" `
            -FileName "hostfxr.dll" `
            -Existence `
            -Is64Bit:$false

        $x64Clause = New-CMDetectionClauseFile `
            -Path "$env:ProgramFiles\dotnet\host\fxr\${SoftwareVersion}" `
            -FileName "hostfxr.dll" `
            -Existence `
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
            -AddDetectionClause @($x86Clause, $x64Clause) `
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

function Get-DotNet8DesktopRuntimeNetworkAppRoot {
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
        $v = Get-LatestDotNet8Version -Quiet
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
    Write-Host ".NET 8 Desktop Runtime (x86+x64) Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Host ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Host "Start location               : $startLocation"
    Write-Host "SiteCode                     : $SiteCode"
    Write-Host "FileServerPath               : $FileServerPath"
    Write-Host "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Host "ReleasesIndexUrl             : $ReleasesIndexUrl"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-DotNet8DesktopRuntimeNetworkAppRoot -FileServerPath $FileServerPath

    $version = Get-LatestDotNet8Version
    if (-not $version) {
        throw "Could not resolve .NET 8 runtime version."
    }

    $x64FileName = $X64FileNamePattern -f $version
    $x86FileName = $X86FileNamePattern -f $version
    $contentPath = Join-Path $networkAppRoot $version

    Initialize-Folder -Path $contentPath

    $localX64  = Join-Path $BaseDownloadRoot $x64FileName
    $localX86  = Join-Path $BaseDownloadRoot $x86FileName
    $netX64    = Join-Path $contentPath $x64FileName
    $netX86    = Join-Path $contentPath $x86FileName

    Write-Host "Version                      : $version"
    Write-Host "Local x64 installer          : $localX64"
    Write-Host "Local x86 installer          : $localX86"
    Write-Host "ContentPath                  : $contentPath"
    Write-Host "Network x64 installer        : $netX64"
    Write-Host "Network x86 installer        : $netX86"
    Write-Host ""

    # Download x64
    if (-not (Test-Path -LiteralPath $localX64)) {
        Write-Host "Downloading x64 installer..."
        $downloadUrl = "${DownloadUrlBase}/${version}/${x64FileName}"
        curl.exe -L --fail --silent --show-error -o $localX64 $downloadUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $downloadUrl" }
    }
    else {
        Write-Host "Local x64 installer exists. Skipping download."
    }

    # Download x86
    if (-not (Test-Path -LiteralPath $localX86)) {
        Write-Host "Downloading x86 installer..."
        $downloadUrl = "${DownloadUrlBase}/${version}/${x86FileName}"
        curl.exe -L --fail --silent --show-error -o $localX86 $downloadUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $downloadUrl" }
    }
    else {
        Write-Host "Local x86 installer exists. Skipping download."
    }

    # Copy x64 to network
    if (-not (Test-Path -LiteralPath $netX64)) {
        Write-Host "Copying x64 installer to network..."
        Copy-Item -LiteralPath $localX64 -Destination $netX64 -Force -ErrorAction Stop
    }
    else {
        Write-Host "Network x64 installer exists. Skipping copy."
    }

    # Copy x86 to network
    if (-not (Test-Path -LiteralPath $netX86)) {
        Write-Host "Copying x86 installer to network..."
        Copy-Item -LiteralPath $localX86 -Destination $netX86 -Force -ErrorAction Stop
    }
    else {
        Write-Host "Network x86 installer exists. Skipping copy."
    }

    $appName   = "Microsoft Windows Desktop Runtime - ${version}"
    $publisher = "Microsoft Corporation"

    Write-Host ""
    Write-Host "CM Application Name          : $appName"
    Write-Host "CM SoftwareVersion           : $version"
    Write-Host ""

    New-MECMDotNet8DesktopRuntimeApplication `
        -AppName $appName `
        -SoftwareVersion $version `
        -ContentPath $contentPath `
        -X64FileName $x64FileName `
        -X86FileName $x86FileName `
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
