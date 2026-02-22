<#
Vendor: Microsoft
App: Microsoft Visual C++ 2015-2022 Redistributable (x86+x64)
CMName: Microsoft Visual C++ 2015-2022 Redistributable

.SYNOPSIS
    Packages Microsoft Visual C++ 2015-2022 Redistributable (x86+x64) for MECM.

.DESCRIPTION
    Downloads the latest vc_redist.x86.exe and vc_redist.x64.exe from Microsoft's
    permalink URLs, reads the version from VersionInfo on the x64 installer, stages
    both to a versioned network location, and creates an MECM Application with
    dual-registry detection.
    Detection uses AND logic: HKLM registry Version string under both X86 and X64
    VC\Runtimes keys must equal the expected value (vMAJOR.MINOR.BUILD.00).

    NOTE: The aka.ms permalink URLs always serve the current release. Installers
    are always re-downloaded to ensure the latest version is packaged.

    GetLatestVersionOnly downloads only the x64 installer to a local staging
    folder, reads the version from VersionInfo, outputs the short version string,
    and exits.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\VC++ 2015-2022 Redistributable\<Version>

.PARAMETER GetLatestVersionOnly
    Downloads the x64 installer, reads the version from VersionInfo, outputs the
    short version string, and exits. No MECM changes are made.

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
$UrlX86 = "https://aka.ms/vc14/vc_redist.x86.exe"
$UrlX64 = "https://aka.ms/vc14/vc_redist.x64.exe"

$FileNameX86 = "vc_redist.x86.exe"
$FileNameX64 = "vc_redist.x64.exe"

$VendorFolder = "Microsoft"
$AppFolder    = "VC++ 2015-2022 Redistributable"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager\MsvcRedist"

$EstimatedRuntimeMins = 15
$MaximumRuntimeMins   = 45

# Registry detection paths (stable across 2015-2022+ releases)
$RegKeyX86    = "SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X86"
$RegKeyX64    = "SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X64"
$RegValueName = "Version"

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

function Get-ExeFileVersion {
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Quiet
    )

    $vi = (Get-Item -LiteralPath $Path).VersionInfo
    if (-not $vi) { throw "Could not read VersionInfo from: $Path" }

    $fv = $vi.FileVersion
    $pv = $vi.ProductVersion

    if (-not [string]::IsNullOrWhiteSpace($fv)) { $fv = $fv.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($pv)) { $pv = $pv.Trim() }

    if (-not $Quiet) {
        Write-Host "EXE FileVersion              : $fv"
        Write-Host "EXE ProductVersion           : $pv"
    }

    if ($fv -match '^\d+\.\d+\.\d+\.\d+$') { return $fv }
    if ($pv -match '^\d+\.\d+\.\d+\.\d+$') { return $pv }

    throw "Could not determine quad version from VersionInfo for: $Path"
}

function Get-ShortVersionFromQuad {
    param([Parameter(Mandatory)][string]$QuadVersion)

    $parts = $QuadVersion -split '\.'
    if ($parts.Count -lt 3) { throw "Unexpected version format: $QuadVersion" }
    return ("{0}.{1}.{2}" -f $parts[0], $parts[1], $parts[2])
}

function Convert-ShortVersionToRegExpected {
    param([Parameter(Mandatory)][string]$ShortVersion)

    return ("v{0}.00" -f $ShortVersion)
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

function New-MECMMsvcRedistApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$RegExpected,
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
"%~dp0$FileNameX86" /install /quiet /norestart /log "%~dp0x86.install.log"
"%~dp0$FileNameX64" /install /quiet /norestart /log "%~dp0x64.install.log"
exit /b 0
"@
            Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
setlocal
"%~dp0$FileNameX86" /uninstall /quiet /norestart /log "%~dp0x86.uninstall.log"
"%~dp0$FileNameX64" /uninstall /quiet /norestart /log "%~dp0x64.uninstall.log"
exit /b 0
"@
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $dtName = $AppName

        $clauseX86 = New-CMDetectionClauseRegistryKeyValue `
            -Hive LocalMachine `
            -KeyName $RegKeyX86 `
            -ValueName $RegValueName `
            -PropertyType String `
            -Value `
            -ExpectedValue $RegExpected `
            -ExpressionOperator IsEquals

        $clauseX64 = New-CMDetectionClauseRegistryKeyValue `
            -Hive LocalMachine `
            -KeyName $RegKeyX64 `
            -ValueName $RegValueName `
            -PropertyType String `
            -Value `
            -ExpectedValue $RegExpected `
            -ExpressionOperator IsEquals

        Write-Host "Adding Script Deployment Type: $dtName"
        Write-Host "Detection (AND)              : HKLM\$RegKeyX86\$RegValueName == $RegExpected"
        Write-Host "Detection (AND)              : HKLM\$RegKeyX64\$RegValueName == $RegExpected"

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
            -AddDetectionClause @($clauseX86, $clauseX64) `
            -ContentFallback `
            -SlowNetworkDeploymentMode Download `
            -PostExecutionBehavior ForceReboot `
            -ErrorAction Stop | Out-Null

        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Host "Created MECM application     : $AppName"
    }
    finally {
        Set-Location $orig -ErrorAction SilentlyContinue
    }
}

function Get-MsvcRedistNetworkAppRoot {
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

        $localX64 = Join-Path $BaseDownloadRoot $FileNameX64

        curl.exe -L --fail --silent --show-error -o $localX64 $UrlX64
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $UrlX64" }

        $quadVersion = Get-ExeFileVersion -Path $localX64 -Quiet
        $shortVersion = Get-ShortVersionFromQuad -QuadVersion $quadVersion

        Write-Output $shortVersion
        exit 0
    }
    catch {
        Write-Error $_.Exception.Message
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Host ""
    Write-Host ("=" * 60)
    Write-Host "MSVC 2015-2022 Redistributable (x86+x64) Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Host ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Host "Start location               : $startLocation"
    Write-Host "SiteCode                     : $SiteCode"
    Write-Host "FileServerPath               : $FileServerPath"
    Write-Host "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Host "UrlX86                       : $UrlX86"
    Write-Host "UrlX64                       : $UrlX64"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-MsvcRedistNetworkAppRoot -FileServerPath $FileServerPath

    # Always re-download (permalink URLs serve latest release)
    $localX86 = Join-Path $BaseDownloadRoot $FileNameX86
    $localX64 = Join-Path $BaseDownloadRoot $FileNameX64

    Write-Host "Downloading x86 installer..."
    curl.exe -L --fail --silent --show-error -o $localX86 $UrlX86
    if ($LASTEXITCODE -ne 0) { throw "Download failed: $UrlX86" }

    Write-Host "Downloading x64 installer..."
    curl.exe -L --fail --silent --show-error -o $localX64 $UrlX64
    if ($LASTEXITCODE -ne 0) { throw "Download failed: $UrlX64" }

    # Version from x64 installer VersionInfo
    $quadVersion  = Get-ExeFileVersion -Path $localX64
    $shortVersion = Get-ShortVersionFromQuad -QuadVersion $quadVersion
    $regExpected  = Convert-ShortVersionToRegExpected -ShortVersion $shortVersion

    $contentPath = Join-Path $networkAppRoot $shortVersion

    Initialize-Folder -Path $contentPath

    $netX86 = Join-Path $contentPath $FileNameX86
    $netX64 = Join-Path $contentPath $FileNameX64

    Write-Host "Version (short)              : $shortVersion"
    Write-Host "Version (quad)               : $quadVersion"
    Write-Host "Registry expected            : $regExpected"
    Write-Host "Local x86                    : $localX86"
    Write-Host "Local x64                    : $localX64"
    Write-Host "ContentPath                  : $contentPath"
    Write-Host "Network x86                  : $netX86"
    Write-Host "Network x64                  : $netX64"
    Write-Host ""

    if (-not (Test-Path -LiteralPath $netX86)) {
        Write-Host "Copying x86 installer to network..."
        Copy-Item -LiteralPath $localX86 -Destination $netX86 -Force -ErrorAction Stop
    }
    else {
        Write-Host "Network x86 exists. Skipping copy."
    }

    if (-not (Test-Path -LiteralPath $netX64)) {
        Write-Host "Copying x64 installer to network..."
        Copy-Item -LiteralPath $localX64 -Destination $netX64 -Force -ErrorAction Stop
    }
    else {
        Write-Host "Network x64 exists. Skipping copy."
    }

    $appName   = "Microsoft Visual C++ 2015-2022 Redistributable (x86+x64) - $shortVersion"
    $publisher = "Microsoft Corporation"

    Write-Host ""
    Write-Host "CM Application Name          : $appName"
    Write-Host "CM SoftwareVersion           : $shortVersion"
    Write-Host ""

    New-MECMMsvcRedistApplication `
        -AppName $appName `
        -SoftwareVersion $shortVersion `
        -RegExpected $regExpected `
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
