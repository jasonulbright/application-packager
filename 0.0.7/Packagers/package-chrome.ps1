<#
Vendor: Google
App: Google Chrome Enterprise (x64)
CMName: Google Chrome Enterprise

.SYNOPSIS
    Packages Google Chrome Enterprise (x64) MSI for MECM.

.DESCRIPTION
    Downloads the latest Google Chrome Enterprise x64 MSI from Google's static
    enterprise download URL, extracts ProductVersion and ProductCode via Windows
    Installer COM, stages content to a versioned network location, and creates
    an MECM Application with Windows Installer (ProductCode) detection.
    Detection uses ProductCode version IsEquals packaged version.

    NOTE: Google's enterprise MSI URL always serves the current stable release.
    The version is read from MSI properties after download.

    GetLatestVersionOnly queries Google's VersionHistory API (tiny JSON response)
    and exits without downloading the installer.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Google\Chrome Enterprise\<Version>

.PARAMETER GetLatestVersionOnly
    Queries the Chrome VersionHistory API for the current stable version, outputs
    the version string, and exits. No download or MECM changes are made.

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
$ChromeVersionApiUrl = "https://versionhistory.googleapis.com/v1/chrome/platforms/win64/channels/stable/versions?order_by=version+desc&pageSize=1"
$MsiDownloadUrl      = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
$MsiFileName         = "googlechromestandaloneenterprise64.msi"

$VendorFolder = "Google"
$AppFolder    = "Chrome Enterprise"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager\Chrome"

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

function Get-ChromeStableVersion {
    param([switch]$Quiet)

    Write-Log "Chrome version API URL        : $ChromeVersionApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $ChromeVersionApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Chrome version API: $ChromeVersionApiUrl" }

        $data = ConvertFrom-Json $json
        $version = $data.versions[0].version
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Could not parse Chrome stable version from API response."
        }

        Write-Log "Latest Chrome stable version : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get Chrome version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Get-MsiPropertyMap {
    param([Parameter(Mandatory)][string]$MsiPath)

    $installer = $null
    $db = $null
    $view = $null
    $record = $null

    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @($MsiPath, 0))

        $wanted = @("ProductVersion", "ProductCode")
        $map = @{}

        foreach ($p in $wanted) {
            $sql  = "SELECT `Value` FROM `Property` WHERE `Property`='$p'"
            $view = $db.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $db, @($sql))
            $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)

            if ($null -ne $record) {
                $val = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
                $map[$p] = $val
            }
            else {
                $map[$p] = $null
            }
        }

        return $map
    }
    finally {
        foreach ($o in @($record, $view, $db, $installer)) {
            if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) {
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) | Out-Null
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
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

function New-MECMChromeApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$ContentPath,
        [Parameter(Mandatory)][string]$MsiFileName,
        [Parameter(Mandatory)][string]$ProductCode,
        [Parameter(Mandatory)][string]$Publisher
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
`$proc = Start-Process msiexec.exe -ArgumentList "/i `"`$PSScriptRoot\$MsiFileName`" /qn /norestart" -Wait -PassThru -NoNewWindow
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
`$proc = Start-Process msiexec.exe -ArgumentList "/x `"`$PSScriptRoot\$MsiFileName`" /qn /norestart" -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@
            Set-Content -LiteralPath $uninstallPs1Path -Value $uninstallPs1 -Encoding UTF8 -ErrorAction Stop
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $dtName = $AppName

        $clause = New-CMDetectionClauseWindowsInstaller `
            -ProductCode $ProductCode `
            -Value `
            -ExpressionOperator IsEquals `
            -ExpectedValue $SoftwareVersion

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

function Get-ChromeNetworkAppRoot {
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
        $v = Get-ChromeStableVersion -Quiet
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
    Write-Log "Google Chrome Enterprise (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "MsiDownloadUrl               : $MsiDownloadUrl"
    Write-Log "ChromeVersionApiUrl          : $ChromeVersionApiUrl"
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-ChromeNetworkAppRoot -FileServerPath $FileServerPath

    # Always download MSI (URL serves latest stable release)
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName

    Write-Log "Downloading Chrome Enterprise MSI..."
    Invoke-DownloadWithRetry -Url $MsiDownloadUrl -OutFile $localMsi

    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productVersion = $props["ProductVersion"]
    $productCode    = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productVersion)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))    { throw "MSI ProductCode missing." }

    $contentPath = Join-Path $networkAppRoot $productVersion

    Initialize-Folder -Path $contentPath

    $netMsi = Join-Path $contentPath $MsiFileName

    Write-Log "Version                      : $productVersion"
    Write-Log "ProductCode                  : $productCode"
    Write-Log "Local MSI                    : $localMsi"
    Write-Log "ContentPath                  : $contentPath"
    Write-Log "Network MSI                  : $netMsi"
    Write-Log ""

    if (-not (Test-Path -LiteralPath $netMsi)) {
        Write-Log "Copying MSI to network..."
        Copy-Item -LiteralPath $localMsi -Destination $netMsi -Force -ErrorAction Stop
    }
    else {
        Write-Log "Network MSI exists. Skipping copy."
    }

    $appName   = "Google Chrome Enterprise $productVersion"
    $publisher = "Google LLC"

    Write-Log ""
    Write-Log "CM Application Name          : $appName"
    Write-Log "CM SoftwareVersion           : $productVersion"
    Write-Log ""

    New-MECMChromeApplication `
        -AppName $appName `
        -SoftwareVersion $productVersion `
        -ContentPath $contentPath `
        -MsiFileName $MsiFileName `
        -ProductCode $productCode `
        -Publisher $publisher

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
