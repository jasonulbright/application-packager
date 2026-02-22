<#
Vendor: Microsoft
App: Microsoft Power BI Desktop (x64)
CMName: Microsoft Power BI Desktop

.SYNOPSIS
    Packages Microsoft Power BI Desktop (x64) for MECM.

.DESCRIPTION
    Parses the Microsoft Download Center page to retrieve the current Power BI
    Desktop x64 installer URL and version, downloads the EXE, stages content to
    a versioned network location, and creates an MECM Application with file
    version-based detection.
    Detection uses PBIDesktop.exe file version GreaterEquals packaged version.

    NOTE: Power BI Desktop is released monthly. The installer URL embedded in the
    download page changes with each release; there is no stable permalink for the
    x64 EXE. Installer is always re-downloaded.

    GetLatestVersionOnly fetches only the Microsoft Download Center page (small
    HTML) to read the current version — no installer download is performed.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\Power BI Desktop\<Version>

.PARAMETER GetLatestVersionOnly
    Parses the Microsoft Download Center page for the current version, outputs
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
$DownloadPageUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=58494"
$ExeFileName     = "PBIDesktopSetup_x64.exe"

$VendorFolder = "Microsoft"
$AppFolder    = "Power BI Desktop"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager\PowerBI"

$EstimatedRuntimeMins = 15
$MaximumRuntimeMins   = 30

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

function Get-PowerBIDownloadInfo {
    param([switch]$Quiet)

    Write-Log "Download page URL            : $DownloadPageUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadPageUrl) -join "`n"
        if ($LASTEXITCODE -ne 0) { throw "Failed to fetch download page: $DownloadPageUrl" }

        $urlMatch = [regex]::Match($html, '"url"\s*:\s*"(https://download\.microsoft\.com/download/[^"]+PBIDesktopSetup_x64\.exe)"')
        if (-not $urlMatch.Success) { throw "Could not locate PBIDesktopSetup_x64.exe URL in download page." }

        $verMatch = [regex]::Match($html, '"[Vv]ersion"\s*:\s*"(\d+\.\d+\.\d+(?:\.\d+)?)"')
        if (-not $verMatch.Success) { throw "Could not parse version from download page." }

        $dlUrl   = $urlMatch.Groups[1].Value
        $version = $verMatch.Groups[1].Value

        Write-Log "Download URL                 : $dlUrl" -Quiet:$Quiet
        Write-Log "Latest Power BI version      : $version" -Quiet:$Quiet

        return [PSCustomObject]@{
            Version     = $version
            DownloadUrl = $dlUrl
        }
    }
    catch {
        Write-Log "Failed to get Power BI download info: $($_.Exception.Message)" -Level ERROR
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

function New-MECMPowerBIApplication {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$SoftwareVersion,
        [Parameter(Mandatory)][string]$ContentPath,
        [Parameter(Mandatory)][string]$InstallerFileName,
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
`$proc = Start-Process "`$PSScriptRoot\$InstallerFileName" -ArgumentList "-quiet ACCEPT_EULA=1" -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@
            Set-Content -LiteralPath $installPs1Path -Value $installPs1 -Encoding UTF8 -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallPs1Path)) {
            $uninstallPs1 = @'
$app = Get-ChildItem `
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' `
    -ErrorAction SilentlyContinue |
    Get-ItemProperty -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like 'Microsoft Power BI Desktop*' } |
    Sort-Object DisplayVersion -Descending |
    Select-Object -First 1
if ($app) {
    Start-Process msiexec.exe -ArgumentList "/x `"$($app.PSChildName)`" /qn /norestart" -Wait -NoNewWindow
}
'@
            Set-Content -LiteralPath $uninstallPs1Path -Value $uninstallPs1 -Encoding UTF8 -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
setlocal
PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
exit /b %ERRORLEVEL%
"@
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

        $dtName = $AppName

        $clause = New-CMDetectionClauseFile `
            -Path "$env:ProgramFiles\Microsoft Power BI Desktop\bin" `
            -FileName "PBIDesktop.exe" `
            -Value `
            -PropertyType Version `
            -ExpressionOperator GreaterEquals `
            -ExpectedValue $SoftwareVersion `
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

function Get-PowerBINetworkAppRoot {
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
        $info = Get-PowerBIDownloadInfo -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
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
    Write-Log "Microsoft Power BI Desktop (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "DownloadPageUrl              : $DownloadPageUrl"
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-PowerBINetworkAppRoot -FileServerPath $FileServerPath

    $info = Get-PowerBIDownloadInfo
    if (-not $info) {
        throw "Could not resolve Power BI Desktop download info."
    }

    $version     = $info.Version
    $downloadUrl = $info.DownloadUrl

    $localExe    = Join-Path $BaseDownloadRoot $ExeFileName
    $contentPath = Join-Path $networkAppRoot $version

    Initialize-Folder -Path $contentPath

    $netExe = Join-Path $contentPath $ExeFileName

    Write-Log "Version                      : $version"
    Write-Log "Local installer              : $localExe"
    Write-Log "ContentPath                  : $contentPath"
    Write-Log "Network installer            : $netExe"
    Write-Log ""

    Write-Log "Downloading installer..."
    Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localExe

    if (-not (Test-Path -LiteralPath $netExe)) {
        Write-Log "Copying installer to network..."
        Copy-Item -LiteralPath $localExe -Destination $netExe -Force -ErrorAction Stop
    }
    else {
        Write-Log "Network installer exists. Skipping copy."
    }

    $appName   = "Microsoft Power BI Desktop $version"
    $publisher = "Microsoft Corporation"

    Write-Log ""
    Write-Log "CM Application Name          : $appName"
    Write-Log "CM SoftwareVersion           : $version"
    Write-Log ""

    New-MECMPowerBIApplication `
        -AppName $appName `
        -SoftwareVersion $version `
        -ContentPath $contentPath `
        -InstallerFileName $ExeFileName `
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
