<#
Vendor: Microsoft
App: Microsoft ODBC Driver 18 for SQL Server (x64)
CMName: Microsoft ODBC Driver 18 for SQL Server

.SYNOPSIS
    Packages Microsoft ODBC Driver 18 for SQL Server (x64) for MECM.

.DESCRIPTION
    Downloads the Microsoft ODBC Driver 18 (x64) installer, stages content to a
    versioned network location, and creates an MECM Application with MSI-based detection.

    This package uses:
      - Static, version-agnostic install.bat and uninstall.bat wrappers
      - Windows Installer (MSI) detection to enforce exact version alignment
      - System installation context with no user logon requirement

    Install and uninstall command files are created only if missing and are not
    regenerated on subsequent runs to prevent drift or accidental command changes.

    GetLatestVersionOnly fetches only the Microsoft documentation page (small HTML)
    to read the current version — no installer download is performed.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Work order or comment string applied to the MECM application description.

.PARAMETER FileServerPath
    UNC root of the SCCM content share (e.g., "\\fileserver\sccm$").

.PARAMETER GetLatestVersionOnly
    Fetches the current ODBC 18 version from the Microsoft documentation page,
    outputs the version string, and exits. No MECM, network share, or download
    actions are performed.

.NOTES
    Requirements:
      - PowerShell 5.1
      - ConfigMgr Admin Console installed (for ConfigurationManager.psd1)
      - RBAC permissions to create Applications and Deployment Types
      - Local administrator
      - Write access to FileServerPath

    Detection:
      - Windows Installer (MSI) ProductCode + ProductVersion (IsEquals)
#>

param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Configuration ---
$LearnPageUrl         = "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server"
$FwLinkUrl            = "https://go.microsoft.com/fwlink/?linkid=2345415&clcid=0x409"
$BaseDownloadRoot     = Join-Path $env:USERPROFILE "Downloads\_AutoPackager"
$NetworkRootPath      = Join-Path $FileServerPath "Applications\Microsoft\ODBC Driver 18 for SQL Server"
$Publisher            = "Microsoft Corporation"
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
        Write-Warning "Admin check failed: $($_.Exception.Message)"
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

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
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
        $tmp = Join-Path $Path ("_write_test_{0}.txt" -f (Get-Random))
        Set-Content -LiteralPath $tmp -Value "test" -Encoding ASCII -ErrorAction Stop
        Remove-Item -LiteralPath $tmp -ErrorAction Stop
        return $true
    }
    catch {
        Write-Error "Network share is not writable: $Path ($($_.Exception.Message))"
        return $false
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
    }
}

function Get-Odbc18Version {
    # Lightweight: parses the Microsoft documentation page — no installer download.
    param([switch]$Quiet)
    if (-not $Quiet) { Write-Host "Fetching ODBC 18 version from: $LearnPageUrl" }
    $html = (curl.exe -L --fail --silent --show-error $LearnPageUrl) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "Failed to fetch ODBC documentation page: $LearnPageUrl" }
    $verMatch = [regex]::Match($html, '\b(18\.\d+\.\d+\.\d+)\b')
    if (-not $verMatch.Success) { throw "Could not parse ODBC 18 version from documentation page." }
    $version = $verMatch.Groups[1].Value
    if (-not $Quiet) { Write-Host "Latest ODBC Driver 18 version: $version" }
    return $version
}

function Resolve-Odbc18MsiUrl {
    param([switch]$Quiet)
    if (-not $Quiet) {
        Write-Host "Learn download page : $LearnPageUrl"
        Write-Host "FWLink (English)    : $FwLinkUrl"
    }
    $final = (curl.exe --max-redirs 10 --silent --show-error --write-out "%{url_effective}" --output NUL $FwLinkUrl) -join ''
    if ($LASTEXITCODE -ne 0) { throw "Failed to resolve URL: $FwLinkUrl" }
    if ([string]::IsNullOrWhiteSpace($final)) { throw "Could not resolve final MSI URL." }
    if ($final -notmatch '\.msi($|\?)') { throw "Resolved URL does not appear to be an MSI: $final" }
    if (-not $Quiet) { Write-Host "Resolved MSI URL    : $final" }
    return $final
}

function Get-MsiPropertyMap {
    param([Parameter(Mandatory)][string]$MsiPath)
    $installer = $null
    $db        = $null
    $view      = $null
    $record    = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @($MsiPath, 0))
        $wanted = @("ProductName", "ProductVersion", "Manufacturer", "ProductCode")
        $map = @{}
        foreach ($p in $wanted) {
            $sql    = "SELECT ``Value`` FROM ``Property`` WHERE ``Property``='$p'"
            $view   = $db.GetType().InvokeMember("OpenView",   "InvokeMethod", $null, $db,   @($sql))
            $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null
            $record = $view.GetType().InvokeMember("Fetch",     "InvokeMethod", $null, $view, $null)
            if ($null -ne $record) {
                $map[$p] = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
            } else {
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
        $version = Get-Odbc18Version -Quiet
        Write-Output $version
        exit 0
    }
    catch {
        Write-Error "Failed to retrieve ODBC 18 version: $($_.Exception.Message)"
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Host ""
    Write-Host ("=" * 60)
    Write-Host "Microsoft ODBC Driver 18 (x64) Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser         : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    Write-Host ("Machine           : {0}"     -f $env:COMPUTERNAME)
    Write-Host "Start location    : $startLocation"
    Write-Host "SiteCode          : $SiteCode"
    Write-Host "BaseDownloadRoot  : $BaseDownloadRoot"
    Write-Host "NetworkRootPath   : $NetworkRootPath"
    Write-Host "LearnPageUrl      : $LearnPageUrl"
    Write-Host "FWLinkUrl         : $FwLinkUrl"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    Ensure-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $NetworkRootPath)) {
        throw "Network root path not accessible: $NetworkRootPath"
    }

    $msiUrl = Resolve-Odbc18MsiUrl
    if (-not $msiUrl) { throw "Could not resolve MSI download URL." }

    $localMsi = Join-Path $BaseDownloadRoot "msodbcsql18.msi"
    Write-Host "Local MSI path    : $localMsi"
    Write-Host ""

    Write-Host "Downloading MSI..."
    curl.exe -L --fail --silent --show-error -o $localMsi $msiUrl
    if ($LASTEXITCODE -ne 0) { throw "Download failed: $msiUrl" }

    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName    = $props["ProductName"]
    $productVersion = $props["ProductVersion"]
    $manufacturer   = $props["Manufacturer"]
    $productCode    = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productName))    { throw "MSI ProductName missing." }
    if ([string]::IsNullOrWhiteSpace($productVersion)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))    { throw "MSI ProductCode missing." }

    Write-Host "MSI ProductName    : $productName"
    Write-Host "MSI ProductVersion : $productVersion"
    Write-Host "MSI Manufacturer   : $manufacturer"
    Write-Host "MSI ProductCode    : $productCode"
    Write-Host ""

    $contentPath = Join-Path $NetworkRootPath $productVersion
    Ensure-Folder -Path $contentPath

    $msiFileName = "msodbcsql18.msi"
    $netMsi      = Join-Path $contentPath $msiFileName

    Write-Host "ContentPath       : $contentPath"
    Write-Host "Network MSI       : $netMsi"
    Write-Host ""

    if (-not (Test-Path -LiteralPath $netMsi)) {
        Write-Host "Copying MSI to network..."
        Copy-Item -LiteralPath $localMsi -Destination $netMsi -Force -ErrorAction Stop
    } else {
        Write-Host "Network MSI exists. Skipping copy."
    }

    $appName = "$productName $productVersion"

    Write-Host "CM Application Name : $appName"
    Write-Host ""

    if (-not (Connect-CMSite -SiteCode $SiteCode)) { throw "CM site connection failed." }

    $existing = Get-CMApplication -Name $appName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Warning "Application already exists: $appName"
        exit 1
    }

    Write-Host "Creating CM Application: $appName"
    $cmApp = New-CMApplication `
        -Name $appName `
        -Publisher $Publisher `
        -SoftwareVersion $productVersion `
        -Description $Comment `
        -ErrorAction Stop

    # Static wrappers (do not overwrite if already present)
    $installBatPath   = Join-Path $contentPath "install.bat"
    $uninstallBatPath = Join-Path $contentPath "uninstall.bat"

    if (-not (Test-Path -LiteralPath $installBatPath)) {
        $installBat = @"
@echo off
setlocal
start /wait "" msiexec.exe /i "%~dp0$msiFileName" /qn /norestart
exit /b 0
"@
        Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
    }

    if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
        $uninstallBat = @"
@echo off
setlocal
start /wait "" msiexec.exe /x "%~dp0$msiFileName" /qn /norestart
exit /b 0
"@
        Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
    }

    $clause = New-CMDetectionClauseWindowsInstaller `
        -ProductCode $productCode `
        -Value `
        -ExpressionOperator IsEquals `
        -ExpectedValue $productVersion

    Write-Host "Adding Script Deployment Type: $appName"
    Add-CMScriptDeploymentType `
        -ApplicationName $appName `
        -DeploymentTypeName $appName `
        -ContentLocation $contentPath `
        -InstallCommand "install.bat" `
        -UninstallCommand "uninstall.bat" `
        -InstallationBehaviorType InstallForSystem `
        -LogonRequirementType WhetherOrNotUserLoggedOn `
        -UserInteractionMode Hidden `
        -EstimatedRuntimeMins $EstimatedRuntimeMins `
        -MaximumRuntimeMins $MaximumRuntimeMins `
        -AddDetectionClause @($clause) `
        -ContentFallback $true `
        -SlowNetworkDeploymentMode Download `
        -ErrorAction Stop | Out-Null

    Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

    Write-Host ""
    Write-Host "Created MECM application: $appName"
    Write-Host "Script execution complete."
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    exit 1
}
