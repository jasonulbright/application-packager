<#
Vendor: Microsoft
App: Microsoft OLE DB Driver 19 for SQL Server (x64)
CMName: Microsoft OLE DB Driver 19 for SQL Server

.SYNOPSIS
    Packages Microsoft OLE DB Driver 19 for SQL Server (x64) for MECM.

.DESCRIPTION
    Downloads the Microsoft OLE DB Driver 19 (x64) installer, stages content to a
    versioned network location, and creates an MECM Application with MSI-based detection.

    This package uses:
      - Static, version-agnostic install.bat and uninstall.bat wrappers
      - Windows Installer (MSI) detection to enforce exact version alignment
      - System installation context with no user logon requirement

    Install and uninstall command files are created only if missing and are not
    regenerated on subsequent runs to prevent drift or accidental command changes.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available OLE DB Driver version string and exits.
    No MECM, network share, or administrative actions are performed when this switch is used.

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8.2
      - ConfigMgr Admin Console installed
      - RBAC permissions to create Applications and Deployment Types

    Detection:
      - Windows Installer (MSI) ProductCode + ProductVersion

    Behavior notes:
      - Static install/uninstall BAT files are intentional
      - MSI detection is preferred to ensure version control and compliance
#>


param(
    [string]$SiteCode       = "MCM",
    [string]$Comment        = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Configuration ---
$LearnDownloadPageUrl = "https://learn.microsoft.com/en-us/sql/connect/oledb/download-oledb-driver-for-sql-server?view=sql-server-ver17"
$FwLinkUrl            = "https://go.microsoft.com/fwlink/?linkid=2318101&clcid=0x409"

$BaseDownloadRoot = Join-Path $env:USERPROFILE "Downloads\_AutoPackager"
$NetworkRootPath  = Join-Path $FileServerPath "Applications\Microsoft\OLE DB Driver 19 for SQL Server"

$Publisher = "Microsoft Corporation"

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

function Resolve-OleDb19MsiUrl {
    param([switch]$Quiet)
    if (-not $Quiet) {
        Write-Host "Learn download page         : $LearnDownloadPageUrl"
        Write-Host "FWLink (English)            : $FwLinkUrl"
    }

    try {
        $final = (curl.exe --max-redirs 10 --silent --show-error --write-out "%{url_effective}" --output NUL $FwLinkUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to resolve URL: $FwLinkUrl" }

        if ([string]::IsNullOrWhiteSpace($final)) {
            throw "Could not resolve final MSI URL."
        }

        if ($final -notmatch '\.msi($|\?)') {
            throw "Resolved URL does not appear to be an MSI: $final"
        }

        if (-not $Quiet) { Write-Host "Resolved MSI URL            : $final" }
        return $final
    }
    catch {
        Write-Error "Failed to resolve MSI URL: $($_.Exception.Message)"
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

        $wanted = @("ProductName", "ProductVersion", "Manufacturer", "ProductCode")
        $map = @{}

        foreach ($p in $wanted) {
            $sql  = "SELECT `Value` FROM `Property` WHERE `Property`='$p'"
            $view = $db.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $db, @($sql))
            $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
            if ($record -ne $null) {
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
            if ($o -ne $null -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) {
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

function New-MECMOleDbMsiApplication {
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
            Write-Warning "Application already exists: $AppName"
            return
        }

        Write-Host "Creating CM Application      : $AppName"
        $cmApp = New-CMApplication -Name $AppName -Publisher $Publisher -SoftwareVersion $SoftwareVersion -Description $Comment -ErrorAction Stop

        # Static wrappers (do not overwrite if already present)
        $installBatPath   = Join-Path $ContentPath "install.bat"
        $uninstallBatPath = Join-Path $ContentPath "uninstall.bat"

        if (-not (Test-Path -LiteralPath $installBatPath)) {
            $installBat = @"
@echo off
setlocal
start /wait "" msiexec.exe /i "%~dp0$MsiFileName" /qn /norestart
exit /b 0
"@
            Set-Content -LiteralPath $installBatPath -Value $installBat -Encoding ASCII -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $uninstallBatPath)) {
            $uninstallBat = @"
@echo off
setlocal
start /wait "" msiexec.exe /x "%~dp0$MsiFileName" /qn /norestart
exit /b 0
"@
            Set-Content -LiteralPath $uninstallBatPath -Value $uninstallBat -Encoding ASCII -ErrorAction Stop
        }

        $dtName = $AppName

        # MSI detection (ProductCode + specific version)
        $clause = New-CMDetectionClauseWindowsInstaller `
            -ProductCode $ProductCode `
            -Value `
            -ExpressionOperator IsEquals `
            -ExpectedValue $SoftwareVersion

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
            -AddDetectionClause @($clause) `
            -ContentFallback $true `
            -SlowNetworkDeploymentMode Download `
            -ErrorAction Stop | Out-Null
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Host "Created MECM application     : $AppName"
    }
    finally {
        Set-Location $orig -ErrorAction SilentlyContinue
    }
}


# --- Latest-only mode ---
if ($GetLatestVersionOnly) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        Ensure-Folder -Path $BaseDownloadRoot

        $msiUrl = Resolve-OleDb19MsiUrl -Quiet
        if (-not $msiUrl) { exit 1 }

        $localMsi = Join-Path $BaseDownloadRoot "msoledbsql.msi"
        curl.exe -L --fail --silent --show-error -o $localMsi $msiUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $msiUrl" }

        $props = Get-MsiPropertyMap -MsiPath $localMsi
        if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) { exit 1 }

        Write-Output $props["ProductVersion"]
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
    Write-Host "Microsoft OLE DB Driver 19 (x64) Auto-Packager starting"
    Write-Host ("=" * 60)
    Write-Host ""
    Write-Host ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Host ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Host "Start location               : $startLocation"
    Write-Host "SiteCode                     : $SiteCode"
    Write-Host "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Host "NetworkRootPath              : $NetworkRootPath"
    Write-Host "LearnDownloadPageUrl         : $LearnDownloadPageUrl"
    Write-Host "FWLinkUrl                    : $FwLinkUrl"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Error "Run PowerShell as Administrator."
        exit 1
    }

    Ensure-Folder -Path $BaseDownloadRoot

    if (-not (Test-NetworkShareAccess -Path $NetworkRootPath)) {
        throw "Network root path not accessible: $NetworkRootPath"
    }

    $msiUrl = Resolve-OleDb19MsiUrl
    if (-not $msiUrl) {
        throw "Could not resolve MSI download URL."
    }

    $localMsi = Join-Path $BaseDownloadRoot "msoledbsql.msi"

    Write-Host "Local MSI path               : $localMsi"
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

    Write-Host "MSI ProductName              : $productName"
    Write-Host "MSI ProductVersion           : $productVersion"
    Write-Host "MSI Manufacturer             : $manufacturer"
    Write-Host "MSI ProductCode              : $productCode"
    Write-Host ""

    $versionFolder = $productVersion
    $contentPath   = Join-Path $NetworkRootPath $versionFolder

    Ensure-Folder -Path $contentPath

    $msiFileName = "msoledbsql.msi"
    $netMsi      = Join-Path $contentPath $msiFileName

    Write-Host "ContentPath                  : $contentPath"
    Write-Host "Network MSI                  : $netMsi"
    Write-Host ""

    if (-not (Test-Path -LiteralPath $netMsi)) {
        Write-Host "Copying MSI to network..."
        Copy-Item -LiteralPath $localMsi -Destination $netMsi -Force -ErrorAction Stop
    }
    else {
        Write-Host "Network MSI exists. Skipping copy."
    }

    $appName = "$productName $productVersion"

    Write-Host ""
    Write-Host "CM Application Name          : $appName"
    Write-Host "CM SoftwareVersion           : $productVersion"
    Write-Host ""

    New-MECMOleDbMsiApplication `
        -AppName $appName `
        -SoftwareVersion $productVersion `
        -ContentPath $contentPath `
        -MsiFileName $msiFileName `
        -ProductCode $productCode `
        -Publisher $Publisher

    Write-Host ""
    Write-Host "Script execution complete."
}
catch {
    Write-Error "SCRIPT FAILED: $($_.Exception.Message)"
    exit 1
}
