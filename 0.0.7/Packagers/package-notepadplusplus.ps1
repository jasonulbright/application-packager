<#
Vendor: Notepad++ Team
App: Notepad++
CMName: Notepad++

.SYNOPSIS
    Packages the latest Notepad++ release and creates an MECM Application.

.DESCRIPTION
    Retrieves the latest Notepad++ release metadata from GitHub, stages content to a versioned
    network path, and creates an MECM Application and Script Deployment Type.

.PARAMETER SiteCode
    ConfigMgr site code for the CM PSDrive.

.PARAMETER Comment
    Value written to the Application Description field.

.PARAMETER FileServerPath
    UNC root for content storage. Content is staged under:
    <FileServerPath>\Applications\<Vendor>\<App>\<Version>

.PARAMETER GetLatestVersionOnly
    Outputs only the latest version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager module available)
    - RBAC rights to create and modify Applications / Deployment Types
    - Local administrator
    - Write access to FileServerPath
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment  = "WO#00000001234567",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [switch]$GetLatestVersionOnly
)

# --- Global Configuration & Variables ---
$ErrorActionPreference = 'Stop'
$GitHubApiUrl   = "https://api.github.com/repos/notepad-plus-plus/notepad-plus-plus/releases/latest"
$LocalDlRoot    = Join-Path $env:USERPROFILE "Downloads\NotepadPlusPlusInstallers"
$Publisher      = "Don Ho"
$VendorFolder   = "Notepad++"
$AppFolder      = "Notepad++"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Functions ---
function Test-IsAdmin {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
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

        if ([string]::IsNullOrWhiteSpace($Path)) {
            Write-Error "Network path is null or empty."
            return $false
        }

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

function Ensure-ContentRootFolders {
    param(
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$Vendor,
        [Parameter(Mandatory)][string]$App
    )

    $originalLocation = Get-Location
    try {
        Set-Location C: -ErrorAction Stop

        $vendorPath = Join-Path $FileServerPath $Vendor
        if (-not (Test-Path -LiteralPath $vendorPath)) {
            New-Item -Path $vendorPath -ItemType Directory -Force | Out-Null
        }

        $appPath = Join-Path $vendorPath $App
        if (-not (Test-Path -LiteralPath $appPath)) {
            New-Item -Path $appPath -ItemType Directory -Force | Out-Null
        }

        return $appPath
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
    }
}

function Invoke-RevisionHistoryCleanup {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$AppNamePattern
    )

    $candidates = @()

    $exact = Get-CMApplication -Name $AppName -Fast -ErrorAction SilentlyContinue
    if ($exact) { $candidates += $exact }

    if ($candidates.Count -eq 0) {
        $wild = Get-CMApplication -Name $AppNamePattern -Fast -ErrorAction SilentlyContinue
        if ($wild) { $candidates += @($wild) }
    }

    foreach ($app in $candidates) {
        if ($app -and $app.NumberOfRevisions -gt 1) {
            Remove-CMApplicationRevisionHistory -InputObject $app -Force
        }
    }
}

function Remove-CMApplicationRevisionHistoryByCIId {
    param(
        [Parameter(Mandatory)][int]$CI_ID,
        [int]$KeepLatest = 1
    )

    $history = Get-CMApplicationRevisionHistory -Id $CI_ID -ErrorAction SilentlyContinue
    if (-not $history) { return }

    $revList = @()
    foreach ($h in @($history)) {
        if ($h.PSObject.Properties.Name -contains 'Revision') {
            $revList += [uint32]$h.Revision
        }
        elseif ($h.PSObject.Properties.Name -contains 'CIVersion') {
            $revList += [uint32]$h.CIVersion
        }
    }

    $revList = $revList | Sort-Object -Unique -Descending
    if ($revList.Count -le $KeepLatest) { return }

    foreach ($rev in ($revList | Select-Object -Skip $KeepLatest)) {
        Remove-CMApplicationRevisionHistory -Id ([uint32]$CI_ID) -Revision ([uint32]$rev) -Force -ErrorAction Stop
    }
}

# --- Main Script Execution ---
$originalLocation = Get-Location

try {
    if ($GetLatestVersionOnly) {
        $release = Invoke-RestMethod -Uri $GitHubApiUrl -Headers @{ "User-Agent" = "PowerShell" } -ErrorAction Stop
        $NppVersion = $release.tag_name -replace '^v'
        Write-Output $NppVersion
        exit 0
    }

    # Log module version when present
    $cmModule = Get-Module ConfigurationManager -ErrorAction SilentlyContinue
    if ($cmModule) {
        Write-Host "Configuration Manager module version: $($cmModule.Version)"
    }

    if (-not (Test-IsAdmin)) {
        Write-Error "This script must be run with admin privileges. Please run PowerShell as Administrator."
        exit 1
    }

    Set-Location C: -ErrorAction Stop

    # 1. Get latest Notepad++ version
    Write-Host "Contacting GitHub for latest Notepad++ release info..." -ForegroundColor Cyan
    $release = Invoke-RestMethod -Uri $GitHubApiUrl -Headers @{ "User-Agent" = "PowerShell" } -ErrorAction Stop
    $NppVersion = $release.tag_name -replace '^v'
    Write-Host "Latest Notepad++ version: $NppVersion" -ForegroundColor Green

    # 2. Download and Stage Content
    if (-not (Test-Path -LiteralPath $LocalDlRoot)) {
        New-Item -Path $LocalDlRoot -ItemType Directory -Force | Out-Null
    }

    $Asset = $release.assets | Where-Object { $_.name -like "*Installer.x64.exe" } | Select-Object -First 1
    if (-not $Asset) {
        Write-Error "Could not find x64 installer asset in GitHub release."
        exit 1
    }

    $NppInstallerName = $Asset.name
    $NppInstallerUrl  = $Asset.browser_download_url
    $LocalInstaller   = Join-Path $LocalDlRoot $NppInstallerName

    if (-not (Test-Path -LiteralPath $LocalInstaller)) {
        Write-Host "Downloading $NppInstallerName..."
        curl.exe -L --fail --silent --show-error -A "PowerShell" -o $LocalInstaller $NppInstallerUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $NppInstallerUrl" }
        Write-Host "Downloaded: $LocalInstaller"
    }
    else {
        Write-Host "Installer already present: $LocalInstaller"
    }

    # 3. Copy to UNC share and create content scripts
    $NetworkRoot = Join-Path $FileServerPath "Applications"

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) { exit 1 }
    if (-not (Test-NetworkShareAccess -Path $NetworkRoot))   { exit 1 }

    Set-Location C: -ErrorAction Stop

    $AppRoot = Ensure-ContentRootFolders -FileServerPath $NetworkRoot -Vendor $VendorFolder -App $AppFolder

    $UncRoot = Join-Path $AppRoot $NppVersion
    if (-not (Test-Path -LiteralPath $UncRoot)) {
        Write-Host "Creating UNC content directory $UncRoot ..."
        New-Item -Path $UncRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    Write-Host "Copying installer to $UncRoot ..."
    Copy-Item -LiteralPath $LocalInstaller -Destination (Join-Path $UncRoot $NppInstallerName) -Force -ErrorAction Stop

    $InstallBat   = Join-Path $UncRoot "install.bat"
    $UninstallBat = Join-Path $UncRoot "uninstall.bat"

    $InstallBatContent = @"
@ECHO OFF
taskkill /f /im notepad++.exe
start /wait "" "%~dp0$NppInstallerName" /S /noUpdater
"@

    $UninstallBatContent = @"
@ECHO OFF
start /wait "" "%ProgramFiles%\Notepad++\uninstall.exe" /S
"@

    Set-Content -LiteralPath $InstallBat -Value $InstallBatContent -Encoding ASCII -ErrorAction Stop
    Set-Content -LiteralPath $UninstallBat -Value $UninstallBatContent -Encoding ASCII -ErrorAction Stop
    Write-Host "Created install.bat and uninstall.bat."

    # 4. Define Detection Rule
    $DetectionClause = New-CMDetectionClauseFile `
        -Path "%ProgramFiles%\Notepad++" `
        -FileName "notepad++.exe" `
        -Value `
        -PropertyType Version `
        -ExpressionOperator GreaterEquals `
        -ExpectedValue $NppVersion

    Write-Host "Defined file detection rule for version $NppVersion."

    # 5. Connect to Configuration Manager
    if (-not (Connect-CMSite -SiteCode $SiteCode)) {
        Write-Error "Cannot proceed without CM connection."
        exit 1
    }

    # 6. Application and deployment type names
    $AppName = "Notepad++ - $NppVersion"
    $DTName  = "Notepad++ - $NppVersion"

    # 7. Check for existing application
    $ExistingApp = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
    if ($ExistingApp) {
        Write-Warning "Application '$AppName' already exists with CI_ID: $($ExistingApp.CI_ID)."
        exit 1
    }

    # 8. Create new application
    Write-Host "Creating new application '$AppName' version $NppVersion." -ForegroundColor Yellow

    $App = New-CMApplication `
        -Name $AppName `
        -SoftwareVersion $NppVersion `
        -Publisher $Publisher `
        -Description $Comment `
        -AutoInstall $true `
        -ErrorAction Stop

    Write-Host "Application CI_ID: $($App.CI_ID)"
    $AppCIId = [int]$App.CI_ID
    Write-Host "Creating new deployment type '$DTName'."

    Add-CMScriptDeploymentType `
        -ApplicationName $AppName `
        -DeploymentTypeName $DTName `
        -InstallCommand "install.bat" `
        -UninstallCommand "uninstall.bat" `
        -ContentLocation $UncRoot `
        -AddDetectionClause @($DetectionClause) `
        -InstallationBehaviorType InstallForSystem `
        -UserInteractionMode Normal `
        -LogonRequirementType WhetherOrNotUserLoggedOn `
        -MaximumRuntimeMins 20 `
        -EstimatedRuntimeMins 10 `
        -ContentFallback `
        -SlowNetworkDeploymentMode Download `
        -ErrorAction Stop | Out-Null

    # Revision history cleanup (CI_ID-based)
    Remove-CMApplicationRevisionHistoryByCIId -CI_ID $AppCIId -KeepLatest 1

    Write-Host "Creation of Application and Deployment Type complete for Notepad++ version $NppVersion." -ForegroundColor Green
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
}
finally {
    Set-Location $originalLocation -ErrorAction SilentlyContinue
    Write-Host "Restored initial location to: ${originalLocation}"
}