<#
Vendor: Microsoft
App: Microsoft Edge WebView2 Runtime (x64)
CMName: Microsoft Edge WebView2 Runtime

.SYNOPSIS
    Packages Microsoft Edge WebView2 Runtime (x64) for MECM.
.DESCRIPTION
    Downloads the latest Microsoft Edge WebView2 Runtime x64 EXE, stages content to a versioned network folder,
    and creates an MECM Application with a Script Deployment Type.
    Detection: msedgewebview2.exe file version >= packaged version.
.PARAMETER SiteCode
    ConfigMgr site code for the CM PSDrive (e.g., "MCM"). The PSDrive is assumed to already exist.
.PARAMETER Comment
    Work order or comment string applied to the MECM application description.
.PARAMETER FileServerPath
    UNC root of the SCCM content share (e.g., "\\fileserver\sccm$").
.PARAMETER GetLatestVersionOnly
    Outputs only the latest version string and exits.
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
$WebView2NetworkShareRoot = Join-Path $FileServerPath "Applications\Microsoft\WebView2"
$WebView2ExeUrl = "https://go.microsoft.com/fwlink/?linkid=2124701"
$EdgeVersionUrl = "https://edgeupdates.microsoft.com/api/products?view=enterprise"
$Product = @{
    Name = "Microsoft Edge WebView2 Runtime - {0}"
    FileName = "MicrosoftEdgeWebView2RuntimeInstallerX64.exe"
    InstallBatContent = 'start /wait "" "%~dp0MicrosoftEdgeWebView2RuntimeInstallerX64.exe" /silent /install'
    UninstallBatContent = 'start /wait "" "%~dp0MicrosoftEdgeWebView2RuntimeInstallerX64.exe" /silent /uninstall'
    InstallPs1Content = 'Start-Process "$PSScriptRoot\MicrosoftEdgeWebView2RuntimeInstallerX64.exe" -ArgumentList "/silent /install" -Wait'
    UninstallPs1Content = 'Start-Process "$PSScriptRoot\MicrosoftEdgeWebView2RuntimeInstallerX64.exe" -ArgumentList "/silent /uninstall" -Wait'
    DetectionFile = "msedgewebview2.exe"
    DetectionPath = "$env:ProgramFiles\Microsoft\EdgeWebView\Application\{0}"
    Publisher = "Microsoft Corporation"
}


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

function Get-LatestEdgeVersion {
    param([switch]$Quiet)
    if (-not $Quiet) { Write-Host "Fetching Microsoft Edge version from: ${EdgeVersionUrl}" }
    try {
        $response = Invoke-RestMethod -Uri $EdgeVersionUrl -ErrorAction Stop
        $stableChannel = $response | Where-Object { $_.Product -eq "Stable" }
        $latestVersion = $stableChannel.Releases | Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" } | 
            Sort-Object { [version]$_.ProductVersion } -Descending | Select-Object -First 1 -ExpandProperty ProductVersion
        if (-not $latestVersion) {
            throw "Could not determine latest Microsoft Edge Stable version."
        }
    if (-not $Quiet) { Write-Host "Found latest Microsoft Edge Stable version: ${latestVersion}" }
        return $latestVersion
    }
    catch {
        if (-not $Quiet) { Write-Error "Failed to fetch Edge version: $($_.Exception.Message)" }
        return "140.0.3485.66"
    }
}

function Test-NetworkShareAccess {
    param ([string]$Path)
    $originalLocation = Get-Location
    Write-Host "Current location before network share validation: ${originalLocation}"
    try {
        Set-Location C: -ErrorAction Stop
        Write-Host "Set location to C: for network share validation"
        if (-not $Path) { Write-Error "Network path is null or empty."; return $false }
        if (-not (Test-Path $Path -ErrorAction Stop)) { Write-Error "Network path '${Path}' does not exist or is inaccessible."; return $false }
        $testFile = Join-Path $Path "test_$(Get-Random).txt"
        Set-Content -Path $testFile -Value "Test" -ErrorAction Stop
        Remove-Item -Path $testFile -ErrorAction Stop
        Write-Host "Network share '${Path}' is accessible and writable."
        return $true
    }
    catch {
        Write-Error "Failed to access network share '${Path}': $($_.Exception.Message)"
        return $false
    }
    finally {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
        Write-Host "Restored location to: ${originalLocation}"
    }
}

function Create-BatchAndPS1Files {
    param (
        [string]$NetworkPath,
        [string]$InstallBatContent,
        [string]$UninstallBatContent,
        [string]$InstallPs1Content,
        [string]$UninstallPs1Content
    )
    $originalLocation = Get-Location
    Write-Host "Current location before file creation: ${originalLocation}"
    try {
        Set-Location C: -ErrorAction Stop
        Write-Host "Set location to C: for file creation"
        if (-not (Test-Path $NetworkPath)) {
            Write-Host "Creating directory: ${NetworkPath}"
            New-Item -ItemType Directory -Path $NetworkPath -Force | Out-Null
        }
        $InstallBatPath = Join-Path $NetworkPath "install.bat"
        $UninstallBatPath = Join-Path $NetworkPath "uninstall.bat"
        Set-Content -Path $InstallBatPath -Value $InstallBatContent -Encoding ASCII -ErrorAction Stop
        Write-Host "Created install.bat at ${NetworkPath}"
        Set-Content -Path $UninstallBatPath -Value $UninstallBatContent -Encoding ASCII -ErrorAction Stop
        Write-Host "Created uninstall.bat at ${NetworkPath}"
        if ($InstallPs1Content -and $UninstallPs1Content) {
            $InstallPs1Path = Join-Path $NetworkPath "install.ps1"
            $UninstallPs1Path = Join-Path $NetworkPath "uninstall.ps1"
            Set-Content -Path $InstallPs1Path -Value $InstallPs1Content -Encoding ASCII -ErrorAction Stop
            Write-Host "Created install.ps1 at ${NetworkPath}"
            Set-Content -Path $UninstallPs1Path -Value $UninstallPs1Content -Encoding ASCII -ErrorAction Stop
            Write-Host "Created uninstall.ps1 at ${NetworkPath}"
        }
    }
    catch {
        Write-Error "Failed to create files in '${NetworkPath}': $($_.Exception.Message)"
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
        [string]$Publisher,
        [string]$DetectionPath,
        [string]$DetectionFile,
        [string]$DetectionStagedPath
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
        $detectionClauses = @()
        # Primary detection clause
        $detectionPathFormatted = if ($DetectionPath -like "*{0}*") { $DetectionPath -f $Version } else { $DetectionPath }
        Write-Host "Creating detection clause for ${DetectionFile} at ${detectionPathFormatted}"
        if (-not (Test-Path $detectionPathFormatted -ErrorAction SilentlyContinue)) {
            Write-Warning "Detection path '${detectionPathFormatted}' is not accessible. Ensure it exists on target systems."
        }
        $clause1 = New-CMDetectionClauseFile -Path $detectionPathFormatted `
            -FileName $DetectionFile `
            -Value `
            -PropertyType Version `
            -ExpressionOperator GreaterEquals `
            -ExpectedValue $Version `
            -ErrorAction Stop
        $detectionClauses += $clause1
        # Staged detection clause (for Edge only)
        if ($DetectionStagedPath) {
            Write-Host "Creating staged detection clause for ${DetectionFile} at ${DetectionStagedPath}"
            if (-not (Test-Path $DetectionStagedPath -ErrorAction SilentlyContinue)) {
                Write-Warning "Staged detection path '${DetectionStagedPath}' is not accessible. Ensure it exists on target systems."
            }
            $clause2 = New-CMDetectionClauseFile -Path $DetectionStagedPath `
                -FileName $DetectionFile `
                -Value `
                -PropertyType Version `
                -ExpressionOperator GreaterEquals `
                -ExpectedValue $Version `
                -ErrorAction Stop
            $detectionClauses += $clause2
        }
        $params = @{
            ApplicationName = $AppName
            DeploymentTypeName = "${AppName}"
            InstallCommand = "install.bat"
            ContentLocation = $NetworkPath
            ContentFallback = $true
            SlowNetworkDeploymentMode = "Download"
            UninstallCommand = "uninstall.bat"
            InstallationBehaviorType = "InstallForSystem"
            LogonRequirementType = "WhetherOrNotUserLoggedOn"
            MaximumRuntimeMins = 20
            EstimatedRuntimeMins = 10
            AddDetectionClause = $detectionClauses
            UserInteractionMode = "Hidden"
            RebootBehavior = "NoAction"
            ErrorAction = "Stop"
        }
        if ($DetectionStagedPath) {
            $params.Add("DetectionClauseConnector", @{"LogicalName"=$clause2.Setting.LogicalName; "Connector"="OR"})
        }
        Write-Host "Adding script deployment type with parameters: $($params | Out-String)"
        Add-CMScriptDeploymentType @params -Verbose
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$app.CI_ID) -KeepLatest 1
        Write-Host "Created MECM application: ${AppName} with file-based detection"
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
    $originalLocation = Get-Location
    Write-Host "Initial location: ${originalLocation}"
    Set-Location C: -ErrorAction Stop
    Write-Host "Set initial location to C: for script execution"

    if ($GetLatestVersionOnly) {
        $versionOnly = Get-LatestEdgeVersion -Quiet
        Write-Output $versionOnly
        return
    }

    $version = Get-LatestEdgeVersion
    $downloadFolderName = "WebView2Installers_${version}"
    $downloadPath = Join-Path $BaseDownloadRoot $downloadFolderName
    if (-not (Test-Path -LiteralPath $downloadPath)) {
        Write-Host "Creating download directory: ${downloadPath}"
        New-Item -ItemType Directory -Path $downloadPath -Force | Out-Null
    }

    $networkPath = Join-Path $WebView2NetworkShareRoot $version
    if (-not (Test-NetworkShareAccess -Path $WebView2NetworkShareRoot)) {
        Write-Error "Network share '${WebView2NetworkShareRoot}' is inaccessible. Skipping WebView2 application."
        exit 1
    }
    if (-not (Test-Path -LiteralPath $networkPath)) {
        Write-Host "Creating WebView2 network directory: ${networkPath}"
        New-Item -ItemType Directory -Path $networkPath -Force -ErrorAction Stop | Out-Null
    }

    $product = $Product
    $appName = $product.Name -f $version
    $fileName = $product.FileName
    $downloadUrl = $WebView2ExeUrl
    $outputPath = Join-Path $downloadPath $fileName
    $networkFilePath = Join-Path $networkPath $fileName

    if (Test-Path -LiteralPath $outputPath) {
        $fileInfo = Get-Item -LiteralPath $outputPath -ErrorAction Stop
        Write-Host "${fileName} already exists at ${outputPath} (Size=$($fileInfo.Length) bytes, LastModified=$($fileInfo.LastWriteTime)), skipping download."
    }
    else {
        Write-Host "Downloading ${appName} (${fileName}) from ${downloadUrl}"
        Set-Location C: -ErrorAction Stop
        curl.exe -L --fail --silent --show-error -o $outputPath $downloadUrl
        if ($LASTEXITCODE -ne 0) { throw "Download failed: $downloadUrl" }
        $fileInfo = Get-Item -LiteralPath $outputPath -ErrorAction Stop
        Write-Host "Downloaded ${appName} (${fileName}) to ${outputPath} (Size=$($fileInfo.Length) bytes)"
    }

    if (Test-Path -LiteralPath $networkFilePath) {
        $fileInfo = Get-Item -LiteralPath $networkFilePath -ErrorAction Stop
        Write-Host "${fileName} already exists at ${networkFilePath} (Size=$($fileInfo.Length) bytes, LastModified=$($fileInfo.LastWriteTime)), skipping copy."
    }
    else {
        Set-Location C: -ErrorAction Stop
        Write-Host "Copying ${fileName} to ${networkPath}"
        Copy-Item -LiteralPath $outputPath -Destination $networkPath -Force -ErrorAction Stop
        $fileInfo = Get-Item -LiteralPath $networkFilePath -ErrorAction Stop
        Write-Host "Copied ${fileName} to ${networkPath} (Size=$($fileInfo.Length) bytes)"
    }

    Create-BatchAndPS1Files -NetworkPath $networkPath `
        -InstallBatContent $product.InstallBatContent `
        -UninstallBatContent $product.UninstallBatContent `
        -InstallPs1Content $product.InstallPs1Content `
        -UninstallPs1Content $product.UninstallPs1Content

    $params = @{
        AppName = $appName
        Version = $version
        NetworkPath = $networkPath
        FileName = $fileName
        Publisher = $product.Publisher
        DetectionPath = ($product.DetectionPath -f $version)
        DetectionFile = $product.DetectionFile
    }
    New-MECMApplication @params

    Write-Host "Script execution complete."
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($originalLocation) {
        Set-Location $originalLocation -ErrorAction SilentlyContinue
        Write-Host "Restored initial location to: ${originalLocation}"
    }
}