<#
.SYNOPSIS
    WinForms front-end for application packager scripts (metadata-driven, no network on launch).

.DESCRIPTION
    Displays locally-discovered packager scripts in a DataGridView with checkboxes and status indicators.
    On launch, the tool performs LOCAL-ONLY operations:
      - Enumerates packager scripts in the PackagersRoot folder
      - Parses metadata tags from each script header:
          # Vendor:
          # App:
          # CMName:   (optional; defaults to App)
      - Populates the grid with Vendor/Application and placeholders for Current/Latest/Status

    No network operations are performed on launch. Network operations are only performed after explicit user action:
      - Check Latest: runs selected packagers with -GetLatestVersionOnly and updates Latest/Status
      - Check MECM: queries MECM for selected products (requires existing CM PSDrive session)
      - Run Selected: runs selected packagers normally and captures output to logs

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM"). The PSDrive is assumed to already exist in the session.
    Used only when MECM actions are invoked by the user.

.PARAMETER PackagersRoot
    Local folder containing packager scripts (e.g., .\Packagers).
    Scripts supported:
      - package-*.ps1
      - package-*.notps1   (treated as PowerShell script content; used for email/security systems)

.EXAMPLE
    .\start-apppackager.ps1

.EXAMPLE
    .\start-apppackager.ps1 -SiteCode "MCM" -PackagersRoot "D:\CM\Packagers"

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8.2
      - Windows Forms (System.Windows.Forms)

    Startup behavior:
      - No MECM queries on launch
      - No internet queries on launch
      - No network share access on launch

    ScriptName : start-apppackager.ps1
    Purpose    : WinForms front-end for packager scripts (metadata-driven selection + actions)
    Owner      : CM Engineering
    Version    : 0.1.0
    Updated    : 2026-01-26
#>


param(
    [string]$SiteCode = "MCM",
    [string]$PackagersRoot = (Join-Path $PSScriptRoot "Packagers")
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }
# -----------------------------
# Helpers
# -----------------------------
function Get-ConfigStorePath {
    Join-Path $PSScriptRoot "AppPackager.configurations.json"
}

function Read-ConfigStore {
    $path = Get-ConfigStorePath
    if (-not (Test-Path -LiteralPath $path)) { return @() }

    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        $data = $raw | ConvertFrom-Json -ErrorAction Stop

        if ($data -is [System.Collections.IEnumerable]) { return @($data) }
        return @($data)
    }
    catch {
        return @()
    }
}

function Write-ConfigStore {
    param([Parameter(Mandatory)][object[]]$Configs)

    $path = Get-ConfigStorePath
    $json = ($Configs | ConvertTo-Json -Depth 6)
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8
}

function Set-ConfigurationInputs {
    param(
        [Parameter(Mandatory)][pscustomobject]$Config,
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TxtComment,
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TxtFSPath,
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TxtSiteCode
    )

    if ($null -ne $Config.WOComment)     { $TxtComment.Text  = [string]$Config.WOComment }
    if ($null -ne $Config.FileShareRoot) { $TxtFSPath.Text   = [string]$Config.FileShareRoot }
    if ($null -ne $Config.SiteCode)      { $TxtSiteCode.Text = [string]$Config.SiteCode }
}

function Save-Configuration {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$WOComment,
        [Parameter(Mandatory)][string]$FileShareRoot,
        [Parameter(Mandatory)][string]$SiteCode
    )

    $configs = @(Read-ConfigStore)

    $existing = $configs | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
    if ($existing) {
        $existing.WOComment     = $WOComment
        $existing.FileShareRoot = $FileShareRoot
        $existing.SiteCode      = $SiteCode
    }
    else {
        $configs += [pscustomobject]@{
            Name          = $Name
            WOComment     = $WOComment
            FileShareRoot = $FileShareRoot
            SiteCode      = $SiteCode
        }
    }

    Write-ConfigStore -Configs $configs
    return $configs
}

function Select-OnlyUpdateAvailable {
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable
    )

    foreach ($row in $DataTable.Rows) {
        $row["Selected"] = ($row["Status"] -eq "Update available")
    }
}

function New-Mdl2IconBitmap {
    param(
        [Parameter(Mandatory)][int]$CodePoint,
        [int]$Size = 20
    )

    $bmp = New-Object System.Drawing.Bitmap $Size, $Size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $g.Clear([System.Drawing.Color]::Transparent)

    $font = New-Object System.Drawing.Font("Segoe MDL2 Assets", ($Size - 4), [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(60, 60, 60))

    $glyph = [char]$CodePoint
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center

    $rect = New-Object System.Drawing.RectangleF 0, 0, $Size, $Size
    $g.DrawString($glyph, $font, $brush, $rect, $sf)

    $brush.Dispose()
    $font.Dispose()
    $sf.Dispose()
    $g.Dispose()

    return $bmp
}
function Set-ModernButtonStyle {
    param(
        [Parameter(Mandatory=$true)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory=$true)][System.Drawing.Color]$BackColor
    )

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $BackColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand

    $hover = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 18),
        [Math]::Max(0, $BackColor.G - 18),
        [Math]::Max(0, $BackColor.B - 18)
    )
    $down = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 36),
        [Math]::Max(0, $BackColor.G - 36),
        [Math]::Max(0, $BackColor.B - 36)
    )

    $Button.FlatAppearance.MouseOverBackColor = $hover
    $Button.FlatAppearance.MouseDownBackColor = $down
}


function Enable-DoubleBuffer {
    param([Parameter(Mandatory=$true)][System.Windows.Forms.Control]$Control)

    $prop = $Control.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
    if ($prop) { $prop.SetValue($Control, $true, $null) | Out-Null }
}

function Add-LogLine {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TextBox,
        [Parameter(Mandatory)][string]$Message
    )
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "{0}  {1}" -f $ts, $Message

    if ([string]::IsNullOrWhiteSpace($TextBox.Text)) {
        $TextBox.Text = $line
    }
    else {
        $TextBox.AppendText([Environment]::NewLine + $line)
    }

    $TextBox.SelectionStart = $TextBox.TextLength
    $TextBox.ScrollToCaret()
}

function Get-PackagerMetadata {
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $meta = [ordered]@{
        Vendor = $null
        App    = $null
        CMName = $null
    }

    $lines = Get-Content -LiteralPath $Path -TotalCount 200 -ErrorAction Stop

    foreach ($line in $lines) {
        $l = $line.TrimStart([char]0xFEFF)

        if (-not $meta.Vendor -and $l -match '^\s*(?:#\s*)?Vendor\s*:\s*(.+?)\s*$') { $meta.Vendor = $Matches[1].Trim(); continue }
        if (-not $meta.App    -and $l -match '^\s*(?:#\s*)?App\s*:\s*(.+?)\s*$')    { $meta.App    = $Matches[1].Trim(); continue }
        if (-not $meta.CMName -and $l -match '^\s*(?:#\s*)?CMName\s*:\s*(.+?)\s*$') { $meta.CMName = $Matches[1].Trim(); continue }

        if (-not $meta.App    -and $l -match '^\s*(?:#\s*)?Application\s*:\s*(.+?)\s*$') { $meta.App = $Matches[1].Trim(); continue }
    }

    if (-not $meta.CMName) { $meta.CMName = $meta.App }

    return [pscustomobject]@{
        Vendor      = $meta.Vendor
        Application = $meta.App
        CMName      = $meta.CMName
        Script      = (Split-Path -Leaf $Path)
        FullPath    = $Path
    }
}

function Get-Packagers {
    param(
        [Parameter(Mandatory)][string]$Root
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        return @()
    }

    $files = Get-ChildItem -LiteralPath $Root -File -ErrorAction Stop |
        Where-Object { $_.Name -match '^package-.*\.(?:ps1|notps1)$' } |
        Sort-Object Name

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($f in $files) {
        try {
            $m = Get-PackagerMetadata -Path $f.FullName

            $status = "Ready"
            if ($f.Extension -ieq ".notps1") {
                $status = "Not runnable (.notps1)"
            }
            if (-not $m.Vendor -or -not $m.Application) {
                $status = "Missing metadata (Vendor/App)"
            }

            $items.Add([pscustomobject]@{
                Selected      = $false
                Vendor        = $m.Vendor
                Application   = $m.Application
                CMName        = $m.CMName
                Script        = $m.Script
                FullPath      = $m.FullPath
                CurrentVersion= ""
                LatestVersion = ""
                Status        = $status
            })
        }
        catch {
            $items.Add([pscustomobject]@{
                Selected      = $false
                Vendor        = ""
                Application   = ""
                CMName        = ""
                Script        = $f.Name
                FullPath      = $f.FullName
                CurrentVersion= ""
                LatestVersion = ""
                Status        = ("Read error: " + $_.Exception.Message)
            })
        }
    }
    return $items
}

function Test-PackagerSupportsFileServerPath {
    param([Parameter(Mandatory)][string]$PackagerPath)

    try {
        $head = Get-Content -LiteralPath $PackagerPath -TotalCount 120 -ErrorAction Stop | Out-String
        return ($head -match '\$FileServerPath')
    }
    catch {
        return $false
    }
}

function Invoke-PackagerGetLatestVersion {
    param(
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$SiteCode,
        [string]$FileServerPath = $null
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $argsBase = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -SiteCode "{1}" -GetLatestVersionOnly' -f $PackagerPath, $SiteCode)
    if ($FileServerPath -and (Test-PackagerSupportsFileServerPath -PackagerPath $PackagerPath)) {
        $argsBase = ($argsBase + (' -FileServerPath "{0}"' -f $FileServerPath))
    }
    $psi.Arguments = $argsBase
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi

    $null = $p.Start()
    $stderr = $p.StandardError.ReadToEnd()
    $stdout = $p.StandardOutput.ReadToEnd()
    $p.WaitForExit()

    if ($p.ExitCode -ne 0) {
        $msg = $stderr
        if ([string]::IsNullOrWhiteSpace($msg)) { $msg = $stdout }
        if ([string]::IsNullOrWhiteSpace($msg)) { $msg = "Packager returned exit code $($p.ExitCode)." }
        throw $msg.Trim()
    }

    $lines = @($stdout -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (-not $lines -or $lines.Count -lt 1) {
        throw "No version output received."
    }

    $version = ([string]$lines[0]).Trim()

    if ($version -notmatch '^\d+(\.\d+){1,3}$') {
        throw ("Unexpected version string: '{0}'" -f $version)
    }

    return $version
}

function Compare-SemVer {
    param(
        [Parameter(Mandatory)][string]$A,
        [Parameter(Mandatory)][string]$B
    )
    try {
        $va = [version]$A
        $vb = [version]$B
        return $va.CompareTo($vb)
    }
    catch {
        return 0
    }
}

function Get-MecmCurrentVersionByCMName {
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$CMName
    )

    if (-not (Get-Command -Name Get-CMApplication -ErrorAction SilentlyContinue)) {
        try {
            if ($env:SMS_ADMIN_UI_PATH) {
                $cmModule = Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH) "ConfigurationManager.psd1"
                if (Test-Path -LiteralPath $cmModule) {
                    Import-Module $cmModule -Force -ErrorAction Stop
                }
            }
        } catch { }
    }
    if (-not (Get-Command -Name Get-CMApplication -ErrorAction SilentlyContinue)) {
        throw "ConfigMgr PowerShell cmdlets not available in this session."
    }

    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
    }
    catch {
        throw ("Failed to connect to CM site PSDrive '{0}:'" -f $SiteCode)
    }

    $apps = @(Get-CMApplication -Name $CMName -ErrorAction SilentlyContinue)

    if (-not $apps -or $apps.Count -eq 0) {
        $apps = @(Get-CMApplication -Name ("{0}*" -f $CMName) -ErrorAction SilentlyContinue)
    }

    if (-not $apps -or $apps.Count -eq 0) {
        return [pscustomobject]@{
            Found          = $false
            DisplayName    = $null
            SoftwareVersion= $null
            MatchCount     = 0
        }
    }

    $exact = $apps | Where-Object { $_.LocalizedDisplayName -eq $CMName -or $_.Name -eq $CMName }
    if ($exact -and $exact.Count -gt 0) {
        $chosen = $exact | Select-Object -First 1
    }
    else {
        $parsable = @()
        $nonParsable = @()

        foreach ($a in $apps) {
            try {
                $null = [version]$a.SoftwareVersion
                $parsable += $a
            }
            catch {
                $nonParsable += $a
            }
        }

        if ($parsable.Count -gt 0) {
            $chosen = $parsable | Sort-Object { [version]$_.SoftwareVersion } -Descending | Select-Object -First 1
        }
        else {
            $chosen = $nonParsable | Sort-Object Name -Descending | Select-Object -First 1
        }
    }

    return [pscustomobject]@{
        Found           = $true
        DisplayName     = $chosen.LocalizedDisplayName
        SoftwareVersion = $chosen.SoftwareVersion
        MatchCount      = $apps.Count
    }
}

function Invoke-PackagerRun {
    param(
        [Parameter(Mandatory)][string]$PackagerPath,
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$Comment,
        [Parameter(Mandatory)][string]$FileServerPath,
        [Parameter(Mandatory)][string]$LogFolder
    )

    if (-not (Test-Path -LiteralPath $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
    }

    $stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    $base  = [IO.Path]::GetFileNameWithoutExtension($PackagerPath)
    $outLog         = Join-Path $LogFolder ("{0}-{1}.out.log" -f $base, $stamp)
    $errLog         = Join-Path $LogFolder ("{0}-{1}.err.log" -f $base, $stamp)
    $structuredLog  = Join-Path $LogFolder ("{0}-{1}.structured.log" -f $base, $stamp)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $argsBase = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -SiteCode "{1}" -Comment "{2}" -LogPath "{3}"' -f $PackagerPath, $SiteCode, $Comment, $structuredLog)
    if (Test-PackagerSupportsFileServerPath -PackagerPath $PackagerPath) {
        $argsBase = ($argsBase + (' -FileServerPath "{0}"' -f $FileServerPath))
    }
    $psi.Arguments = $argsBase
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi

    $null = $p.Start()
    $stderr = $p.StandardError.ReadToEnd()
    $stdout = $p.StandardOutput.ReadToEnd()
    $p.WaitForExit()

    Set-Content -LiteralPath $outLog -Value $stdout -Encoding UTF8
    Set-Content -LiteralPath $errLog -Value $stderr -Encoding UTF8

    return [pscustomobject]@{
        ExitCode      = $p.ExitCode
        OutLog        = $outLog
        ErrLog        = $errLog
        StructuredLog = $structuredLog
        StdErr        = $stderr
    }
}

# -----------------------------
# UI
# -----------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "AppPackager"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1120, 720)
$form.MinimumSize = New-Object System.Drawing.Size(980, 640)

$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.BackColor = [System.Drawing.Color]::White
$script:AppIconBitmap = $null
$script:AppIconHandle = [IntPtr]::Zero

$logoPath = Join-Path $PSScriptRoot "apppackager-logo.jpg"
if (-not (Test-Path -LiteralPath $logoPath)) {
    $logoPath = Join-Path $PSScriptRoot "apppackager-logo.png"
}

try {
    if (Test-Path -LiteralPath $logoPath) {
        $script:AppIconBitmap = New-Object System.Drawing.Bitmap $logoPath
        $script:AppIconHandle = $script:AppIconBitmap.GetHicon()
        $form.Icon = [System.Drawing.Icon]::FromHandle($script:AppIconHandle)
    }
    else {
        $form.Icon = [System.Drawing.SystemIcons]::Application
    }
}
catch {
    $form.Icon = [System.Drawing.SystemIcons]::Application
}

# Header logo
$picLogo = New-Object System.Windows.Forms.PictureBox
$picLogo.Size = New-Object System.Drawing.Size(196, 196)
$picLogo.Location = New-Object System.Drawing.Point(20, 20)
$picLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
$picLogo.BackColor = [System.Drawing.Color]::Transparent

if (Test-Path -LiteralPath $logoPath) {
    try { $picLogo.Image = [System.Drawing.Image]::FromFile($logoPath) } catch { }
}

$form.Controls.Add($picLogo)

# -----------------------------
# In-app descriptor text block (compact + wrapped)
# -----------------------------
$descPanel = New-Object System.Windows.Forms.Panel
$descPanel.BackColor   = [System.Drawing.Color]::FromArgb(247, 249, 252)
$descPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$descPanel.Location    = New-Object System.Drawing.Point(($picLogo.Right + 20), ($picLogo.Top + 20))
$descPanel.Size        = New-Object System.Drawing.Size(10, 112)  # width is set in Set-UILayout()
$descPanel.Anchor      = "Top,Left,Right"
$form.Controls.Add($descPanel)

$descFont = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Regular)
$descColor = [System.Drawing.Color]::FromArgb(32,32,32)

# Internal padding for the panel
$padX = 12
$padY = 8

# Line 1 (wrapped)
$lblDesc1 = New-Object System.Windows.Forms.Label
$lblDesc1.AutoSize  = $true
$lblDesc1.Font      = $descFont
$lblDesc1.ForeColor = $descColor
$lblDesc1.Location  = New-Object System.Drawing.Point($padX, $padY)
$lblDesc1.MaximumSize = New-Object System.Drawing.Size(($descPanel.Width - ($padX * 2)), 0)
$lblDesc1.Anchor    = "Top,Left,Right"
$lblDesc1.Text      = "AppPackager automates version discovery and MECM application updates using approved packager scripts."
$descPanel.Controls.Add($lblDesc1)

# Line 2 (muted)
$lblDesc2 = New-Object System.Windows.Forms.Label
$lblDesc2.AutoSize  = $true
$lblDesc2.Font      = $descFont
$lblDesc2.ForeColor = $descColor
$lblDesc2.Location  = New-Object System.Drawing.Point($padX, ($padY + 22))
$lblDesc2.MaximumSize = New-Object System.Drawing.Size(($descPanel.Width - ($padX * 2)), 0)
$lblDesc2.Anchor    = "Top,Left,Right"
$lblDesc2.Text      = "No network or MECM actions occur until a button is explicitly selected."
$descPanel.Controls.Add($lblDesc2)

# Lines 3-5 (clean, no pipes)
$lblDesc3 = New-Object System.Windows.Forms.Label
$lblDesc3.AutoSize  = $true
$lblDesc3.Font      = $descFont
$lblDesc3.ForeColor = $descColor
$lblDesc3.Location  = New-Object System.Drawing.Point($padX, ($padY + 44))
$lblDesc3.MaximumSize = New-Object System.Drawing.Size(($descPanel.Width - ($padX * 2)), 0)
$lblDesc3.Anchor    = "Top,Left,Right"
$lblDesc3.Text      = "Check Latest - queries vendor sources"
$descPanel.Controls.Add($lblDesc3)

$lblDesc4 = New-Object System.Windows.Forms.Label
$lblDesc4.AutoSize  = $true
$lblDesc4.Font      = $descFont
$lblDesc4.ForeColor = $descColor
$lblDesc4.Location  = New-Object System.Drawing.Point($padX, ($padY + 64))
$lblDesc4.MaximumSize = New-Object System.Drawing.Size(($descPanel.Width - ($padX * 2)), 0)
$lblDesc4.Anchor    = "Top,Left,Right"
$lblDesc4.Text      = "Check MECM - reads current versions from the ConfigMgr site"
$descPanel.Controls.Add($lblDesc4)

$lblDesc5 = New-Object System.Windows.Forms.Label
$lblDesc5.AutoSize  = $true
$lblDesc5.Font      = $descFont
$lblDesc5.ForeColor = $descColor
$lblDesc5.Location  = New-Object System.Drawing.Point($padX, ($padY + 84))
$lblDesc5.MaximumSize = New-Object System.Drawing.Size(($descPanel.Width - ($padX * 2)), 0)
$lblDesc5.Anchor    = "Top,Left,Right"
$lblDesc5.Text      = "Debug columns expose internal identifiers (CMName, script source) for troubleshooting only. Internal tool for controlled administrative use."
$descPanel.Controls.Add($lblDesc5)

# File share root
$lblFSPath = New-Object System.Windows.Forms.Label
$lblFSPath.Text = "File Share Root:"
$lblFSPath.AutoSize = $true
$lblFSPath.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblFSPath.BackColor = $form.BackColor
$lblFSPath.Anchor = "Top,Right"
$form.Controls.Add($lblFSPath)

$txtFSPath = New-Object System.Windows.Forms.TextBox
$txtFSPath.Text = "\\fileserver\sccm$"
$txtFSPath.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtFSPath.Width = 200
$txtFSPath.MaxLength = 200
$txtFSPath.Anchor = "Top,Right"
$form.Controls.Add($txtFSPath)

# Site code
$lblSiteCode = New-Object System.Windows.Forms.Label
$lblSiteCode.Text = "Site Code:"
$lblSiteCode.AutoSize = $true
$lblSiteCode.BackColor = $form.BackColor
$lblSiteCode.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSiteCode.Anchor = "Top,Right"
$form.Controls.Add($lblSiteCode)

$txtSiteCode = New-Object System.Windows.Forms.TextBox
$txtSiteCode.Text = $SiteCode
$txtSiteCode.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtSiteCode.Width = 80
$txtSiteCode.MaxLength = 5
$txtSiteCode.Anchor = "Top,Right"
$form.Controls.Add($txtSiteCode)

# Work Order / Comment
$lblComment = New-Object System.Windows.Forms.Label
$lblComment.Text = "WO / Comment:"
$lblComment.AutoSize = $true
$lblComment.BackColor = $form.BackColor
$lblComment.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblComment.Anchor = "Top,Right"
$form.Controls.Add($lblComment)

$txtComment = New-Object System.Windows.Forms.TextBox
$txtComment.Text = "WO#00000001234567"
$txtComment.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtComment.Width = 220
$txtComment.MaxLength = 64
$txtComment.Anchor = "Top,Right"
$form.Controls.Add($txtComment)

# -----------------------------
# Configuration row (between descriptor and WO/Comment row)
# -----------------------------
$lblLoadConfig = New-Object System.Windows.Forms.Label
$lblLoadConfig.Text = "Load Configuration:"
$lblLoadConfig.AutoSize = $true
$lblLoadConfig.BackColor = $form.BackColor
$lblLoadConfig.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblLoadConfig.Anchor = "Top,Right"
$form.Controls.Add($lblLoadConfig)

$cmbLoadConfig = New-Object System.Windows.Forms.ComboBox
$cmbLoadConfig.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbLoadConfig.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$cmbLoadConfig.Width = 220
$cmbLoadConfig.Anchor = "Top,Right"
$form.Controls.Add($cmbLoadConfig)

$lblSaveConfig = New-Object System.Windows.Forms.Label
$lblSaveConfig.Text = "Save Configuration:"
$lblSaveConfig.AutoSize = $true
$lblSaveConfig.BackColor = $form.BackColor
$lblSaveConfig.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSaveConfig.Anchor = "Top,Right"
$form.Controls.Add($lblSaveConfig)

$txtSaveConfigName = New-Object System.Windows.Forms.TextBox
$txtSaveConfigName.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$txtSaveConfigName.Width = 200
$txtSaveConfigName.MaxLength = 64
$txtSaveConfigName.Anchor = "Top,Right"
$txtSaveConfigName.Text = "Configuration Name"
$form.Controls.Add($txtSaveConfigName)

$btnSaveConfig = New-Object System.Windows.Forms.Button
$btnSaveConfig.Width = 44
$btnSaveConfig.Height = 26
$btnSaveConfig.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSaveConfig.FlatAppearance.BorderSize = 0
$btnSaveConfig.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34) # green
$btnSaveConfig.ForeColor = [System.Drawing.Color]::White
$btnSaveConfig.UseVisualStyleBackColor = $false
$btnSaveConfig.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnSaveConfig.Text = [char]::ConvertFromUtf32(0x2713)  # checkmark
$btnSaveConfig.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$btnSaveConfig.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$btnSaveConfig.Anchor = "Top,Right"
$form.Controls.Add($btnSaveConfig)

# Keep a script-scoped cache for the combo binding
$script:ConfigCache = @()

function Update-ConfigDropdown {
    param([string]$SelectName = $null)

    $script:ConfigCache = @(Read-ConfigStore | Sort-Object Name)

    $cmbLoadConfig.BeginUpdate()
    $cmbLoadConfig.Items.Clear()
    foreach ($c in $script:ConfigCache) {
        [void]$cmbLoadConfig.Items.Add($c.Name)
    }
    $cmbLoadConfig.EndUpdate()

    if ($SelectName) {
        $idx = $cmbLoadConfig.Items.IndexOf($SelectName)
        if ($idx -ge 0) { $cmbLoadConfig.SelectedIndex = $idx }
    }
    elseif ($cmbLoadConfig.Items.Count -gt 0 -and $cmbLoadConfig.SelectedIndex -lt 0) {
        $cmbLoadConfig.SelectedIndex = 0
    }
}

$cmbLoadConfig.Add_SelectedIndexChanged({
    $name = [string]$cmbLoadConfig.SelectedItem
    if ([string]::IsNullOrWhiteSpace($name)) { return }

    $cfg = $script:ConfigCache | Where-Object { $_.Name -eq $name } | Select-Object -First 1
    if ($cfg) {
        Set-ConfigurationInputs -Config $cfg -TxtComment $txtComment -TxtFSPath $txtFSPath -TxtSiteCode $txtSiteCode
    }
})

$btnSaveConfig.Add_Click({
    $name = $txtSaveConfigName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($name) -or $name -eq "Configuration Name") { return }

    $null = Save-Configuration -Name $name -WOComment $txtComment.Text -FileShareRoot $txtFSPath.Text -SiteCode $txtSiteCode.Text
    Update-ConfigDropdown -SelectName $name
})

# Initial population
Update-ConfigDropdown

# Data grid
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 70)
$grid.Size = New-Object System.Drawing.Size(1060, 430)
$grid.Anchor = "Top,Left,Right,Bottom"
$grid.ReadOnly = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $false
$grid.AutoGenerateColumns = $false
$grid.ScrollBars = "Both"
$grid.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
$grid.AutoSizeRowsMode = "None"
$grid.ColumnHeadersHeightSizeMode = "DisableResizing"
$grid.ColumnHeadersHeight = 34
$grid.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$grid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$grid.GridColor = [System.Drawing.Color]::FromArgb(220,220,220)
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::White
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
$grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $grid.ColumnHeadersDefaultCellStyle.BackColor
$grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = $grid.ColumnHeadersDefaultCellStyle.ForeColor
$grid.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
$grid.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
$grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(225, 235, 245)
$grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(32,32,32)
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$grid.RowTemplate.Height = 32
$grid.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::None

$colSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSel.HeaderText = ""
$colSel.Width = 44
$colSel.Frozen = $true
$colSel.Name = "Selected"
$grid.Columns.Add($colSel) | Out-Null

$colVendor = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colVendor.HeaderText = "Vendor"
$colVendor.Width = 170
$colVendor.ReadOnly = $true
$colVendor.Name = "Vendor"
$grid.Columns.Add($colVendor) | Out-Null

$colApp = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colApp.HeaderText = "Application"
$colApp.Width = 360
$colApp.ReadOnly = $true
$colApp.Name = "Application"
$grid.Columns.Add($colApp) | Out-Null

$colCurrent = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCurrent.HeaderText = "Current (MECM)"
$colCurrent.Width = 160
$colCurrent.ReadOnly = $true
$colCurrent.Name = "CurrentVersion"
$grid.Columns.Add($colCurrent) | Out-Null

$colLatest = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colLatest.HeaderText = "Latest"
$colLatest.Width = 140
$colLatest.ReadOnly = $true
$colLatest.Name = "LatestVersion"
$grid.Columns.Add($colLatest) | Out-Null

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.HeaderText = "Status"
$colStatus.Width = 180
$colStatus.ReadOnly = $true
$colStatus.Name = "Status"
$grid.Columns.Add($colStatus) | Out-Null

$colCMName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCMName.HeaderText = "CMName (Debug)"
$colCMName.Width = 240
$colCMName.ReadOnly = $true
$colCMName.Name = "CMName"
$colCMName.Visible = $false
$grid.Columns.Add($colCMName) | Out-Null

$colScript = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colScript.HeaderText = "Script (Debug)"
$colScript.Width = 220
$colScript.ReadOnly = $true
$colScript.Name = "Script"
$colScript.Visible = $false
$grid.Columns.Add($colScript) | Out-Null

$form.Controls.Add($grid)

Enable-DoubleBuffer -Control $form
Enable-DoubleBuffer -Control $grid

# -----------------------------
# Select-All header checkbox
# -----------------------------
$chkSelectAll = New-Object System.Windows.Forms.CheckBox
$chkSelectAll.Size = New-Object System.Drawing.Size(16,16)
$chkSelectAll.BackColor = [System.Drawing.Color]::Transparent
$chkSelectAll.Checked = $false
$form.Controls.Add($chkSelectAll)
$chkSelectAll.BringToFront()

function Set-SelectAllCheckboxPosition {
    try {
        $headerRect = $grid.GetCellDisplayRectangle(0, -1, $true)

        $x = $grid.Left + $headerRect.X + [int](($headerRect.Width - $chkSelectAll.Width) / 2)
        $y = $grid.Top  + $headerRect.Y + [int](($headerRect.Height - $chkSelectAll.Height) / 2)

        $chkSelectAll.Location = New-Object System.Drawing.Point($x, $y)
        $chkSelectAll.Visible = $true
    }
    catch {
        $chkSelectAll.Visible = $false
    }
}

$grid.Add_Scroll({ Set-SelectAllCheckboxPosition })
$grid.Add_ColumnWidthChanged({ Set-SelectAllCheckboxPosition })
$grid.Add_SizeChanged({ Set-SelectAllCheckboxPosition })
$form.Add_Shown({ Set-SelectAllCheckboxPosition })

$chkSelectAll.Add_CheckedChanged({

    if ($script:SuppressSelectAll) {
        return
    }

    $grid.EndEdit()

    $target = $chkSelectAll.Checked
    foreach ($row in $dt.Rows) {
        $row["Selected"] = $target
    }
})

$grid.Add_CellValueChanged({
    param($s, $e)

    if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
        $s.EndEdit()

        $allSelected = $true
        $anySelected = $false

        foreach ($row in $dt.Rows) {
            if ($row["Selected"] -eq $true) {
                $anySelected = $true
            }
            else {
                $allSelected = $false
            }
        }

        $script:SuppressSelectAll = $true
        try {
            $chkSelectAll.Checked = ($anySelected -and $allSelected)
        }
        finally {
            $script:SuppressSelectAll = $false
        }
    }
})

$grid.Add_CurrentCellDirtyStateChanged({
    if ($grid.IsCurrentCellDirty) {
        $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

# Debug toggle
$chkDebug = New-Object System.Windows.Forms.CheckBox
$chkDebug.Text = "Show Debug Columns"
$chkDebug.AutoSize = $true
$chkDebug.Location = New-Object System.Drawing.Point(20, 510)
$chkDebug.Anchor = "Top,Left"
$form.Controls.Add($chkDebug)

# Unicode emoji code points
$emojiSearch = [char]::ConvertFromUtf32(0x1F50D)  # search
$emojiBox    = [char]::ConvertFromUtf32(0x1F5C4)  # file cabinet
$emojiPlay   = [char]::ConvertFromUtf32(0x25B6)   # play

# Buttons row
$btnLatest = New-Object System.Windows.Forms.Button
$btnLatest.Text = "$emojiSearch  Check Latest"
$btnLatest.Size = New-Object System.Drawing.Size(320, 52)
$btnLatest.Location = New-Object System.Drawing.Point(20, 548)
$btnLatest.Anchor = "Bottom,Left"
$btnLatest.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$btnLatest.ImageAlign = "MiddleLeft"
$btnLatest.TextAlign = "MiddleCenter"
$btnLatest.Padding = 0
$form.Controls.Add($btnLatest)

$btnMecm = New-Object System.Windows.Forms.Button
$btnMecm.Text   = "$emojiBox  Check MECM"
$btnMecm.Size = New-Object System.Drawing.Size(320, 52)
$btnMecm.Location = New-Object System.Drawing.Point(360, 548)
$btnMecm.Anchor = "Bottom,Left"
$btnMecm.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$btnMecm.ImageAlign = "MiddleLeft"
$btnMecm.TextAlign = "MiddleCenter"
$btnMecm.Padding = 0
$form.Controls.Add($btnMecm)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text    = "$emojiPlay  Run Selected Packagers"
$btnRun.Size = New-Object System.Drawing.Size(400, 52)
$btnRun.Location = New-Object System.Drawing.Point(700, 548)
$btnRun.Anchor = "Bottom,Right"
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$btnRun.ImageAlign = "MiddleLeft"
$btnRun.TextAlign = "MiddleCenter"
$btnRun.Padding = 0
$form.Controls.Add($btnRun)

Set-ModernButtonStyle -Button $btnLatest -BackColor ([System.Drawing.Color]::FromArgb(0, 120, 212))
Set-ModernButtonStyle -Button $btnMecm   -BackColor ([System.Drawing.Color]::FromArgb(16, 124, 16))
Set-ModernButtonStyle -Button $btnRun    -BackColor ([System.Drawing.Color]::FromArgb(217, 95, 2))

# 3-line log output
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.ReadOnly = $true
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.WordWrap = $false
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 10)
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtLog.Anchor = "Bottom,Left,Right"
$txtLog.Location = New-Object System.Drawing.Point(20, 612)
$txtLog.Size = New-Object System.Drawing.Size(1060, 54)
$form.Controls.Add($txtLog)

# Status strip
$status = New-Object System.Windows.Forms.StatusStrip
$status.Dock = [System.Windows.Forms.DockStyle]::Bottom
$status.SizingGrip = $false
$status.BackColor = [System.Drawing.Color]::White
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready. No network actions are performed until you click a button."
$status.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($status)

# -----------------------------
# Layout
# -----------------------------
function Set-UILayout {

    $padding   = 20
    $gap       = 20
    $btnHeight = 52
    $logHeight = 54

    # Descriptor panel: left aligns after logo; right aligns with form padding
    $descLeft = ($picLogo.Right + 16)
    $descWidth = ($form.ClientSize.Width - $padding - $descLeft)
    if ($descWidth -lt 200) { $descWidth = 200 }

    $descPanel.SetBounds($descLeft, ($picLogo.Top + 20), $descWidth, 112)

    $wrapW = [Math]::Max(200, ($descPanel.Width - (2 * $padX)))
    $lblDesc1.MaximumSize = New-Object System.Drawing.Size($wrapW, 0)
    $lblDesc2.MaximumSize = New-Object System.Drawing.Size($wrapW, 0)
    $lblDesc3.MaximumSize = New-Object System.Drawing.Size($wrapW, 0)
    $lblDesc4.MaximumSize = New-Object System.Drawing.Size($wrapW, 0)
    $lblDesc5.MaximumSize = New-Object System.Drawing.Size($wrapW, 0)

    $siteBoxWidth    = 80
    $fsBoxWidth      = 200
    $commentBoxWidth = 220
    $headerGap       = 12
    $topY = ($picLogo.Top + 172)
    $configY = $topY - 34
    if ($configY -lt 28) { $configY = 28 }  # safety for small window sizes

    # Config row: right-aligned, fits between descriptor and WO row
    $saveBtnW = 44
    $saveNameW = 200
    $loadW = 220
    $cfgGap = 12

    # Right edge padding
    $right = ($form.ClientSize.Width - $padding)

    # Save button at far right
    $btnSaveConfig.SetBounds(($right - $saveBtnW), $configY, $saveBtnW, 26)

    # Save name textbox to the left of save button
    $txtSaveConfigName.SetBounds(($btnSaveConfig.Left - $cfgGap - $saveNameW), $configY, $saveNameW, 24)

    # Save label to the left of save name
    $lblSaveConfig.Location = New-Object System.Drawing.Point(
        ($txtSaveConfigName.Left - 8 - $lblSaveConfig.PreferredWidth),
        ($configY + 3)
    )

    # Load combo to the left of save label area (leave a gap)
    $cmbLoadConfig.SetBounds(($lblSaveConfig.Left - $gap - $loadW), $configY, $loadW, 24)

    # Load label to the left of load combo
    $lblLoadConfig.Location = New-Object System.Drawing.Point(
        ($cmbLoadConfig.Left - 8 - $lblLoadConfig.PreferredWidth),
        ($configY + 3)
    )

    $txtSiteCode.SetBounds(
        ($form.ClientSize.Width - $padding - $siteBoxWidth),
        $topY,
        $siteBoxWidth,
        24
    )

    $lblSiteCode.Location = New-Object System.Drawing.Point(
        ($txtSiteCode.Left - 8 - $lblSiteCode.PreferredWidth),
        ($topY + 3)
    )

    $fsRight = ($lblSiteCode.Left - $headerGap)
    $minLeft = ($picLogo.Right + 20)

    $minFsBoxWidth = 180
    for ($i = 0; $i -lt 5; $i++) {
        $fsLeft = ($fsRight - $fsBoxWidth)

        $fsLabelLeft = ($fsLeft - 8 - $lblFSPath.PreferredWidth)
        if ($fsLabelLeft -ge $minLeft) { break }

        $delta = ($minLeft - $fsLabelLeft)
        $fsBoxWidth = ($fsBoxWidth - $delta)
        if ($fsBoxWidth -lt $minFsBoxWidth) { $fsBoxWidth = $minFsBoxWidth; break }
    }

    $fsLeft = ($fsRight - $fsBoxWidth)
    $txtFSPath.SetBounds($fsLeft, $topY, $fsBoxWidth, 24)
    $lblFSPath.Location = New-Object System.Drawing.Point(
        ($txtFSPath.Left - 8 - $lblFSPath.PreferredWidth),
        ($topY + 3)
    )

    $commentRight = ($lblFSPath.Left - $headerGap)
    $commentLeft  = ($commentRight - $commentBoxWidth)

    $minCommentBoxWidth = 160
    if ($commentLeft -lt $minLeft) {
        $delta = ($minLeft - $commentLeft)
        $commentBoxWidth = ($commentBoxWidth - $delta)
        if ($commentBoxWidth -lt $minCommentBoxWidth) { $commentBoxWidth = $minCommentBoxWidth }
        $commentLeft = ($commentRight - $commentBoxWidth)
    }

    $txtComment.SetBounds($commentLeft, $topY, $commentBoxWidth, 24)
    $lblComment.Location = New-Object System.Drawing.Point(
        ($txtComment.Left - 8 - $lblComment.PreferredWidth),
        ($topY + 3)
    )

    $gridTop = ([Math]::Max($picLogo.Bottom, $txtFSPath.Bottom) + 12)

    $bottomY = $form.ClientSize.Height - $status.Height - $padding
    $txtLog.SetBounds($padding, ($bottomY - $logHeight), ($form.ClientSize.Width - (2 * $padding)), $logHeight)

    $btnWidth = [int](($form.ClientSize.Width - (2 * $padding) - (2 * $gap)) / 3)
    if ($btnWidth -lt 200) { $btnWidth = 200 }

    $btnY = $txtLog.Top - 10 - $btnHeight
    $btnLatest.SetBounds($padding, $btnY, $btnWidth, $btnHeight)
    $btnMecm.SetBounds(($padding + $btnWidth + $gap), $btnY, $btnWidth, $btnHeight)
    $btnRun.SetBounds(($padding + (2 * ($btnWidth + $gap))), $btnY, $btnWidth, $btnHeight)

    $chkDebug.Location = New-Object System.Drawing.Point($padding, ($btnY - 34))

    $grid.SetBounds($padding, $gridTop, ($form.ClientSize.Width - (2 * $padding)), ($chkDebug.Top - 10 - $gridTop))

    Set-SelectAllCheckboxPosition
}

$form.Add_Shown({ Set-UILayout })
$form.Add_Resize({ Set-UILayout })

# -----------------------------
# Data model
# -----------------------------
$dt = New-Object System.Data.DataTable
[void]$dt.Columns.Add("Selected", [bool])
[void]$dt.Columns.Add("Vendor", [string])
[void]$dt.Columns.Add("Application", [string])
[void]$dt.Columns.Add("CurrentVersion", [string])
[void]$dt.Columns.Add("LatestVersion", [string])
[void]$dt.Columns.Add("Status", [string])
[void]$dt.Columns.Add("CMName", [string])
[void]$dt.Columns.Add("Script", [string])
[void]$dt.Columns.Add("FullPath", [string])

$grid.DataSource = $dt

$grid.Columns["Selected"].DataPropertyName = "Selected"
$grid.Columns["Vendor"].DataPropertyName = "Vendor"
$grid.Columns["Application"].DataPropertyName = "Application"
$grid.Columns["CurrentVersion"].DataPropertyName = "CurrentVersion"
$grid.Columns["LatestVersion"].DataPropertyName = "LatestVersion"
$grid.Columns["Status"].DataPropertyName = "Status"
$grid.Columns["CMName"].DataPropertyName = "CMName"
$grid.Columns["Script"].DataPropertyName = "Script"

$grid.Add_CellBeginEdit({
    param($s, $e)
    if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
        $path = [string]$dt.Rows[$e.RowIndex]["FullPath"]
        if ($path -match "\.notps1$") {
            $e.Cancel = $true
        }
    }
})

$grid.Add_RowPrePaint({
    param($s, $e)
    try {
        $path = [string]$dt.Rows[$e.RowIndex]["FullPath"]
        if ($path -match "\.notps1$") {
            $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGray
        }
    }
    catch { }
})

# -----------------------------
# Events
# -----------------------------
$chkDebug.Add_CheckedChanged({
    $show = $chkDebug.Checked
    $grid.Columns["CMName"].Visible = $show
    $grid.Columns["Script"].Visible = $show
})

$btnLatest.Add_Click({
    $siteCodeValue = $txtSiteCode.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -TextBox $txtLog -Message "SiteCode is required."
        $statusLabel.Text = "SiteCode is required."
        return
    }

    $selectedRows = @()
    foreach ($r in $dt.Rows) {
        if ($r["Selected"] -eq $true) { $selectedRows += $r }
    }

    if (-not $selectedRows -or $selectedRows.Count -eq 0) {
        Add-LogLine -TextBox $txtLog -Message "No rows selected."
        return
    }

    $btnLatest.Enabled = $false
    $btnMecm.Enabled   = $false
    $btnRun.Enabled    = $false
    $form.UseWaitCursor = $true

    try {
        $statusLabel.Text = "Checking latest versions for selected packagers..."

        foreach ($row in $selectedRows) {
            [System.Windows.Forms.Application]::DoEvents()

            $app     = [string]$row["Application"]
            $script  = [string]$row["Script"]
            $path    = [string]$row["FullPath"]

            Add-LogLine -TextBox $txtLog -Message ("Latest: {0} ({1})" -f $app, $script)
            $row["Status"] = "Checking latest..."

            try {
                $latest = Invoke-PackagerGetLatestVersion -PackagerPath $path -SiteCode $siteCodeValue -FileServerPath $txtFSPath.Text.Trim()
                $row["LatestVersion"] = $latest

                $current = [string]$row["CurrentVersion"]
                if (-not [string]::IsNullOrWhiteSpace($current)) {
                    $cmp = Compare-SemVer -A $current -B $latest
                    if ($cmp -lt 0) {
                        $row["Status"] = "Update available"
                    }
                    elseif ($cmp -eq 0) {
                        $row["Status"] = "Up to date"
                    }
                    else {
                        $row["Status"] = "Current newer"
                    }
                }
                else {
                    $row["Status"] = "Latest retrieved"
                }

                Add-LogLine -TextBox $txtLog -Message ("Latest version: {0}" -f $latest)
            }
            catch {
                $row["Status"] = "Error"
                Add-LogLine -TextBox $txtLog -Message ("Error: {0}" -f $_.Exception.Message)
            }
        }

        $statusLabel.Text = "Latest check complete."
    }
    finally {
        $form.UseWaitCursor = $false
        $btnLatest.Enabled = $true
        $btnMecm.Enabled   = $true
        $btnRun.Enabled    = $true
    }
})

$btnMecm.Add_Click({
    $siteCodeValue = $txtSiteCode.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -TextBox $txtLog -Message "SiteCode is required."
        $statusLabel.Text = "SiteCode is required."
        return
    }

    $selectedRows = @()
    foreach ($r in $dt.Rows) {
        if ($r["Selected"] -eq $true) { $selectedRows += $r }
    }

    if (-not $selectedRows -or $selectedRows.Count -eq 0) {
        Add-LogLine -TextBox $txtLog -Message "No rows selected."
        return
    }

    $btnLatest.Enabled = $false
    $btnMecm.Enabled   = $false
    $btnRun.Enabled    = $false
    $form.UseWaitCursor = $true

    try {
        $statusLabel.Text = "Querying MECM for selected products..."

        foreach ($row in $selectedRows) {
            [System.Windows.Forms.Application]::DoEvents()

            $app    = [string]$row["Application"]
            $cmName = [string]$row["CMName"]

            Add-LogLine -TextBox $txtLog -Message ("MECM: {0}" -f $app)
            $row["Status"] = "Querying MECM..."

            try {
                $res = Get-MecmCurrentVersionByCMName -SiteCode $siteCodeValue -CMName $cmName

                if (-not $res.Found) {
                    $row["CurrentVersion"] = ""
                    $row["Status"] = "Not found in MECM"
                    Add-LogLine -TextBox $txtLog -Message "Not found."
                    continue
                }

                $row["CurrentVersion"] = [string]$res.SoftwareVersion

                $latest = [string]$row["LatestVersion"]
                if (-not [string]::IsNullOrWhiteSpace($latest) -and -not [string]::IsNullOrWhiteSpace($res.SoftwareVersion)) {
                    $cmp = Compare-SemVer -A ([string]$res.SoftwareVersion) -B $latest
                    if ($cmp -lt 0)      { $row["Status"] = "Update available" }
                    elseif ($cmp -eq 0)  { $row["Status"] = "Up to date" }
                    else                 { $row["Status"] = "Current newer" }
                }
                else {
                    $row["Status"] = "MECM version retrieved"
                }

                if ($res.MatchCount -gt 1) {
                    Add-LogLine -TextBox $txtLog -Message ("Found {0} matches; using: {1} ({2})" -f $res.MatchCount, $res.DisplayName, $res.SoftwareVersion)
                }
                else {
                    Add-LogLine -TextBox $txtLog -Message ("Current version: {0}" -f $res.SoftwareVersion)
                }
            }
            catch {
                $row["Status"] = "Error"
                Add-LogLine -TextBox $txtLog -Message ("Error: {0}" -f $_.Exception.Message)
            }
        }

        Select-OnlyUpdateAvailable -DataTable $dt
        $grid.EndEdit()
        $grid.Refresh()

        $statusLabel.Text = "MECM query complete."
    }
    finally {
        $form.UseWaitCursor = $false
        $btnLatest.Enabled = $true
        $btnMecm.Enabled   = $true
        $btnRun.Enabled    = $true
    }
})

$btnRun.Add_Click({
    $siteCodeValue = $txtSiteCode.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($siteCodeValue)) {
        Add-LogLine -TextBox $txtLog -Message "SiteCode is required."
        $statusLabel.Text = "SiteCode is required."
        return
    }

    $commentValue = $txtComment.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($commentValue)) {
        Add-LogLine -TextBox $txtLog -Message "Work Order / Comment is required."
        $statusLabel.Text = "Work Order / Comment is required."
        return
    }
    $fsPathValue = $txtFSPath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($fsPathValue)) {
        Add-LogLine -TextBox $txtLog -Message "File Share Root is required."
        $statusLabel.Text = "File Share Root is required."
        return
    }

    $selectedRows = @()
    foreach ($r in $dt.Rows) {
        if ($r["Selected"] -eq $true) { $selectedRows += $r }
    }

    if (-not $selectedRows -or $selectedRows.Count -eq 0) {
        Add-LogLine -TextBox $txtLog -Message "No rows selected."
        return
    }

    $btnLatest.Enabled = $false
    $btnMecm.Enabled   = $false
    $btnRun.Enabled    = $false
    $form.UseWaitCursor = $true

    try {
        $logFolder = Join-Path $PSScriptRoot "Logs"
        $statusLabel.Text = "Running selected packagers..."

        foreach ($row in $selectedRows) {
            [System.Windows.Forms.Application]::DoEvents()

            $app    = [string]$row["Application"]
            $script = [string]$row["Script"]
            $path   = [string]$row["FullPath"]

            $row["Status"] = "Running..."
            Add-LogLine -TextBox $txtLog -Message ("Run: {0} ({1})" -f $app, $script)

            try {
                $res = Invoke-PackagerRun -PackagerPath $path -SiteCode $siteCodeValue -Comment $commentValue -FileServerPath $fsPathValue -LogFolder $logFolder

                if ($res.ExitCode -eq 0) {
                    $row["Status"] = "Complete"
                    Add-LogLine -TextBox $txtLog -Message ("Complete. Logs: {0}" -f (Split-Path -Leaf $res.OutLog))
                }
                else {
                    $row["Status"] = "Error"

                    $stderrLines = @($res.StdErr -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })

                    if ($stderrLines.Count -gt 0) {
                        $linesToShow = [Math]::Min($stderrLines.Count, 10)
                        for ($i = 0; $i -lt $linesToShow; $i++) {
                            Add-LogLine -TextBox $txtLog -Message ("  stderr: {0}" -f $stderrLines[$i])
                        }
                        if ($stderrLines.Count -gt $linesToShow) {
                            Add-LogLine -TextBox $txtLog -Message ("  ... and {0} more line(s) in: {1}" -f ($stderrLines.Count - $linesToShow), (Split-Path -Leaf $res.ErrLog))
                        }
                    }
                    else {
                        Add-LogLine -TextBox $txtLog -Message ("Error: Exit code {0}, no stderr output." -f $res.ExitCode)
                    }

                    Add-LogLine -TextBox $txtLog -Message ("Logs: {0}" -f (Split-Path -Leaf $res.OutLog))
                }
            }
            catch {
                $row["Status"] = "Error"
                Add-LogLine -TextBox $txtLog -Message ("Error: {0}" -f $_.Exception.Message)
            }
        }

        $statusLabel.Text = "Run complete."
    }
    finally {
        $form.UseWaitCursor = $false
        $btnLatest.Enabled = $true
        $btnMecm.Enabled   = $true
        $btnRun.Enabled    = $true
    }
})

$form.Add_Shown({
    Add-LogLine -TextBox $txtLog -Message ("Loading packagers from: {0}" -f $PackagersRoot)

    $items = Get-Packagers -Root $PackagersRoot
    foreach ($m in $items) {
        $row = $dt.NewRow()
        $row["Selected"] = $false
        $row["Vendor"] = $m.Vendor
        $row["Application"] = $m.Application
        $row["CurrentVersion"] = ""
        $row["LatestVersion"] = ""
        $row["Status"] = $m.Status
        $row["CMName"] = $m.CMName
        $row["Script"] = $m.Script
        $row["FullPath"] = $m.FullPath
        $dt.Rows.Add($row) | Out-Null
    }

    $statusLabel.Text = ("Loaded {0} packager(s). Ready." -f $dt.Rows.Count)
})

[void]$form.ShowDialog()
