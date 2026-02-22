<#
.SYNOPSIS
    Shared logging and download helpers for AppPackager packager scripts.

.DESCRIPTION
    Dot-source this file at the top of every packager script to get structured
    logging (Write-Log) and download retry (Invoke-DownloadWithRetry).

    Designed to migrate directly into a PowerShell module (AppPackagerCommon.psm1)
    when punchlist item #4 (Code Consolidation) is implemented.

.EXAMPLE
    . "$PSScriptRoot\AppPackagerCommon.ps1"
    Initialize-Logging -LogPath $LogPath

    Write-Log "Starting packager..."
    Write-Log "Something went wrong" -Level ERROR
    Invoke-DownloadWithRetry -Url $url -OutFile $file
#>

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

$script:__AppPackagerLogPath = $null

function Initialize-Logging {
    param([string]$LogPath)

    $script:__AppPackagerLogPath = $LogPath

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, severity-tagged log message.

    .DESCRIPTION
        INFO  -> Write-Host (stdout)
        WARN  -> Write-Host (stdout)
        ERROR -> Write-Host (stdout) + $host.UI.WriteErrorLine (stderr)

        -Quiet suppresses all console output but still writes to the log file.
    #>
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',

        [switch]$Quiet
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[{0}] [{1,-5}] {2}" -f $timestamp, $Level, $Message

    if (-not $Quiet) {
        Write-Host $formatted

        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__AppPackagerLogPath) {
        Add-Content -LiteralPath $script:__AppPackagerLogPath -Value $formatted -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Download with retry
# ---------------------------------------------------------------------------

function Invoke-DownloadWithRetry {
    <#
    .SYNOPSIS
        Downloads a file via curl.exe with a single retry on failure.

    .DESCRIPTION
        Wraps curl.exe file-download calls (curl.exe -L --fail --silent --show-error -o <file> <url>)
        with retry logic. Throws on final failure.

        Does NOT wrap scraping/variable-capture calls or URL-resolution calls.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Url,

        [Parameter(Mandatory)]
        [string]$OutFile,

        [string[]]$ExtraCurlArgs = @(),

        [int]$RetryCount = 1,

        [int]$RetryDelaySec = 5,

        [switch]$Quiet
    )

    $maxAttempts = 1 + $RetryCount

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        if ($attempt -gt 1) {
            Write-Log ("Retrying download (attempt {0} of {1}) after {2}s delay..." -f $attempt, $maxAttempts, $RetryDelaySec) -Level WARN -Quiet:$Quiet
            Start-Sleep -Seconds $RetryDelaySec
        }

        $allArgs = @('-L', '--fail', '--silent', '--show-error') + $ExtraCurlArgs + @('-o', $OutFile, $Url)
        & curl.exe @allArgs 2>$null
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            return
        }

        if ($attempt -lt $maxAttempts) {
            Write-Log ("Download attempt {0} failed (curl exit code {1}). Will retry." -f $attempt, $exitCode) -Level WARN -Quiet:$Quiet
        }
    }

    $msg = "Download failed after $maxAttempts attempt(s): $Url"
    Write-Log $msg -Level ERROR -Quiet:$Quiet
    throw $msg
}
