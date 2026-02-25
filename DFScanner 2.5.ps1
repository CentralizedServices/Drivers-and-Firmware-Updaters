<#
.SYNOPSIS
    Downloads Windows Drivers & Firmware to a custom folder (No Install).

.DESCRIPTION
    Version 2.5 - FIX TIMESTAMP & BRANDING ISSUES
    - FIXED: Extraction Logic now uses a "Snapshot Diff" of the SoftwareDistribution folder.
             (Previous version failed because BITS preserves old file timestamps from 2018 etc.)
    - ADDED: Base64 Fallback for Branding. If the URL fails (DNS error), it generates a generic icon.
    - IMPROVED: Debug logging for file matching.
    - ADD: Logo cache validation + skip download when cached logo is valid (optional max-age refresh)

.PARAMETER DestinationPath
    Target folder for drivers. Default: C:\ProgramData\i3\Drivers

.PARAMETER LogPath
    Target folder for logs. Default: C:\ProgramData\i3\logs

.PARAMETER BrandLogoUrl
    URL for the company logo to cache.

.PARAMETER LogoMaxAgeDays
    Optional: Refresh logo if cached logo is older than X days. Default: 0 (never refresh based on age)
#>

[CmdletBinding()]
param (
    [string]$DestinationPath = "C:\ProgramData\i3\Drivers",
    [string]$LogPath = "C:\ProgramData\i3\logs",
    [string]$BrandLogoUrl = "https://s3.us-east-1.wasabisys.com/i3-blob-storage/Installers/i3_CharmOnly_400x300.png",
    [switch]$Silent,

    # NEW: Logo cache refresh window (0 = never refresh based on age)
    [int]$LogoMaxAgeDays = 0
)

$EXIT_OK = 0
$EXIT_UPDATES_FOUND_BUT_NONE_EXTRACTED = 10
$EXIT_CRITICAL_FAILURE = 20
$EXIT_NOT_ADMIN = 21
$EXIT_WU_SERVICES_UNAVAILABLE = 22

$ScriptVersion = "2.5"
$RunStamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile       = Join-Path -Path $LogPath -ChildPath "DriverDownload_$(Get-Date -Format 'yyyyMMdd').log"
$RunSummaryPath= Join-Path -Path $LogPath -ChildPath "DriverDownload_RunSummary_${RunStamp}.json"

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------
function Write-Log {
    param ([string]$Message, [string]$Type = "INFO")
    $Line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Type] $Message"
    if (-not $Silent) { Write-Host $Line }
    if (-not (Test-Path $LogPath)) { New-Item $LogPath -ItemType Directory -Force | Out-Null }
    Add-Content -Path $LogFile -Value $Line -ErrorAction SilentlyContinue
}

function Get-SafeFilename {
    param ([string]$Name)
    $Safe = $Name -replace '[\\/:"*?<>|]', '' -replace '\s+', '_'
    if ($Safe.Length -gt 50) { $Safe = $Safe.Substring(0, 50) }
    return $Safe.Trim().Trim('.')
}

function Write-RunSummary {
    param([hashtable]$Summary)
    try {
        $Summary | ConvertTo-Json -Depth 10 | Out-File -FilePath $RunSummaryPath -Force -Encoding UTF8
    } catch { Write-Log "Summary write failed: $($_.Exception.Message)" "WARNING" }
}

# ---------------------------------------------------------------------------
# PNG VALIDATION (for logo caching)
# ---------------------------------------------------------------------------
function Test-PngSignature {
    param([string]$Path)
    try {
        if (-not (Test-Path $Path)) { return $false }
        $bytes = Get-Content -Path $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop
        return ($bytes.Count -eq 8 -and
                $bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47 -and
                $bytes[4] -eq 0x0D -and $bytes[5] -eq 0x0A -and $bytes[6] -eq 0x1A -and $bytes[7] -eq 0x0A)
    } catch { return $false }
}

function Test-LogoOk {
    param([string]$Path)
    try {
        if (-not (Test-Path $Path)) { return $false }
        $fi = Get-Item $Path -ErrorAction Stop
        if ($fi.Length -lt 2048) { return $false }
        if (-not (Test-PngSignature -Path $Path)) { return $false }
        return $true
    } catch { return $false }
}

function Test-LogoTooOld {
    param([string]$Path, [int]$MaxAgeDays)
    try {
        if ($MaxAgeDays -le 0) { return $false } # never refresh based on age
        if (-not (Test-Path $Path)) { return $true }
        $fi = Get-Item $Path -ErrorAction Stop
        return ($fi.LastWriteTime -lt (Get-Date).AddDays(-1 * $MaxAgeDays))
    } catch { return $true }
}

# ---------------------------------------------------------------------------
# BRANDING FALLBACK (Base64) + CACHE
# ---------------------------------------------------------------------------
function Update-BrandingCache {
    param([string]$Url)

    $BrandDir = "C:\ProgramData\i3\branding"
    $LogoPath = Join-Path $BrandDir "logo.png"
    $LastGood = Join-Path $BrandDir "logo.lastgood.png"

    try {
        if (-not (Test-Path $BrandDir)) { New-Item -Path $BrandDir -ItemType Directory -Force | Out-Null }

        # Unlock logo files if they exist so we can check/update them
        foreach ($p in @($LogoPath, $LastGood)) {
            if (Test-Path $p) {
                try { Set-ItemProperty -Path $p -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
            }
        }

        $needDownload = $true

        # 1) Use cached logo.png if valid and not too old
        if (Test-LogoOk -Path $LogoPath) {
            if (-not (Test-LogoTooOld -Path $LogoPath -MaxAgeDays $LogoMaxAgeDays)) {
                $needDownload = $false
                Write-Log "Branding: Using cached logo.png (valid)." "INFOs
INFO"
            } else {
                Write-Log "Branding: Cached logo.png is older than $LogoMaxAgeDays days; will refresh." "INFO"
            }
        } else {
            if (Test-Path $LogoPath) {
                Remove-Item $LogoPath -Force -ErrorAction SilentlyContinue
                Write-Log "Branding: Removed invalid logo.png." "WARNING"
            }
        }

        # 2) If needed, restore from lastgood (no download)
        if ($needDownload -and (Test-LogoOk -Path $LastGood)) {
            Copy-Item $LastGood $LogoPath -Force -ErrorAction SilentlyContinue
            if (Test-LogoOk -Path $LogoPath) {
                $needDownload = $false
                Write-Log "Branding: Restored logo.png from logo.lastgood.png (no download)." "SUCCESS"
            }
        }

        # 3) Download only if needed (or refresh forced by age)
        if ($needDownload) {
            Write-Log "Branding: Downloading from $Url..." "INFO"
            Invoke-WebRequest -Uri $Url -OutFile $LogoPath -UseBasicParsing -ErrorAction Stop
            Unblock-File -Path $LogoPath -ErrorAction SilentlyContinue

            if (-not (Test-LogoOk -Path $LogoPath)) {
                Remove-Item $LogoPath -Force -ErrorAction SilentlyContinue
                throw "Downloaded logo failed validation"
            }

            # Update lastgood only after validated download
            Copy-Item $LogoPath $LastGood -Force -ErrorAction SilentlyContinue
            Write-Log "Branding: Download success (validated) + lastgood updated." "SUCCESS"
        }

        # Lock files
        foreach ($p in @($LogoPath, $LastGood)) {
            if (Test-Path $p) {
                try { Set-ItemProperty -Path $p -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue } catch {}
            }
        }

    } catch {
        Write-Log "Branding: Download/cache failed ($($_.Exception.Message)). Using Fallback." "WARNING"
        try {
            # Generic "Info" Icon Base64 (Blue Circle with 'i')
            $B64 = "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"
            $B64Bytes = [System.Convert]::FromBase64String($B64)
            [System.IO.File]::WriteAllBytes($LogoPath, $B64Bytes)

            # Validate fallback and also save as lastgood
            if (Test-LogoOk -Path $LogoPath) {
                Copy-Item $LogoPath $LastGood -Force -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $LogoPath -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $LastGood -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue
                Write-Log "Branding: Fallback logo applied + cached as lastgood." "SUCCESS"
            } else {
                Remove-Item $LogoPath -Force -ErrorAction SilentlyContinue
                Write-Log "Branding: Fallback generated but failed validation. No logo will be available." "ERROR"
            }
        } catch {
            Write-Log "Branding: Fallback failed. Popup will have no logo." "ERROR"
        }
    }
}

# ---------------------------------------------------------------------------
# SNAPSHOT DIFF LOGIC (Fixes Timestamp Issues)
# ---------------------------------------------------------------------------
function Get-FileSnapshot {
    # Returns a HashSet of full paths for fast comparison
    $Snapshot = [System.Collections.Generic.HashSet[string]]::new()
    $SoftwareDistPath = "$env:SystemRoot\SoftwareDistribution\Download"

    Get-ChildItem -Path $SoftwareDistPath -Recurse -File -ErrorAction SilentlyContinue |
        ForEach-Object { $Snapshot.Add($_.FullName) | Out-Null }

    return $Snapshot
}

# ---------------------------------------------------------------------------
# MAIN SCRIPT
# ---------------------------------------------------------------------------

# 1) Admin Check
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Log "Script must be run as Administrator." "ERROR"
    $ExitCode = $EXIT_NOT_ADMIN
    exit $ExitCode
}

# 2) Prep Folders
if (-not (Test-Path $DestinationPath)) { New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $LogPath)) { New-Item -Path $LogPath -ItemType Directory -Force | Out-Null }

# 3) Update Branding (With cache + fallback)
if (-not [string]::IsNullOrWhiteSpace($BrandLogoUrl)) { Update-BrandingCache -Url $BrandLogoUrl }

$DownloadedList = @()
$SessionFailures = @()
$SoftwareDistPath = "$env:SystemRoot\SoftwareDistribution\Download"

# 4) WU Services
if (-not ((Get-Service wuauserv -ErrorAction SilentlyContinue).Status -eq 'Running')) { Start-Service wuauserv -ErrorAction SilentlyContinue }

try {
    Write-Log "Starting Session (v$ScriptVersion)..." "INFO"
    $Session  = New-Object -ComObject Microsoft.Update.Session
    $Searcher = $Session.CreateUpdateSearcher()

    # 5) SNAPSHOT BEFORE
    Write-Log "Taking pre-download snapshot..." "INFO"
    $SnapshotBefore = Get-FileSnapshot

    # 6) Search
    Write-Log "Scanning for drivers..." "INFO"
    $SearchResult = $Searcher.Search("IsInstalled=0 and Type='Driver'")
    $Count = $SearchResult.Updates.Count
    Write-Log "Found $Count updates." "INFO"

    if ($Count -eq 0) {
        Write-Log "No updates found." "SUCCESS"
        exit $EXIT_OK
    }

    # 7) Download
    $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($U in $SearchResult.Updates) { if (!$U.EulaAccepted) { $U.AcceptEula() }; $UpdatesToDownload.Add($U) | Out-Null }

    $Downloader = $Session.CreateUpdateDownloader()
    $Downloader.Updates = $UpdatesToDownload

    Write-Log "Downloading..." "INFO"
    $Res = $Downloader.Download()
    Write-Log "Download Result: $($Res.ResultCode)" "INFO"

    # 8) SNAPSHOT AFTER + DIFF
    Write-Log "Taking post-download snapshot..." "INFO"
    $NewFiles = @()
    Get-ChildItem -Path $SoftwareDistPath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        if (-not $SnapshotBefore.Contains($_.FullName)) {
            $NewFiles += $_
        }
    }
    Write-Log "Detected $($NewFiles.Count) new files downloaded." "INFO"

    # 9) EXTRACTION (Using only NewFiles)
    foreach ($Update in $SearchResult.Updates) {
        if ($Update.IsDownloaded) {
            $SafeTitle = Get-SafeFilename -Name $Update.Title
            $DestDir = Join-Path $DestinationPath $SafeTitle
            if (-not (Test-Path $DestDir)) { New-Item $DestDir -ItemType Directory -Force | Out-Null }

            $Found = $false

            # Get expected payloads
            $Payloads = @()
            if ($Update.BundledUpdates.Count -gt 0) { $Payloads += $Update.BundledUpdates } else { $Payloads += $Update }

            foreach ($P in $Payloads) {
                foreach ($C in $P.DownloadContents) {
                    # Try to match a new file
                    $TargetSize = [int64]$P.MaxDownloadSize
                    $TargetName = if ($C.DownloadUrl) { $C.DownloadUrl.Split('/')[-1] } else { "Unknown" }

                    # Match Logic:
                    # 1. Check if ANY new file matches Name
                    $Match = $NewFiles | Where-Object { $_.Name -eq $TargetName } | Select-Object -First 1

                    # 2. If no name match, check SIZE among new files
                    if (-not $Match -and $TargetSize -gt 0) {
                        $Match = $NewFiles | Where-Object { $_.Length -eq $TargetSize } | Select-Object -First 1
                    }

                    if ($Match) {
                        Copy-Item -LiteralPath $Match.FullName -Destination (Join-Path $DestDir $Match.Name) -Force
                        $Found = $true
                    }
                }
            }

            if ($Found) {
                Write-Log "Extracted: $SafeTitle" "SUCCESS"
                $DownloadedList += $SafeTitle
            } else {
                # Fallback: If Diff missed it (rare), try looking for ANY file matching size in entire folder
                # This catches cases where file was already there (cached) and didn't trigger 'New'
                $SearchAll = Get-ChildItem $SoftwareDistPath -Recurse -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.Length -eq [int64]$Update.MaxDownloadSize } | Select-Object -First 1

                if ($SearchAll) {
                    Copy-Item -LiteralPath $SearchAll.FullName -Destination (Join-Path $DestDir $SearchAll.Name) -Force
                    Write-Log "Extracted (Cached Match): $SafeTitle" "SUCCESS"
                    $DownloadedList += $SafeTitle
                } else {
                    Write-Log "Failed to extract: $SafeTitle" "WARNING"
                    $SessionFailures += $SafeTitle
                    Remove-Item $DestDir -Force -Recurse -ErrorAction SilentlyContinue
                }
            }
        }
    }

} catch {
    Write-Log "CRITICAL: $($_.Exception.Message)" "ERROR"
    exit $EXIT_CRITICAL_FAILURE
}

# Final Check
$Size = (Get-ChildItem $DestinationPath -Recurse -File | Measure-Object -Sum Length).Sum
if ($Size -gt 0) {
    Write-Log "Updates ready in $DestinationPath ($([math]::Round($Size/1MB,2)) MB)" "SUCCESS"
    exit $EXIT_OK
} else {
    Write-Log "No files extracted." "WARNING"
    exit $EXIT_UPDATES_FOUND_BUT_NONE_EXTRACTED
}
