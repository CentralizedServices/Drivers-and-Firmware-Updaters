<#
.SYNOPSIS
    Downloads Windows Drivers & Firmware to a custom folder (No Install).

.DESCRIPTION
    Version 2.6 - RELIABILITY & CORRECTNESS FIXES

    Version History:
        2.5  - FIX: Extraction Logic now uses a "Snapshot Diff" of the SoftwareDistribution folder.
                    (Previous version failed because BITS preserves old file timestamps from 2018 etc.)
               ADDED: Base64 Fallback for Branding. If the URL fails (DNS error), it generates a generic icon.
               IMPROVED: Debug logging for file matching.
               ADD: Logo cache validation + skip download when cached logo is valid (optional max-age refresh)

        2.6  - FIX [#4]:  Log typo corrected — "INFOs`nINFO" was written as the log Type in
                           Update-BrandingCache. Changed to "INFO".
               FIX [#6]:  Write-RunSummary is now called at every exit point. Previously the function
                           was defined but never invoked — no scanner exit status was ever recorded on disk.
               FIX [#11]: Size-based fallback (both in the diff loop and the global fallback) now requires
                           an unambiguous match (exactly 1 file at that size). If multiple files share the
                           same size, the match is skipped and logged as ambiguous to prevent copying the
                           wrong driver into the wrong update folder.
               FIX [#12]: Test-PngSignature now uses -AsByteStream on PS 6+; -Encoding Byte on PS 5.
                           Prevents runtime error if CW RMM ever invokes this via pwsh.exe (PS 7).
               FIX [#15]: Final trigger decision now uses $DownloadedList.Count (files extracted THIS
                           session) instead of total folder size. The old size check included cached files
                           from previous days, causing TRIGGER=UPDATES_READY to fire on re-runs even when
                           nothing new was downloaded.
               ADD:       Explicit TRIGGER=UPDATES_READY log line written on success. DFPopup checks for
                           this pattern first, making trigger detection unambiguous and version-stable.

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
    [string]$LogPath         = "C:\ProgramData\i3\logs",
    [string]$BrandLogoUrl    = "https://s3.us-east-1.wasabisys.com/i3-blob-storage/Installers/i3_CharmOnly_400x300.png",
    [switch]$Silent,
    [int]$LogoMaxAgeDays     = 0
)

$EXIT_OK                               = 0
$EXIT_UPDATES_FOUND_BUT_NONE_EXTRACTED = 10
$EXIT_CRITICAL_FAILURE                 = 20
$EXIT_NOT_ADMIN                        = 21
$EXIT_WU_SERVICES_UNAVAILABLE          = 22

$ScriptVersion  = "2.6"
$RunStamp       = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile        = Join-Path -Path $LogPath -ChildPath "DriverDownload_$(Get-Date -Format 'yyyyMMdd').log"
$RunSummaryPath = Join-Path -Path $LogPath -ChildPath "DriverDownload_RunSummary_${RunStamp}.json"

# ---------------------------------------------------------------------------
# FIX #6: Script-level summary hashtable.
# Write-RunSummary was previously defined but never called. This hashtable is
# populated throughout the run and flushed at every exit point.
# ---------------------------------------------------------------------------
$Script:ScanSummary = @{
    Version      = $ScriptVersion
    RunStamp     = $RunStamp
    StartTime    = (Get-Date).ToString("o")
    EndTime      = $null
    ComputerName = $env:COMPUTERNAME
    ExitCode     = $null
    UpdatesFound = 0
    Downloaded   = @()
    Failures     = @()
    LogFile      = $null   # assigned after $LogPath is confirmed
}

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

# FIX #6: Write-RunSummary now flushes the script-level summary hashtable.
# Called at every exit point — previously this function existed but was never invoked.
function Write-RunSummary {
    param([int]$ExitCode)
    try {
        $Script:ScanSummary.EndTime  = (Get-Date).ToString("o")
        $Script:ScanSummary.ExitCode = $ExitCode
        $Script:ScanSummary | ConvertTo-Json -Depth 10 | Out-File -FilePath $RunSummaryPath -Force -Encoding UTF8
    } catch {
        Write-Log "Summary write failed: $($_.Exception.Message)" "WARNING"
    }
}

# ---------------------------------------------------------------------------
# PNG VALIDATION (for logo caching)
# ---------------------------------------------------------------------------

# FIX #12: -Encoding Byte is deprecated in PS 6+. Use -AsByteStream on PS 6+
# and fall back to -Encoding Byte on PS 5. Prevents failure if invoked via pwsh.exe.
function Test-PngSignature {
    param([string]$Path)
    try {
        if (-not (Test-Path $Path)) { return $false }

        if ($PSVersionTable.PSVersion.Major -ge 6) {
            $bytes = Get-Content -Path $Path -AsByteStream -TotalCount 8 -ErrorAction Stop
        } else {
            $bytes = Get-Content -Path $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop
        }

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
        if ($MaxAgeDays -le 0) { return $false }
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

        foreach ($p in @($LogoPath, $LastGood)) {
            if (Test-Path $p) {
                try { Set-ItemProperty -Path $p -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch {}
            }
        }

        $needDownload = $true

        if (Test-LogoOk -Path $LogoPath) {
            if (-not (Test-LogoTooOld -Path $LogoPath -MaxAgeDays $LogoMaxAgeDays)) {
                $needDownload = $false
                # FIX #4: Was Write-Log "..." "INFOs`nINFO" (multi-line typo). Corrected to "INFO".
                Write-Log "Branding: Using cached logo.png (valid)." "INFO"
            } else {
                Write-Log "Branding: Cached logo.png is older than $LogoMaxAgeDays days; will refresh." "INFO"
            }
        } else {
            if (Test-Path $LogoPath) {
                Remove-Item $LogoPath -Force -ErrorAction SilentlyContinue
                Write-Log "Branding: Removed invalid logo.png." "WARNING"
            }
        }

        if ($needDownload -and (Test-LogoOk -Path $LastGood)) {
            Copy-Item $LastGood $LogoPath -Force -ErrorAction SilentlyContinue
            if (Test-LogoOk -Path $LogoPath) {
                $needDownload = $false
                Write-Log "Branding: Restored logo.png from logo.lastgood.png (no download)." "SUCCESS"
            }
        }

        if ($needDownload) {
            Write-Log "Branding: Downloading from $Url..." "INFO"
            Invoke-WebRequest -Uri $Url -OutFile $LogoPath -UseBasicParsing -ErrorAction Stop
            Unblock-File -Path $LogoPath -ErrorAction SilentlyContinue

            if (-not (Test-LogoOk -Path $LogoPath)) {
                Remove-Item $LogoPath -Force -ErrorAction SilentlyContinue
                throw "Downloaded logo failed validation"
            }

            Copy-Item $LogoPath $LastGood -Force -ErrorAction SilentlyContinue
            Write-Log "Branding: Download success (validated) + lastgood updated." "SUCCESS"
        }

        foreach ($p in @($LogoPath, $LastGood)) {
            if (Test-Path $p) {
                try { Set-ItemProperty -Path $p -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue } catch {}
            }
        }

    } catch {
        Write-Log "Branding: Download/cache failed ($($_.Exception.Message)). Using Fallback." "WARNING"
        try {
            $B64 = "iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAYAAACtWK6eAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAA3HpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8++IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwveDp4bXBtZXRhPiA8P3hwYWNrZXQgZW5kPSJyIj8+1S41JAAAC+BJREFUeNrsnQtwVNUZx/9333k8NiQECAkhDxIeCY9AwBAeQ7E8BGTEdpS2M05nKD5wqqOij2GqM310tLbaoVIdR6qjU61jRjufCAI+wBAFBAIBEhKAhDw2yW6y2f3O9a633CSS3Zvc3bv35j/fzGSy2XvPvbvnnO/7zvnOuec4wzCAiMioeJwuABFJJIAIJBJABBIJIAKJBBCBRAKIQCIBCME4nC5AuImJicK0adMwfvx4TJgwAQ6HwzR9r9eLzs5O1NfXo6amBvX19abtIyLxjUAgI4aUlBQkJSVhxowZmDlzJmJjY8P6Hn6/H+fPn0dFRQUqKipQXV0d1u8jIoEEs/iIj4/H3Ll"
            $B64Bytes = [System.Convert]::FromBase64String($B64)
            [System.IO.File]::WriteAllBytes($LogoPath, $B64Bytes)

            if (Test-LogoOk -Path $LogoPath) {
                Copy-Item $LogoPath $LastGood -Force -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $LogoPath -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $LastGood  -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue
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
    # FIX #6: Write-RunSummary now called before every exit
    Write-RunSummary -ExitCode $EXIT_NOT_ADMIN
    exit $EXIT_NOT_ADMIN
}

# 2) Prep Folders
if (-not (Test-Path $DestinationPath)) { New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $LogPath))         { New-Item -Path $LogPath         -ItemType Directory -Force | Out-Null }

# LogPath is now confirmed — record it in the summary
$Script:ScanSummary.LogFile = $LogFile

# 3) Update Branding (With cache + fallback)
if (-not [string]::IsNullOrWhiteSpace($BrandLogoUrl)) { Update-BrandingCache -Url $BrandLogoUrl }

$DownloadedList   = @()
$SessionFailures  = @()
$SoftwareDistPath = "$env:SystemRoot\SoftwareDistribution\Download"

# 4) WU Services
if (-not ((Get-Service wuauserv -ErrorAction SilentlyContinue).Status -eq 'Running')) {
    Start-Service wuauserv -ErrorAction SilentlyContinue
}

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
    Write-Log "Found $Count driver update(s)." "INFO"
    $Script:ScanSummary.UpdatesFound = $Count

    if ($Count -eq 0) {
        Write-Log "No updates found." "SUCCESS"
        # FIX #6: Write-RunSummary called here (was missing)
        Write-RunSummary -ExitCode $EXIT_OK
        exit $EXIT_OK
    }

    # 7) Download
    $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($U in $SearchResult.Updates) {
        if (-not $U.EulaAccepted) { $U.AcceptEula() }
        $UpdatesToDownload.Add($U) | Out-Null
    }

    $Downloader         = $Session.CreateUpdateDownloader()
    $Downloader.Updates = $UpdatesToDownload

    Write-Log "Downloading $Count update(s)..." "INFO"
    $Res = $Downloader.Download()
    Write-Log "Download result code: $($Res.ResultCode)" "INFO"

    # 8) SNAPSHOT AFTER + DIFF
    Write-Log "Taking post-download snapshot..." "INFO"
    $NewFiles = @()
    Get-ChildItem -Path $SoftwareDistPath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        if (-not $SnapshotBefore.Contains($_.FullName)) { $NewFiles += $_ }
    }
    Write-Log "Detected $($NewFiles.Count) new file(s) downloaded." "INFO"

    # 9) EXTRACTION (Using only NewFiles from this session)
    foreach ($Update in $SearchResult.Updates) {
        if ($Update.IsDownloaded) {
            $SafeTitle = Get-SafeFilename -Name $Update.Title
            $DestDir   = Join-Path $DestinationPath $SafeTitle
            if (-not (Test-Path $DestDir)) { New-Item $DestDir -ItemType Directory -Force | Out-Null }

            $Found    = $false
            $Payloads = @()
            if ($Update.BundledUpdates.Count -gt 0) { $Payloads += $Update.BundledUpdates } else { $Payloads += $Update }

            foreach ($P in $Payloads) {
                foreach ($C in $P.DownloadContents) {
                    $TargetSize = [int64]$P.MaxDownloadSize
                    $TargetName = if ($C.DownloadUrl) { $C.DownloadUrl.Split('/')[-1] } else { "Unknown" }

                    # --- Match 1: Exact filename among new files (most reliable) ---
                    $Match = $NewFiles | Where-Object { $_.Name -eq $TargetName } | Select-Object -First 1

                    # --- Match 2: Size among new files ---
                    # FIX #11: Only accept if exactly ONE new file has this size (unambiguous).
                    # The original code used Select-Object -First 1 with no ambiguity check — two
                    # different drivers with the same file size could silently cross-copy each other.
                    if (-not $Match -and $TargetSize -gt 0) {
                        $sizeMatches = $NewFiles | Where-Object { $_.Length -eq $TargetSize }
                        if ($sizeMatches.Count -eq 1) {
                            $Match = $sizeMatches[0]
                            Write-Log "  Size-match (unambiguous, 1 candidate): $($Match.Name) for '$SafeTitle'" "INFO"
                        } elseif ($sizeMatches.Count -gt 1) {
                            Write-Log "  Size-match AMBIGUOUS ($($sizeMatches.Count) files at $TargetSize bytes) for '$SafeTitle' — skipping size fallback to avoid cross-copy." "WARNING"
                        }
                    }

                    if ($Match) {
                        Copy-Item -LiteralPath $Match.FullName -Destination (Join-Path $DestDir $Match.Name) -Force
                        $Found = $true
                        Write-Log "  Matched and copied: $($Match.Name)" "INFO"
                    }
                }
            }

            if ($Found) {
                Write-Log "Extracted: $SafeTitle" "SUCCESS"
                $DownloadedList += $SafeTitle
            } else {
                # --- Global fallback: file may already have been cached from a prior incomplete run ---
                # FIX #11: Same unambiguous-match rule applies here. Only accept if exactly 1 file
                # in the entire SoftwareDistribution folder matches the size.
                $allFiles     = Get-ChildItem $SoftwareDistPath -Recurse -File -ErrorAction SilentlyContinue
                $fallbackSize = [int64]$Update.MaxDownloadSize
                $sizeMatchAll = $allFiles | Where-Object { $_.Length -eq $fallbackSize }

                if ($sizeMatchAll.Count -eq 1) {
                    Copy-Item -LiteralPath $sizeMatchAll[0].FullName -Destination (Join-Path $DestDir $sizeMatchAll[0].Name) -Force
                    Write-Log "Extracted (cached fallback, unambiguous): $SafeTitle" "SUCCESS"
                    $DownloadedList += $SafeTitle
                } elseif ($sizeMatchAll.Count -gt 1) {
                    Write-Log "Failed to extract — global size match AMBIGUOUS ($($sizeMatchAll.Count) candidates at $fallbackSize bytes): $SafeTitle" "WARNING"
                    $SessionFailures += $SafeTitle
                    Remove-Item $DestDir -Force -Recurse -ErrorAction SilentlyContinue
                } else {
                    Write-Log "Failed to extract (no match found anywhere): $SafeTitle" "WARNING"
                    $SessionFailures += $SafeTitle
                    Remove-Item $DestDir -Force -Recurse -ErrorAction SilentlyContinue
                }
            }
        }
    }

} catch {
    Write-Log "CRITICAL: $($_.Exception.Message)" "ERROR"
    # FIX #6: Write-RunSummary called here (was missing)
    Write-RunSummary -ExitCode $EXIT_CRITICAL_FAILURE
    exit $EXIT_CRITICAL_FAILURE
}

# ---------------------------------------------------------------------------
# FINAL CHECK & EXIT
# ---------------------------------------------------------------------------

# FIX #15: Base the trigger decision on $DownloadedList.Count (files extracted THIS session),
# not total folder size. The old check used Get-ChildItem on $DestinationPath which included
# driver files cached from previous days — this caused TRIGGER=UPDATES_READY to fire on
# every re-run even when today's scan found and extracted nothing new.
$Script:ScanSummary.Downloaded = $DownloadedList
$Script:ScanSummary.Failures   = $SessionFailures

if ($DownloadedList.Count -gt 0) {
    # ADD: Explicit machine-readable trigger line.
    # DFPopup checks for 'TRIGGER=UPDATES_READY' as its first pattern — this makes
    # trigger detection unambiguous and independent of the human-readable summary text.
    Write-Log "TRIGGER=UPDATES_READY" "INFO"
    Write-Log "Updates ready in $DestinationPath. Drivers extracted: $($DownloadedList.Count). Failures: $($SessionFailures.Count)." "SUCCESS"

    # FIX #6: Write-RunSummary called here (was missing)
    Write-RunSummary -ExitCode $EXIT_OK
    exit $EXIT_OK
} else {
    Write-Log "No driver files extracted this session. Updates found: $($SearchResult.Updates.Count). Failures: $($SessionFailures.Count)." "WARNING"

    # FIX #6: Write-RunSummary called here (was missing)
    Write-RunSummary -ExitCode $EXIT_UPDATES_FOUND_BUT_NONE_EXTRACTED
    exit $EXIT_UPDATES_FOUND_BUT_NONE_EXTRACTED
}
