<#
.SYNOPSIS
    Analyzes Driver Logs and prompts the user to restart if updates are ready.

.DESCRIPTION
    Runs as SYSTEM/Admin via RMM (CW).
    - Detects "updates ready" in today's driver download log
    - Skips popup entirely if Windows is already pending a reboot (exit 22)
    - Launches a user-context WPF popup via schtasks.exe /IT with DOMAIN\user
    - Waits on a JSON status file for user response
    - BEHAVIOR 1: If user clicks "Restart now", arm a SYSTEM "install at next boot" task + marker,
      then initiate reboot (optional).

    Version History:
        5.10  - ADD: Behavior 1 "Install at next boot" arming ONLY when user clicks Restart
        5.11  - FIX: Changed Schtasks scheduling to +1 Hour to prevent "Start time is in the past" race conditions.
        5.12  - IMP: Enhanced "Best Effort" installer logic with exit code validation (0, 3010, 259) and depth limits.
        5.13  - FIX [Critical #1]:      Start-Sleep -Seconds 3 added between Release-Lock and shutdown.exe.
                                         Write-RunSummary and Release-Lock already ran first; the sleep gives the
                                         filesystem a flush grace period before the OS begins teardown.
               FIX [Critical #3]:      Notes are now structured {Severity, Message} objects via Add-SummaryNote.
                                         Only 'Error'-severity notes surface as RMM alerts in DFReport — success
                                         events (e.g. "Armed install-at-boot") no longer cause false positives.
               FIX [Significant #5]:   Invoke-RunSummaryCleanup added — old RunSummary JSON files are
                                         auto-pruned on each run (retains last 10).
               FIX [Significant #8]:   Set-PublicReadWrite called once after interactive-user check to establish
                                         permissions. StatusFile content is reset to "Waiting" immediately before
                                         Invoke-PopupTask_Schtasks — not at the top of the script — eliminating
                                         the stale-state race condition.
               FIX [Moderate #9]:      $MaxSnoozes parameter removed (was declared but never implemented).
               FIX [Moderate #10]:     XAML TextBlocks use __XML_CAPTION__ / __XAML_BODY__ tokens (pre-computed
                                         in SYSTEM context) instead of PS variable interpolation inside the XAML
                                         here-string. Body lines are joined with <LineBreak/> elements — XML
                                         normalizes whitespace in attribute values to spaces, but element content
                                         with <LineBreak/> renders correctly.
               FIX [Moderate #12]:     Test-PngSignature uses -AsByteStream on PS 6+; -Encoding Byte on PS 5.
               FIX [Moderate #18]:     DISM /Add-Package corrected to /Add-Driver in the boot installer template.
                                         /Add-Package is for Windows packages; /Add-Driver is for driver CABs.
               FIX [Minor #14]:        Test-LogoTooOld added to popup; $LogoMaxAgeDays now applied in branding.
#>

[CmdletBinding()]
param(
    [string]$LogFolder           = 'C:\ProgramData\i3\logs',
    [string]$DriverSource        = 'C:\ProgramData\i3\Drivers',
    [string]$BrandLogoUrl        = 'https://s3.us-east-1.wasabisys.com/i3-blob-storage/Installers/i3_CharmOnly_400x300.png',
    [int]$MaxWaitTimeMinutes     = 55,
    [string]$Caption             = "i3 Updates Available",
    [string]$Body                = @"
New driver and firmware updates are ready to install.
Please save your work.
When you click 'Restart Now', updates will install and the PC will reboot automatically.
"@,
    [bool]$EnableDriverDiff      = $true,
    [int]$LogoMaxAgeDays         = 0,

    # Behavior 1 control:
    [bool]$RebootOnRestartChoice = $true
)

# ---------------------------------------------------------------------------
# EXIT CODES
# ---------------------------------------------------------------------------
$EXIT_NOACTION_OR_DONE          = 0
$EXIT_PROMPT_TIMEOUT_ASSUMED    = 10
$EXIT_USER_CHOSE_RESTART_START  = 11
$EXIT_INSTALL_FAILED            = 20
$EXIT_NO_INTERACTIVE_USER       = 21
$EXIT_PENDING_REBOOT_SKIP       = 22
$EXIT_INSTALL_IN_PROGRESS       = 23
$EXIT_ALREADY_RUNNING_LOCK      = 24

$ScriptVersion         = "5.13"
$ErrorActionPreference = 'Stop'

$InternalLogPath  = 'C:\ProgramData\i3\RestartPrompt.log'
$SnoozeMetricPath = 'C:\ProgramData\i3\logs\UserSnoozeMetrics.csv'
$StatusFile       = 'C:\ProgramData\i3\PromptStatus.json'

$RunStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$RunStart = Get-Date

# ---------------------------------------------------------------------------
# LOG ROTATION & INITIALIZATION
# ---------------------------------------------------------------------------
$null = New-Item -ItemType Directory -Force -Path (Split-Path $InternalLogPath) -ErrorAction SilentlyContinue

try
{
    if (Test-Path $InternalLogPath)
    {
        $logFile = Get-Item $InternalLogPath
        if ($logFile.Length -gt 5MB)
        {
            Move-Item $InternalLogPath "$InternalLogPath.bak" -Force -ErrorAction SilentlyContinue
            Write-Host "Log rotated."
        }
    }
}
catch
{
}

function Write-InternalLog
{
    param(
        [string]$Message
    )

    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] $Message"
    Write-Host $line
    Add-Content -Path $InternalLogPath -Value $line -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------------
# SINGLE-INSTANCE LOCK
# ---------------------------------------------------------------------------
$LockPath          = 'C:\ProgramData\i3\DriverPrompt.lock'
$LockMaxAgeMinutes = 120

function Acquire-Lock
{
    param(
        [string]$Path,
        [int]$MaxAgeMinutes
    )

    try
    {
        $dir = Split-Path $Path -Parent
        if (-not (Test-Path $dir))
        {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        if (Test-Path $Path)
        {
            try
            {
                $raw = Get-Content $Path -Raw -ErrorAction SilentlyContinue
                $obj = $raw | ConvertFrom-Json -ErrorAction SilentlyContinue
                $ts  = $null

                if ($obj -and $obj.Timestamp)
                {
                    $ts = [datetime]$obj.Timestamp
                }

                if ($ts)
                {
                    $ageMin = ((Get-Date) - $ts).TotalMinutes
                    if ($ageMin -lt $MaxAgeMinutes)
                    {
                        return $false
                    }

                    Remove-Item $Path -Force -ErrorAction SilentlyContinue
                }
                else
                {
                    $fi = Get-Item $Path -ErrorAction SilentlyContinue
                    if ($fi -and (((Get-Date) - $fi.LastWriteTime).TotalMinutes -ge $MaxAgeMinutes))
                    {
                        Remove-Item $Path -Force -ErrorAction SilentlyContinue
                    }
                    else
                    {
                        return $false
                    }
                }
            }
            catch
            {
                return $false
            }
        }

        $payload = [pscustomobject]@{
            Timestamp = (Get-Date).ToString("o")
            PID       = $PID
            Hostname  = $env:COMPUTERNAME
            User      = "SYSTEM"
            RunStamp  = $RunStamp
            Version   = $ScriptVersion
        }

        $payload | ConvertTo-Json | Set-Content -Path $Path -Force -Encoding UTF8
        return $true
    }
    catch
    {
        Write-InternalLog "Lock warning: $($_.Exception.Message)"
        return $true
    }
}

function Release-Lock
{
    param(
        [string]$Path
    )

    try
    {
        if (Test-Path $Path)
        {
            Remove-Item $Path -Force -ErrorAction SilentlyContinue
        }
    }
    catch
    {
    }
}

if (-not (Acquire-Lock -Path $LockPath -MaxAgeMinutes $LockMaxAgeMinutes))
{
    Write-InternalLog "Another instance appears to be running (lock present). Exiting."
    exit $EXIT_ALREADY_RUNNING_LOCK
}

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
function Escape-Xml
{
    param(
        [string]$txt
    )

    if ([string]::IsNullOrEmpty($txt))
    {
        return ""
    }

    return [System.Security.SecurityElement]::Escape($txt)
}

function Set-PublicReadWrite
{
    param(
        [string]$Path
    )

    if (-not (Test-Path $Path))
    {
        New-Item -Path $Path -ItemType File -Force | Out-Null
    }

    $null = Start-Process -FilePath "icacls.exe" -ArgumentList "`"$Path`" /grant Users:(M) /Q" -PassThru -WindowStyle Hidden -Wait
}

function Set-PublicReadExecute
{
    param(
        [string]$Path
    )

    if (Test-Path $Path)
    {
        $null = Start-Process -FilePath "icacls.exe" -ArgumentList "`"$Path`" /grant Users:(RX) /Q" -PassThru -WindowStyle Hidden -Wait
    }
}

function Ensure-UsersRead
{
    param(
        [string]$Path
    )

    try
    {
        if (Test-Path $Path)
        {
            $null = Start-Process icacls.exe -ArgumentList "`"$Path`" /grant Users:(R) /inheritance:e /Q" -WindowStyle Hidden -Wait
        }
    }
    catch
    {
    }
}

# FIX #12: Use -AsByteStream on PS 6+; -Encoding Byte on PS 5.
# -Encoding Byte was deprecated in PS 6 and removed in PS 7.
function Test-PngSignature
{
    param(
        [string]$Path
    )

    try
    {
        if (-not (Test-Path $Path))
        {
            return $false
        }

        if ($PSVersionTable.PSVersion.Major -ge 6)
        {
            $bytes = Get-Content -Path $Path -AsByteStream -TotalCount 8 -ErrorAction Stop
        }
        else
        {
            $bytes = Get-Content -Path $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop
        }

        return (
            $bytes.Count -eq 8 -and
            $bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47 -and
            $bytes[4] -eq 0x0D -and $bytes[5] -eq 0x0A -and $bytes[6] -eq 0x1A -and $bytes[7] -eq 0x0A
        )
    }
    catch
    {
        return $false
    }
}

# FIX #14: Added Test-LogoTooOld so $LogoMaxAgeDays is actually applied in the popup
# branding check (previously this parameter existed in the popup but was silently ignored).
function Test-LogoTooOld
{
    param(
        [string]$Path,
        [int]$MaxAgeDays
    )

    try
    {
        if ($MaxAgeDays -le 0) { return $false }   # 0 = never refresh based on age
        if (-not (Test-Path $Path)) { return $true }
        $fi = Get-Item $Path -ErrorAction Stop
        return ($fi.LastWriteTime -lt (Get-Date).AddDays(-1 * $MaxAgeDays))
    }
    catch { return $true }
}

function Test-PendingReboot
{
    try
    {
        if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
        {
            return $true
        }
    }
    catch
    {
    }

    try
    {
        if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
        {
            return $true
        }
    }
    catch
    {
    }

    try
    {
        $p = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
        if ($p -and $p.PendingFileRenameOperations)
        {
            return $true
        }
    }
    catch
    {
    }

    try
    {
        if (Test-Path "HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData")
        {
            return $true
        }
    }
    catch
    {
    }

    return $false
}

function Test-InstallInProgress
{
    $hits = @()

    foreach ($pname in @('pnputil.exe', 'expand.exe', 'dism.exe', 'wusa.exe'))
    {
        try
        {
            $base = [IO.Path]::GetFileNameWithoutExtension($pname)
            if (Get-Process -Name $base -ErrorAction SilentlyContinue)
            {
                $hits += "Process:$pname"
            }
        }
        catch
        {
        }
    }

    foreach ($taskName in @('CW-RMM_DriverInstall', 'i3_DriverInstall', 'DriverInstall', 'CW-RMM_DFInstall', 'i3_DriverInstallAtBoot'))
    {
        try
        {
            $t = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($t)
            {
                $ti = Get-ScheduledTaskInfo -TaskName $taskName -ErrorAction SilentlyContinue
                if ($ti -and $ti.State -eq 'Running')
                {
                    $hits += "ScheduledTask:$taskName(Running)"
                }
            }
        }
        catch
        {
        }
    }

    if ($hits.Count -gt 0)
    {
        return $hits
    }

    return $null
}

function Get-InteractiveUserDomainUser
{
    $candidate = $null

    try
    {
        $candidate = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName
    }
    catch
    {
    }

    if (-not $candidate)
    {
        return $null
    }

    if ($candidate -notmatch '\\')
    {
        $candidate = "$env:COMPUTERNAME\$candidate"
    }

    return $candidate
}

function Invoke-PopupTask_Schtasks
{
    param(
        [string]$TaskName,
        [string]$RunAsUser,
        [string]$ScriptPath
    )

    $start = (Get-Date).AddHours(1)
    $st    = $start.ToString('HH:mm')
    $sd    = $start.ToString('MM/dd/yyyy')

    $psCmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $ScriptPath + '"'
    $tr    = '"' + $psCmd + '"'

    $argsCreate = @(
        "/Create", "/F",
        "/TN", $TaskName,
        "/SC", "ONCE",
        "/SD", $sd,
        "/ST", $st,
        "/TR", $tr,
        "/RU", $RunAsUser,
        "/RL", "LIMITED",
        "/IT"
    )

    Write-InternalLog "Creating interactive popup task via schtasks. User=$RunAsUser SD=$sd ST=$st"
    $p1 = Start-Process -FilePath "schtasks.exe" -ArgumentList $argsCreate -Wait -NoNewWindow -PassThru

    if ($p1.ExitCode -ne 0)
    {
        throw "schtasks /Create failed with exit code $($p1.ExitCode)"
    }

    $argsRun = @("/Run", "/TN", $TaskName)
    Write-InternalLog "Starting popup task via schtasks..."
    $p2 = Start-Process -FilePath "schtasks.exe" -ArgumentList $argsRun -Wait -NoNewWindow -PassThru

    if ($p2.ExitCode -ne 0)
    {
        throw "schtasks /Run failed with exit code $($p2.ExitCode)"
    }
}

function Remove-PopupTask_Schtasks
{
    param(
        [string]$TaskName
    )

    try
    {
        Start-Process -FilePath "schtasks.exe" -ArgumentList @("/Delete", "/F", "/TN", $TaskName) -Wait -NoNewWindow | Out-Null
    }
    catch
    {
    }
}

# ---------------------------------------------------------------------------
# BEHAVIOR 1: INSTALL-AT-BOOT ARMING (ONLY WHEN USER CLICKS RESTART)
# ---------------------------------------------------------------------------
$BootTaskName        = 'i3_DriverInstallAtBoot'
$BootMarkerPath      = 'C:\ProgramData\i3\DriverInstallAtBoot.marker.json'
$BootInstallerScript = 'C:\ProgramData\i3\InstallDriversAtBoot.ps1'
$BootInstallerLog    = 'C:\ProgramData\i3\logs\DriverInstallAtBoot.log'

function Write-BootInstallerScript
{
    param(
        [string]$Path,
        [string]$MarkerPath,
        [string]$LogPath,
        [string]$TaskName
    )

    $script = @'
[CmdletBinding()]
param(
    [string]$MarkerPath = "__MARKER__",
    [string]$LogPath    = "__LOG__",
    [string]$TaskName   = "__TASK__"
)

$ErrorActionPreference = "Stop"

function Write-Log
{
    param(
        [string]$Message
    )

    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $Message"
    Write-Host $line
    Add-Content -Path $LogPath -Value $line -ErrorAction SilentlyContinue
}

function Safe-DeleteTask
{
    param(
        [string]$Name
    )

    try
    {
        Start-Process -FilePath "schtasks.exe" -ArgumentList @("/Delete", "/F", "/TN", $Name) -Wait -NoNewWindow | Out-Null
    }
    catch
    {
    }
}

try
{
    $null = New-Item -ItemType Directory -Force -Path (Split-Path $LogPath) -ErrorAction SilentlyContinue

    if (-not (Test-Path $MarkerPath))
    {
        Write-Log "No marker present. Nothing to do. Removing task '$TaskName'."
        Safe-DeleteTask -Name $TaskName
        return
    }

    $markerRaw = Get-Content $MarkerPath -Raw -ErrorAction SilentlyContinue
    $marker    = $markerRaw | ConvertFrom-Json -ErrorAction SilentlyContinue

    if (-not $marker -or -not $marker.Timestamp)
    {
        Write-Log "Marker invalid. Deleting marker and removing task."
        Remove-Item $MarkerPath -Force -ErrorAction SilentlyContinue
        Safe-DeleteTask -Name $TaskName
        return
    }

    $ts       = [datetime]$marker.Timestamp
    $ageHours = ((Get-Date) - $ts).TotalHours

    if ($ageHours -gt 168)
    {
        Write-Log "Marker too old ($([math]::Round($ageHours,2))h). Deleting marker and removing task."
        Remove-Item $MarkerPath -Force -ErrorAction SilentlyContinue
        Safe-DeleteTask -Name $TaskName
        return
    }

    $driverSource = $marker.DriverSource
    if ([string]::IsNullOrWhiteSpace($driverSource))
    {
        $driverSource = "C:\ProgramData\i3\Drivers"
    }

    Write-Log "Marker present. Beginning install. DriverSource=$driverSource"

    if (-not (Test-Path $driverSource))
    {
        throw "DriverSource path not found: $driverSource"
    }

    $installedSomething = $false

    # --- 1. CAB files ---
    # FIX #18: Use DISM /Add-Driver for driver CAB packages.
    # /Add-Package is for Windows packages (cumulative updates, language packs, etc.).
    # /Add-Driver is the correct command for .inf-based driver packages wrapped in a CAB.
    $cabFiles = Get-ChildItem -Path $driverSource -Recurse -Filter "*.cab" -ErrorAction SilentlyContinue
    foreach ($cab in $cabFiles)
    {
        Write-Log "Installing CAB via DISM /Add-Driver: $($cab.Name)"
        $dArgs = @("/Online", "/Add-Driver", "/Driver:`"$($cab.FullName)`"")
        $p = Start-Process -FilePath "dism.exe" -ArgumentList $dArgs -Wait -PassThru -WindowStyle Hidden

        # 0=Success, 3010=Reboot Required, 2/50=Not a driver CAB (will fall through to INF handler)
        if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010)
        {
            Write-Log "  SUCCESS (Exit $($p.ExitCode))"
            $installedSomething = $true
        }
        elseif ($p.ExitCode -eq 2 -or $p.ExitCode -eq 50)
        {
            Write-Log "  SKIPPED — not a recognized driver CAB (Exit $($p.ExitCode)). INF handler will retry if .inf is present."
        }
        else
        {
            Write-Log "  FAILED (Exit $($p.ExitCode))"
        }
    }

    # --- 2. INF files ---
    $infFiles = Get-ChildItem -Path $driverSource -Recurse -Depth 2 -Filter "*.inf" -ErrorAction SilentlyContinue
    foreach ($inf in $infFiles)
    {
        if ($inf.Name -eq 'autorun.inf') { continue }

        Write-Log "Installing INF via pnputil: $($inf.Name)"
        $args = @("/add-driver", "`"$($inf.FullName)`"", "/install")
        $p = Start-Process -FilePath "pnputil.exe" -ArgumentList $args -Wait -PassThru -WindowStyle Hidden

        # 0=Success, 3010=Reboot Required, 259=Already installed
        if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010 -or $p.ExitCode -eq 259)
        {
            Write-Log "  SUCCESS (Exit $($p.ExitCode))"
            $installedSomething = $true
        }
        else
        {
            Write-Log "  FAILED (Exit $($p.ExitCode))"
        }
    }

    # --- 3. MSU files ---
    $msuFiles = Get-ChildItem -Path $driverSource -Recurse -Filter "*.msu" -ErrorAction SilentlyContinue
    foreach ($msu in $msuFiles)
    {
        Write-Log "Installing MSU via wusa: $($msu.Name)"
        $p = Start-Process -FilePath "wusa.exe" -ArgumentList @("`"$($msu.FullName)`"", "/quiet", "/norestart") -Wait -PassThru -WindowStyle Hidden

        if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010)
        {
            Write-Log "  SUCCESS (Exit $($p.ExitCode))"
            $installedSomething = $true
        }
        else
        {
            Write-Log "  FAILED (Exit $($p.ExitCode))"
        }
    }

    if (-not $installedSomething)
    {
        Write-Log "No .msu/.cab/.inf files found or none succeeded."
    }

    Write-Log "Install step complete. Deleting marker and removing task."
    Remove-Item $MarkerPath -Force -ErrorAction SilentlyContinue
    Safe-DeleteTask -Name $TaskName
}
catch
{
    Write-Log "CRITICAL: $($_.Exception.Message)"
    # Marker retained so the task retries on next boot.
}
'@

    $script = $script.Replace("__MARKER__", $MarkerPath.Replace('\', '\\'))
    $script = $script.Replace("__LOG__",    $LogPath.Replace('\', '\\'))
    $script = $script.Replace("__TASK__",   $TaskName)

    $script | Set-Content -Path $Path -Encoding UTF8 -Force
}

function Enable-InstallAtNextBoot_Schtasks
{
    param(
        [string]$TaskName,
        [string]$InstallerScriptPath,
        [string]$MarkerPath,
        [string]$InstallerLogPath,
        [string]$DriverSource
    )

    try
    {
        $null = New-Item -ItemType Directory -Force -Path (Split-Path $InstallerScriptPath) -ErrorAction SilentlyContinue
        $null = New-Item -ItemType Directory -Force -Path (Split-Path $InstallerLogPath) -ErrorAction SilentlyContinue

        $marker = [pscustomobject]@{
            Timestamp    = (Get-Date).ToString("o")
            ComputerName = $env:COMPUTERNAME
            ArmedBy      = "DriverPrompt"
            DriverSource = $DriverSource
            Version      = $ScriptVersion
        }

        $marker | ConvertTo-Json | Set-Content -Path $MarkerPath -Force -Encoding UTF8

        Write-BootInstallerScript -Path $InstallerScriptPath -MarkerPath $MarkerPath -LogPath $InstallerLogPath -TaskName $TaskName

        $psCmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $InstallerScriptPath + '"'
        $tr    = '"' + $psCmd + '"'

        try
        {
            Start-Process -FilePath "schtasks.exe" -ArgumentList @("/Delete", "/F", "/TN", $TaskName) -Wait -NoNewWindow | Out-Null
        }
        catch
        {
        }

        $argsCreate = @(
            "/Create", "/F",
            "/TN", $TaskName,
            "/SC", "ONSTART",
            "/RU", "SYSTEM",
            "/RL", "HIGHEST",
            "/TR", $tr
        )

        Write-InternalLog "Arming install-at-boot. Task='$TaskName' Marker='$MarkerPath' Script='$InstallerScriptPath'"
        $p1 = Start-Process -FilePath "schtasks.exe" -ArgumentList $argsCreate -Wait -NoNewWindow -PassThru

        if ($p1.ExitCode -ne 0)
        {
            throw "schtasks /Create (ONSTART) failed with exit code $($p1.ExitCode)"
        }

        Ensure-UsersRead -Path (Split-Path $InstallerScriptPath -Parent)
        Ensure-UsersRead -Path (Split-Path $InstallerLogPath -Parent)
        Ensure-UsersRead -Path $InstallerScriptPath
    }
    catch
    {
        Write-InternalLog "Enable-InstallAtNextBoot error: $($_.Exception.Message)"
        throw
    }
}

# ---------------------------------------------------------------------------
# RUN SUMMARY JSON
# ---------------------------------------------------------------------------
$RunSummaryPath = Join-Path $LogFolder "DriverPrompt_RunSummary_${RunStamp}.json"

# FIX #3: Notes are now structured {Severity, Message} objects.
# Only 'Error'-severity notes surface as RMM alerts in DFReport.
# 'Info' and 'Warning' notes are preserved in the JSON for diagnostics but do NOT trigger alerts.
$Summary = [ordered]@{
    Version         = $ScriptVersion
    RunStamp        = $RunStamp
    StartTime       = $RunStart.ToString("o")
    EndTime         = $null
    ComputerName    = $env:COMPUTERNAME
    TriggerLog      = $null
    TriggerFound    = $false
    InteractiveUser = $null
    UserChoice      = $null
    ExitCode        = $null
    Notes           = [System.Collections.Generic.List[pscustomobject]]::new()
}

# FIX #3: Helper adds a structured note and logs it internally.
# Severity 'Error' → surfaces as RMM alert in DFReport.
# Severity 'Info' / 'Warning' → preserved in JSON only; not flagged as errors.
function Add-SummaryNote
{
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Severity = 'Info'
    )

    $Summary.Notes.Add([pscustomobject]@{ Severity = $Severity; Message = $Message })
    Write-InternalLog "Note[$Severity]: $Message"
}

function Write-RunSummary
{
    param(
        [int]$ExitCode
    )

    try
    {
        if (-not (Test-Path $LogFolder))
        {
            New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
        }

        $Summary.EndTime  = (Get-Date).ToString("o")
        $Summary.ExitCode = $ExitCode
        $Summary | ConvertTo-Json -Depth 6 | Out-File -FilePath $RunSummaryPath -Force -Encoding UTF8
    }
    catch
    {
        Write-InternalLog "Run summary write warning: $($_.Exception.Message)"
    }
}

# FIX #5: Auto-prune old RunSummary JSON files — retain the 10 most recent.
# Previously these accumulated indefinitely; DFReport already selects only the latest,
# but the folder grew unbounded over weeks of daily/weekly runs.
function Invoke-RunSummaryCleanup
{
    try
    {
        $old = Get-ChildItem $LogFolder -Filter 'DriverPrompt_RunSummary_*.json' -ErrorAction SilentlyContinue |
               Sort-Object LastWriteTime -Descending |
               Select-Object -Skip 10

        $old | ForEach-Object { Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue }

        if ($old -and $old.Count -gt 0)
        {
            Write-InternalLog "Pruned $($old.Count) old RunSummary file(s)."
        }
    }
    catch
    {
        Write-InternalLog "RunSummary cleanup warning: $($_.Exception.Message)"
    }
}

Invoke-RunSummaryCleanup

# ---------------------------------------------------------------------------
# TRIGGER CHECK
# ---------------------------------------------------------------------------
$DateStamp = Get-Date -Format "yyyyMMdd"
$TargetLog = Join-Path -Path $LogFolder -ChildPath "DriverDownload_$DateStamp.log"
$Summary.TriggerLog = $TargetLog

$TriggerPatterns = @(
    'TRIGGER=UPDATES_READY',
    'Updates\s+are\s+ready\s+to\s+install',
    'Updates\s+ready\s+in'
)

$TriggerHit            = $null
$TriggerMatchedPattern = $null

if (Test-Path $TargetLog)
{
    foreach ($pat in $TriggerPatterns)
    {
        $hit = Select-String -Path $TargetLog -Pattern $pat -Encoding UTF8 -ErrorAction SilentlyContinue | Select-Object -Last 1
        if ($hit)
        {
            $TriggerHit            = $hit
            $TriggerMatchedPattern = $pat
            break
        }
    }
}

if (-not $TriggerHit)
{
    Write-Host "DONE | No updates found in today's log."
    Write-InternalLog "No trigger in today's log. Exiting."
    Write-RunSummary -ExitCode $EXIT_NOACTION_OR_DONE
    Release-Lock -Path $LockPath
    exit $EXIT_NOACTION_OR_DONE
}

$Summary.TriggerFound = $true
Write-InternalLog "Trigger Found. Pattern='$TriggerMatchedPattern' Line='$($TriggerHit.Line.Trim())'"

# ---------------------------------------------------------------------------
# INSTALL-IN-PROGRESS GUARD
# ---------------------------------------------------------------------------
$inProgressHits = Test-InstallInProgress
if ($inProgressHits)
{
    Write-InternalLog "Install already in progress. Bailing. Hits: $($inProgressHits -join '; ')"
    Write-RunSummary -ExitCode $EXIT_INSTALL_IN_PROGRESS
    Release-Lock -Path $LockPath
    exit $EXIT_INSTALL_IN_PROGRESS
}

# ---------------------------------------------------------------------------
# PENDING REBOOT SKIP
# ---------------------------------------------------------------------------
$PendingRebootNow = Test-PendingReboot
if ($PendingRebootNow)
{
    Write-Host "DONE | Pending reboot already detected. Skipping popup."
    Write-InternalLog "Pending reboot detected. Skipping popup."

    # FIX #3: Warning-severity — does not surface as an RMM error
    Add-SummaryNote "Skipped popup: Windows pending reboot detected." -Severity 'Warning'
    Write-RunSummary -ExitCode $EXIT_PENDING_REBOOT_SKIP

    Release-Lock -Path $LockPath
    exit $EXIT_PENDING_REBOOT_SKIP
}

# ---------------------------------------------------------------------------
# INTERACTIVE USER
# ---------------------------------------------------------------------------
$domainUser = Get-InteractiveUserDomainUser
if (-not $domainUser)
{
    Write-InternalLog "No interactive user. Cannot show popup. Exiting."
    Write-RunSummary -ExitCode $EXIT_NO_INTERACTIVE_USER
    Release-Lock -Path $LockPath
    exit $EXIT_NO_INTERACTIVE_USER
}

$Summary.InteractiveUser = $domainUser
Write-InternalLog "Interactive user resolved: $domainUser"
Write-Host "RESTART (OPTIONAL) | Trigger Found."

# FIX #8: Establish StatusFile permissions ONCE here, immediately after confirming an
# interactive user exists. The content is reset to "Waiting" just before the task
# launches (see Execution section) — not here — so a stale "Restart" or "Later" value
# from a prior run cannot break the wait loop before the new task even starts.
Set-PublicReadWrite -Path $StatusFile

# ---------------------------------------------------------------------------
# BRANDING (cache / validate)
# ---------------------------------------------------------------------------
$BrandDir   = 'C:\ProgramData\i3\branding'
$BrandLogo  = Join-Path $BrandDir 'logo.png'
$LastGood   = Join-Path $BrandDir 'logo.lastgood.png'
$ScriptPath = 'C:\ProgramData\i3\ShowPopup.ps1'

function Test-LogoOk
{
    param(
        [string]$Path
    )

    try
    {
        if (-not (Test-Path $Path))
        {
            return $false
        }

        $fi = Get-Item $Path -ErrorAction Stop
        if ($fi.Length -lt 2048)
        {
            return $false
        }

        if (-not (Test-PngSignature -Path $Path))
        {
            return $false
        }

        return $true
    }
    catch
    {
        return $false
    }
}

try
{
    if (-not (Test-Path $BrandDir))
    {
        New-Item -ItemType Directory -Path $BrandDir -Force | Out-Null
    }

    Ensure-UsersRead -Path $BrandDir

    foreach ($p in @($BrandLogo, $LastGood))
    {
        if (Test-Path $p)
        {
            try
            {
                Set-ItemProperty -Path $p -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
            }
            catch
            {
            }
        }
    }

    $needDownload = $true

    if (Test-LogoOk -Path $BrandLogo)
    {
        # FIX #14: $LogoMaxAgeDays is now applied in the popup branding check.
        # Previously this parameter was accepted but silently ignored here —
        # the age check only existed in the scanner's Update-BrandingCache function.
        if (Test-LogoTooOld -Path $BrandLogo -MaxAgeDays $LogoMaxAgeDays)
        {
            Write-InternalLog "Branding: Logo is older than $LogoMaxAgeDays days; will refresh."
        }
        else
        {
            $needDownload = $false
            Write-InternalLog "Branding: Using cached logo.png (valid)."
        }
    }

    if ($needDownload -and (Test-LogoOk -Path $LastGood))
    {
        Copy-Item $LastGood $BrandLogo -Force -ErrorAction SilentlyContinue
        if (Test-LogoOk -Path $BrandLogo)
        {
            $needDownload = $false
            Write-InternalLog "Branding: Restored logo.png from lastgood."
        }
    }

    if ($needDownload)
    {
        Write-InternalLog "Branding: Downloading logo from $BrandLogoUrl..."
        Invoke-WebRequest -Uri $BrandLogoUrl -OutFile $BrandLogo -UseBasicParsing -ErrorAction Stop
        Unblock-File -Path $BrandLogo -ErrorAction SilentlyContinue

        if (-not (Test-LogoOk -Path $BrandLogo))
        {
            Remove-Item $BrandLogo -Force -ErrorAction SilentlyContinue
            throw "Downloaded logo failed validation"
        }

        Copy-Item $BrandLogo $LastGood -Force -ErrorAction SilentlyContinue
        Write-InternalLog "Branding: Downloaded logo validated and cached as lastgood."
    }

    Ensure-UsersRead -Path $BrandLogo

    foreach ($p in @($BrandLogo, $LastGood))
    {
        if (Test-Path $p)
        {
            try
            {
                Set-ItemProperty -Path $p -Name IsReadOnly -Value $true -ErrorAction SilentlyContinue
            }
            catch
            {
            }
        }
    }

    $li = Get-Item $BrandLogo -ErrorAction SilentlyContinue
    if ($li)
    {
        Write-InternalLog "Branding: Logo ready. Size=$($li.Length) LastWrite=$($li.LastWriteTime)"
    }
}
catch
{
    Write-InternalLog "Branding Error: $($_.Exception.Message) Popup will show text only."
    $BrandLogo = "C:\Null\Ignored.png"
}

Start-Sleep -Milliseconds 250

# ---------------------------------------------------------------------------
# USER POPUP SCRIPT
#
# FIX #10: XAML TextBlocks now use pre-computed XML-safe tokens:
#   __XML_CAPTION__  — XML-escaped single-line caption (safe as a XAML attribute value)
#   __XAML_BODY__    — XML-escaped body lines joined with <LineBreak/> inline elements,
#                      embedded as TextBlock element content (not an attribute).
#
# Why element content instead of attribute?
#   XML normalizes whitespace in attribute values — newlines become spaces. Putting body
#   text in a Text="..." attribute would collapse all line breaks to a single line. Using
#   <LineBreak/> inline elements inside the TextBlock element content avoids this entirely
#   and renders correct line breaks in the WPF popup.
#
# Why pre-compute in SYSTEM context instead of PS interpolation in the generated script?
#   The old approach mixed PS variable interpolation ($($caption -replace ...)) inside a
#   double-quoted XAML here-string. If $Caption or $Body contained $, backticks, or quotes,
#   the inner here-string could break silently. Pre-escaping and embedding as tokens removes
#   all runtime injection risk.
# ---------------------------------------------------------------------------
$userScript = @'
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase

$statusFile = "__STATUSFILE__"
$logoPath   = "__LOGO_PATH__"

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        ShowInTaskbar="False"
        Topmost="True"
        Width="560" Height="360">
  <Grid>
    <Border CornerRadius="16" Background="#F9FAFB" Padding="24">
      <Border.Effect>
        <DropShadowEffect BlurRadius="18" ShadowDepth="0" Opacity="0.35"/>
      </Border.Effect>
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" HorizontalAlignment="Center">
          <Image x:Name="LogoImg" Width="160" Height="120"
                 HorizontalAlignment="Center" Margin="0,8,0,12"
                 Stretch="Uniform" Visibility="Collapsed"/>
        </StackPanel>

        <StackPanel Grid.Row="1" VerticalAlignment="Center" Margin="8,12,8,12">
          <TextBlock Text="__XML_CAPTION__"
                     FontFamily="Segoe UI Variable" FontWeight="SemiBold"
                     FontSize="20" TextAlignment="Center" Margin="0,0,0,12"/>
          <TextBlock FontFamily="Segoe UI Variable" FontSize="14"
                     TextWrapping="Wrap" TextAlignment="Center">__XAML_BODY__</TextBlock>
        </StackPanel>

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,16,0,0">
          <Button x:Name="RestartBtn" Content="Restart now" MinWidth="132" Padding="12,8"
                  FontFamily="Segoe UI Variable" FontWeight="SemiBold" Margin="6,0">
            <Button.Background>
              <SolidColorBrush Color="{DynamicResource {x:Static SystemColors.HotTrackColorKey}}"/>
            </Button.Background>
            <Button.Foreground><SolidColorBrush Color="White"/></Button.Foreground>
          </Button>
          <Button x:Name="LaterBtn" Content="Later" MinWidth="160" Padding="12,8"
                  FontFamily="Segoe UI Variable" Margin="6,0"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$restartBtn = $window.FindName('RestartBtn')
$laterBtn   = $window.FindName('LaterBtn')
$logoImg    = $window.FindName('LogoImg')

try
{
    if ([System.IO.File]::Exists($logoPath))
    {
        $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit()
        $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmp.UriSource   = New-Object System.Uri($logoPath)
        $bmp.EndInit()
        $bmp.Freeze()
        $logoImg.Source     = $bmp
        $logoImg.Visibility = 'Visible'
    }
}
catch
{
}

$restartBtn.Add_Click({
    $obj = [pscustomobject]@{ Status = "Restart"; Timestamp = (Get-Date).ToString() }
    $obj | ConvertTo-Json | Set-Content -Path $statusFile -Force
    $window.Close()
})

$laterBtn.Add_Click({
    $obj = [pscustomobject]@{ Status = "Later"; Timestamp = (Get-Date).ToString() }
    $obj | ConvertTo-Json | Set-Content -Path $statusFile -Force
    $window.Close()
})

$null = $window.ShowDialog()
'@

# FIX #10: Pre-compute XML-safe replacements in SYSTEM context before writing ShowPopup.ps1.
#
# Caption: XML-escaped string — safe to embed as a XAML Text attribute value.
# Body: each line is XML-escaped individually, then joined with <LineBreak/> elements.
#       This renders as real line breaks in the WPF TextBlock element content.
#       (XML normalizes newlines in attribute values to spaces — element content does not.)
$xmlCaption = Escape-Xml $Caption
$bodyLines  = $Body.Trim() -split "`r?`n"
$xamlBody   = ($bodyLines | ForEach-Object { Escape-Xml $_.Trim() }) -join '<LineBreak/>'

$userScript = $userScript.Replace('__XML_CAPTION__', $xmlCaption)
$userScript = $userScript.Replace('__XAML_BODY__',   $xamlBody)
$userScript = $userScript.Replace('__STATUSFILE__',  $StatusFile.Replace('\', '\\'))
$userScript = $userScript.Replace('__LOGO_PATH__',   $BrandLogo.Replace('\', '\\'))

$userScript | Set-Content -Path $ScriptPath -Encoding UTF8 -Force
Set-PublicReadExecute -Path $ScriptPath

# ---------------------------------------------------------------------------
# EXECUTION & WAIT LOOP
# ---------------------------------------------------------------------------
$TaskName        = "CW-RMM_ShowPopup"
$timedOut        = $false
$nextExit        = $EXIT_NOACTION_OR_DONE
$Summary.UserChoice = "None"
$shouldRebootNow = $false

try
{
    Remove-PopupTask_Schtasks -TaskName $TaskName

    # FIX #8: Reset StatusFile to "Waiting" immediately before launching the task —
    # not at the top of the script. This eliminates the race condition where a leftover
    # "Restart" or "Later" from a prior run would break the wait loop on the very first poll.
    # Permissions were already established via Set-PublicReadWrite above.
    $statusObj = [pscustomobject]@{
        Status    = "Waiting"
        Timestamp = (Get-Date).ToString("o")
    }
    $statusObj | ConvertTo-Json | Set-Content -Path $StatusFile -Force -Encoding UTF8

    Invoke-PopupTask_Schtasks -TaskName $TaskName -RunAsUser $domainUser -ScriptPath $ScriptPath
    Write-InternalLog "Popup launched. Entering Wait Loop..."

    $StartTime  = Get-Date
    $MaxSeconds = $MaxWaitTimeMinutes * 60

    while ($true)
    {
        Start-Sleep -Seconds 5

        $elapsed = ((Get-Date) - $StartTime).TotalSeconds
        if ($elapsed -gt $MaxSeconds)
        {
            Write-InternalLog "Timeout reached. Assuming 'Later'."
            $timedOut           = $true
            $Summary.UserChoice = "Timeout"
            break
        }

        try
        {
            if (Test-Path $StatusFile)
            {
                $data = Get-Content $StatusFile -Raw -ErrorAction Stop | ConvertFrom-Json

                if ($data.Status -eq "Restart")
                {
                    Write-InternalLog "User selected RESTART."
                    $Summary.UserChoice = "Restart"
                    $nextExit           = $EXIT_USER_CHOSE_RESTART_START

                    Enable-InstallAtNextBoot_Schtasks `
                        -TaskName            $BootTaskName `
                        -InstallerScriptPath $BootInstallerScript `
                        -MarkerPath          $BootMarkerPath `
                        -InstallerLogPath    $BootInstallerLog `
                        -DriverSource        $DriverSource

                    # FIX #3: Info-severity — this is a success event, not an error.
                    # DFReport will NOT surface this as an RMM alert.
                    Add-SummaryNote "Armed install-at-boot task '$BootTaskName' (Behavior 1)." -Severity 'Info'
                    $shouldRebootNow = $true

                    break
                }
                elseif ($data.Status -eq "Later")
                {
                    Write-InternalLog "User selected LATER."
                    $Summary.UserChoice = "Later"
                    $nextExit           = $EXIT_NOACTION_OR_DONE
                    break
                }
            }
        }
        catch
        {
            # ignore locked / partial reads
        }
    }
}
catch
{
    Write-InternalLog "CRITICAL ERROR: $($_.Exception.Message)"
    # FIX #3: Error-severity — this WILL surface as an RMM alert in DFReport.
    Add-SummaryNote "Critical error: $($_.Exception.Message)" -Severity 'Error'
    $nextExit = 1
}
finally
{
    Write-InternalLog "Cleaning up Task..."
    Remove-PopupTask_Schtasks -TaskName $TaskName
}

if ($timedOut -and $nextExit -eq $EXIT_NOACTION_OR_DONE)
{
    $nextExit = $EXIT_PROMPT_TIMEOUT_ASSUMED
}

Write-InternalLog "Session Ended. ExitCode=$nextExit"
Write-RunSummary -ExitCode $nextExit
Release-Lock -Path $LockPath

# If user chose restart and we armed install-at-boot, reboot now (optional)
if ($shouldRebootNow -and $RebootOnRestartChoice)
{
    try
    {
        # FIX #1: Write-RunSummary and Release-Lock have already completed above.
        # Start-Sleep -Seconds 3 gives the filesystem a brief flush grace period before
        # the OS begins shutdown teardown. Without this, /t 0 /f can interrupt pending
        # I/O on a fast SSD before the summary file is fully committed to disk.
        Write-InternalLog "Initiating reboot now (user-approved). Brief pause for filesystem flush..."
        Start-Sleep -Seconds 3
        Start-Process -FilePath 'shutdown.exe' -ArgumentList @('/r', '/t', '0', '/f') -WindowStyle Hidden
    }
    catch
    {
        Write-InternalLog "Reboot attempt failed: $($_.Exception.Message)"
    }
}

exit $nextExit
