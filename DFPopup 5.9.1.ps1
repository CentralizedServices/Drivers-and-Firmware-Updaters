<#
.SYNOPSIS
    Analyzes Driver Logs and prompts the user to restart if updates are ready.

.DESCRIPTION
    Runs as SYSTEM/Admin via RMM.
    - Detects "updates ready" in today's driver download log
    - Launches a user-context WPF popup via scheduled task
    - Waits on a JSON status file for user response

    Version History:
        5.8 - FIX: ScheduledTask LogonType enum compatibility
        5.9 - FIX: Avoid ScheduledTasks module SID translation errors
              Uses schtasks.exe /IT with DOMAIN\user instead of Register-ScheduledTask

        5.9.1 - FIX: schtasks /TR quoting so PowerShell args (e.g., -WindowStyle) are not parsed as schtasks args
#>

[CmdletBinding()]
param(
    [string]$LogFolder = 'C:\ProgramData\i3\logs',
    [string]$DriverSource = 'C:\ProgramData\i3\Drivers',
    [string]$BrandLogoUrl = 'https://s3.us-east-1.wasabisys.com/i3-blob-storage/Installers/i3_CharmOnly_400x300.png',
    [int]$MaxWaitTimeMinutes = 55,
    [int]$MaxSnoozes = 5,
    [string]$Caption = "i3 Updates Available",
    [string]$Body = @"
New driver and firmware updates are ready to install.
Please save your work.
When you click 'Restart Now', updates will install and the PC will reboot automatically.
"@,
    [bool]$EnableDriverDiff = $true,
    [int]$LogoMaxAgeDays = 0
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

$ScriptVersion = "5.9.1"
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

    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] $Message"
    Write-Host $line
    Add-Content -Path $InternalLogPath -Value $line -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------------
# SINGLE-INSTANCE LOCK
# ---------------------------------------------------------------------------
$LockPath = 'C:\ProgramData\i3\DriverPrompt.lock'
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

        $bytes = Get-Content -Path $Path -Encoding Byte -TotalCount 8 -ErrorAction Stop

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

    foreach ($pname in @('pnputil.exe', 'expand.exe'))
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

    foreach ($taskName in @('CW-RMM_DriverInstall', 'i3_DriverInstall', 'DriverInstall', 'CW-RMM_DFInstall'))
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
        [string]$RunAsUser,     # DOMAIN\user or COMPUTERNAME\user
        [string]$ScriptPath
    )

    # schtasks has minute granularity; schedule for next minute
    $start = (Get-Date).AddMinutes(1)
    $st    = $start.ToString('HH:mm')
    $sd    = $start.ToString('MM/dd/yyyy')

    # IMPORTANT:
    # schtasks.exe requires /TR to be ONE quoted string. If it's not, flags like
    # -WindowStyle get parsed as schtasks arguments and you get:
    #   ERROR: Invalid argument/option - '-WindowStyle'
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
# RUN SUMMARY JSON (kept lightweight)
# ---------------------------------------------------------------------------
$RunSummaryPath = Join-Path $LogFolder "DriverPrompt_RunSummary_${RunStamp}.json"
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
    Notes           = @()
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

$TriggerHit = $null
$TriggerMatchedPattern = $null

if (Test-Path $TargetLog)
{
    foreach ($pat in $TriggerPatterns)
    {
        $hit = Select-String -Path $TargetLog -Pattern $pat -Encoding UTF8 -ErrorAction SilentlyContinue | Select-Object -Last 1
        if ($hit)
        {
            $TriggerHit = $hit
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
# Determine interactive user (DOMAIN\user)
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

# ---------------------------------------------------------------------------
# Pending reboot UX
# ---------------------------------------------------------------------------
$PendingRebootNow = Test-PendingReboot
if ($PendingRebootNow)
{
    $Caption = "i3 Restart Required"
    $Body = @"
A restart is already required to finish applying updates on this PC.
Please save your work.
Click 'Restart now' to restart and complete the update process.
"@
}

Write-Host "RESTART (OPTIONAL) | Trigger Found."

# ---------------------------------------------------------------------------
# Reset Status File
# ---------------------------------------------------------------------------
$statusObj = [pscustomobject]@{
    Status    = "Waiting"
    Timestamp = (Get-Date).ToString()
}

$statusObj | ConvertTo-Json | Set-Content -Path $StatusFile -Force -Encoding UTF8
Set-PublicReadWrite -Path $StatusFile

# ---------------------------------------------------------------------------
# Branding (cache/validate)
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
        $needDownload = $false
        Write-InternalLog "Branding: Using cached logo.png (valid)."
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
# User popup script
# ---------------------------------------------------------------------------
$userScript = @'
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase

$caption = "__CAPTION__"
$messageTemplate = @"
__BODY__
"@
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
          <TextBlock Text="$($caption -replace '"','"')"
                     FontFamily="Segoe UI Variable" FontWeight="SemiBold"
                     FontSize="20" TextAlignment="Center" Margin="0,0,0,12"/>
          <TextBlock Text="$($messageTemplate -replace '"','"')"
                     FontFamily="Segoe UI Variable" FontSize="14"
                     TextWrapping="Wrap" TextAlignment="Center"/>
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
        $bmp.UriSource = New-Object System.Uri($logoPath)
        $bmp.EndInit()
        $bmp.Freeze()
        $logoImg.Source = $bmp
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

$SafeCaption = Escape-Xml $Caption
$SafeBody    = Escape-Xml $Body

$userScript = $userScript.Replace('__CAPTION__',    $SafeCaption.Replace('"', '""'))
$userScript = $userScript.Replace('__BODY__',       $SafeBody)
$userScript = $userScript.Replace('__STATUSFILE__', $StatusFile.Replace('\', '\\'))
$userScript = $userScript.Replace('__LOGO_PATH__',  $BrandLogo.Replace('\', '\\'))

$userScript | Set-Content -Path $ScriptPath -Encoding UTF8 -Force
Set-PublicReadExecute -Path $ScriptPath

# ---------------------------------------------------------------------------
# EXECUTION & WAIT LOOP
# ---------------------------------------------------------------------------
$TaskName = "CW-RMM_ShowPopup"
$timedOut = $false
$nextExit = $EXIT_NOACTION_OR_DONE
$Summary.UserChoice = "None"

try
{
    Remove-PopupTask_Schtasks -TaskName $TaskName

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
            $timedOut = $true
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
                    $nextExit = $EXIT_USER_CHOSE_RESTART_START
                    break
                }
                elseif ($data.Status -eq "Later")
                {
                    Write-InternalLog "User selected LATER."
                    $Summary.UserChoice = "Later"
                    $nextExit = $EXIT_NOACTION_OR_DONE
                    break
                }
            }
        }
        catch
        {
            # ignore locked/partial reads
        }
    }
}
catch
{
    Write-InternalLog "CRITICAL ERROR: $($_.Exception.Message)"
    $Summary.Notes += "Critical error: $($_.Exception.Message)"
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
exit $nextExit
