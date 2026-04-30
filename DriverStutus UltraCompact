<#
.SYNOPSIS
    Driver Status - Ultra-Compact Output (fits in ~70 chars)

.DESCRIPTION
    Outputs driver count and abbreviated names to fit tight space.
    Format: "4 drivers: Intel Serial IO, Realtek Audio, Intel ME"
    
.PARAMETER DriverPath
    Path where drivers are stored. Default: C:\ProgramData\i3\Drivers

.PARAMETER MaxLength
    Maximum total characters. Default: 70
#>
param([string]$p = "C:\ProgramData\i3\Drivers", [int]$max = 70)

if (-not (Test-Path $p)) { "0"; exit 0 }

$d = @(); Get-ChildItem $p -Directory -EA 0 | ForEach-Object {
    if (@(Get-ChildItem $_.FullName -File -EA 0).Count -gt 0) { $d += $_.Name }
}

if ($d.Count -eq 0) { "0"; exit 0 }

# Ultra-short format
$out = "$($d.Count): "
$names = ($d | ForEach-Object { 
    $n = $_ -replace '_', ' ' -replace 'Controller', 'C' -replace 'Interface', 'I' -replace 'High Definition', 'HD'
    ($n -split ' ')[0..1] -join ' '
}) -join ', '

$available = $max - $out.Length
if ($names.Length -gt $available) {
    $names = $names.Substring(0, $available - 3) + "..."
}

$out + $names
