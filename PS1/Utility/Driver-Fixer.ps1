#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$SkipClean,
    [switch]$SkipRepair,
    [switch]$SkipUpdate,
    [string]$LogPath = "$env:USERPROFILE\Desktop\DriverFixer-$(Get-Date -Format 'yyyyMMdd-HHmm').log"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# Helpers
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level.PadRight(5), $Message
    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red    }
        'OK'    { Write-Host $line -ForegroundColor Green  }
        'HEAD'  { Write-Host "`n$line" -ForegroundColor Cyan }
        default { Write-Host $line }
    }
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}

function Invoke-Privileged {
    # Confirm the session is elevated; abort if not.
    $id  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pri = New-Object Security.Principal.WindowsPrincipal($id)
    if (-not $pri.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Log 'Script must be run as Administrator.' 'ERROR'
        exit 1
    }
}

# Pre-flight

Invoke-Privileged

$null = New-Item -ItemType File -Path $LogPath -Force
Write-Log '  Driver Fixer  -  starting run'               'HEAD'
Write-Log "  Log: $LogPath"                             'HEAD'

# Audit

Write-Log 'PHASE 1  -  Auditing devices' 'HEAD'

# All present devices with an error code
$errorDevices = @(Get-PnpDevice -PresentOnly:$false |
    Where-Object { $_.Problem -ne 'CM_PROB_NOT_CONFIGURED' -and $_.Status -eq 'Error' })

# All non-present (phantom) devices
$phantomDevices = @(Get-PnpDevice -PresentOnly:$false |
    Where-Object { $_.Present -eq $false })

# All staged driver packages
$allDrivers = & pnputil /enum-drivers 2>&1 |
    Select-String 'Published Name|Original Name|Driver Version|Signer Name|Class Name|Driver Date' |
    ForEach-Object { $_.Line.Trim() }

Write-Log "Devices with errors  : $($errorDevices.Count)"
Write-Log "Phantom devices      : $($phantomDevices.Count)"

if ($errorDevices.Count -eq 0 -and $phantomDevices.Count -eq 0) {
    Write-Log 'No broken or phantom devices found.' 'OK'
} else {
    Write-Log '--- Broken Devices ---'
    foreach ($dev in $errorDevices) {
        Write-Log ("  [{0}] {1,-45} Problem: {2}" -f $dev.Class, $dev.FriendlyName, $dev.Problem) 'WARN'
        Write-Log ("    InstanceId: {0}" -f $dev.InstanceId)

        # Show which INF is bound
        try {
            $infProp = Get-PnpDeviceProperty -InstanceId $dev.InstanceId `
                           -KeyName 'DEVPKEY_Device_DriverInfPath' -ErrorAction Stop
            Write-Log ("    Bound INF : {0}" -f $infProp.Data)
        } catch {
            Write-Log '    Bound INF : (could not determine)' 'WARN'
        }
    }

    Write-Log '--- Phantom Devices ---'
    foreach ($dev in $phantomDevices) {
        Write-Log ("  [{0}] {1,-45} Status: {2}" -f $dev.Class, $dev.FriendlyName, $dev.Status)
    }
}

# Identify stale / duplicate INF packages (same Original Name, multiple Published Names)
Write-Log '--- Staged Driver Packages ---'
$rawDrivers = & pnputil /enum-drivers 2>&1
$driverBlocks = @()
$current = [ordered]@{}
foreach ($line in $rawDrivers) {
    if ($line -match '^\s*$') {
        if ($current.Count -gt 0) { $driverBlocks += [PSCustomObject]$current }
        $current = [ordered]@{}
        continue
    }
    if ($line -match '^\s*([\w\s]+?)\s*:\s*(.+)$') {
        $current[$Matches[1].Trim()] = $Matches[2].Trim()
    }
}
if ($current.Count -gt 0) { $driverBlocks += [PSCustomObject]$current }

# Group by Original Name to spot duplicates
$grouped = $driverBlocks | Where-Object { $_.'Original Name' } |
    Group-Object 'Original Name' | Where-Object { $_.Count -gt 1 }

if ($grouped) {
    Write-Log 'Duplicate driver packages detected (same .inf, multiple oem##.inf entries):' 'WARN'
    foreach ($g in $grouped) {
        Write-Log ("  Original: {0}  ({1} copies)" -f $g.Name, $g.Count) 'WARN'
        foreach ($d in $g.Group) {
            Write-Log ("    Published: {0}  Version: {1}" -f $d.'Published Name', $d.'Driver Version')
        }
    }
} else {
    Write-Log 'No duplicate staged driver packages found.' 'OK'
}

# Clean
if (-not $SkipClean) {
    Write-Log 'PHASE 2  -  Cleaning phantom devices and stale drivers' 'HEAD'

    # Remove phantom devices
    foreach ($dev in $phantomDevices) {
        if ($PSCmdlet.ShouldProcess($dev.InstanceId, 'Remove phantom device')) {
            Write-Log ("Removing phantom: {0} [{1}]" -f $dev.FriendlyName, $dev.InstanceId)
            try {
                & pnputil /remove-device $dev.InstanceId 2>&1 | ForEach-Object { Write-Log "  $_" }
            } catch {
                Write-Log "  Failed to remove: $_" 'ERROR'
            }
        }
    }

    # Remove older duplicate packages  -  keep highest version
    foreach ($g in $grouped) {
        $sorted = $g.Group | Sort-Object {
            $parts = ($_.'Driver Version' -replace '[^\d\.]','').Split('.') | ForEach-Object { [int64]$_ }
            # Pad to 4 parts so all tuples are the same length
            while ($parts.Count -lt 4) { $parts += 0 }
            [tuple]::Create($parts[0], $parts[1], $parts[2], $parts[3])
        } -Descending
        $keep   = $sorted | Select-Object -First 1
        $remove = $sorted | Select-Object -Skip 1
        Write-Log ("Keeping newest '{0}' ({1})" -f $keep.'Published Name', $keep.'Driver Version') 'OK'
        foreach ($old in $remove) {
            $pub = $old.'Published Name'
            if ($PSCmdlet.ShouldProcess($pub, 'Delete stale driver package')) {
                Write-Log ("Deleting stale driver: {0}" -f $pub) 'WARN'
                $result = & pnputil /delete-driver $pub /uninstall 2>&1
                $result | ForEach-Object { Write-Log "  $_" }
                if ($LASTEXITCODE -ne 0) {
                    Write-Log "  pnputil returned $LASTEXITCODE  -  package may still be in use." 'WARN'
                }
            }
        }
    }

    Write-Log 'Clean phase complete.' 'OK'
} else {
    Write-Log 'Skipping Phase 2 (SkipClean specified).'
}

# Repair 
if (-not $SkipRepair) {
    Write-Log 'PHASE 3  -  Repairing broken devices' 'HEAD'

    foreach ($dev in $errorDevices) {
        $id   = $dev.InstanceId
        $name = $dev.FriendlyName

        Write-Log ("Attempting repair for: {0}" -f $name)

        # Step A  -  disable / enable cycle
        Write-Log "  Step A: disable/enable cycle"
        try {
            Disable-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop
            Start-Sleep -Milliseconds 1500
            Enable-PnpDevice  -InstanceId $id -Confirm:$false -ErrorAction Stop
            Start-Sleep -Milliseconds 2000

            $after = Get-PnpDevice -InstanceId $id
            if ($after.Status -eq 'OK') {
                Write-Log "  Disable/enable resolved the issue." 'OK'
                continue
            }
            Write-Log ("  Still erroring after cycle: {0}" -f $after.Problem) 'WARN'
        } catch {
            Write-Log ("  Disable/enable failed: {0}" -f $_) 'WARN'
        }

        # Step B  -  remove device and rescan
        Write-Log "  Step B: remove device and rescan"
        if ($PSCmdlet.ShouldProcess($id, 'Remove device and rescan')) {
            try {
                & pnputil /remove-device $id 2>&1 | ForEach-Object { Write-Log "  $_" }
                & pnputil /scan-devices   2>&1 | ForEach-Object { Write-Log "  $_" }
                Start-Sleep -Seconds 3

                $after2 = Get-PnpDevice -InstanceId $id -ErrorAction SilentlyContinue
                if ($after2 -and $after2.Status -eq 'OK') {
                    Write-Log "  Remove/rescan resolved the issue." 'OK'
                } elseif ($after2) {
                    Write-Log ("  Device re-appeared but still has problem: {0}" -f $after2.Problem) 'WARN'
                } else {
                    Write-Log "  Device did not re-enumerate (may need manual driver install)." 'WARN'
                }
            } catch {
                Write-Log ("  Remove/rescan failed: {0}" -f $_) 'ERROR'
            }
        }
    }

    Write-Log 'Repair phase complete.' 'OK'
} else {
    Write-Log 'Skipping Phase 3 (SkipRepair specified).'
}

# Update / Download 
if (-not $SkipUpdate) {
    Write-Log 'PHASE 4  -  Attempting driver updates via Windows Update' 'HEAD'

    # Check for PSWindowsUpdate module (preferred  -  gives fine-grained control)
    $wuModule = Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue

    if (-not $wuModule) {
        Write-Log 'PSWindowsUpdate not found  -  attempting to install from PSGallery...' 'WARN'
        try {
            Install-Module -Name PSWindowsUpdate -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log 'PSWindowsUpdate installed successfully.' 'OK'
            $wuModule = Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue
        } catch {
            Write-Log ("Failed to install PSWindowsUpdate: {0}" -f $_) 'WARN'
        }
    }

    if ($wuModule) {
        Write-Log 'PSWindowsUpdate module found  -  using it for driver updates.'
        Import-Module PSWindowsUpdate -Force

        Write-Log 'Searching Windows Update for driver updates...'
        try {
            $updates = @(Get-WUList -UpdateType Driver -ErrorAction Stop)
            if ($updates.Count -eq 0) {
                Write-Log 'No driver updates available via Windows Update.' 'OK'
            } else {
                Write-Log ("Found {0} driver update(s):" -f $updates.Count)
                $updates | ForEach-Object { Write-Log ("  {0}" -f $_.Title) }

                if ($PSCmdlet.ShouldProcess('All available driver updates', 'Install')) {
                    Install-WindowsUpdate -UpdateType Driver -AcceptAll -AutoReboot:$false |
                        ForEach-Object { Write-Log ("  {0}" -f $_.Title) }
                    Write-Log 'Driver updates installed.' 'OK'
                }
            }
        } catch {
            Write-Log ("PSWindowsUpdate error: {0}" -f $_) 'WARN'
            Write-Log 'Falling back to built-in Windows Update scan.' 'WARN'
            & usoclient StartScan 2>&1 | ForEach-Object { Write-Log "  $_" }
            Write-Log 'Windows Update scan triggered. Check Settings > Windows Update for results.' 'OK'
        }
    } else {
        Write-Log 'PSWindowsUpdate unavailable  -  falling back to built-in Windows Update / pnputil.' 'WARN'

        # Trigger WU driver scan via usoclient
        Write-Log 'Triggering Windows Update driver scan via usoclient...'
        & usoclient StartScan 2>&1 | ForEach-Object { Write-Log "  $_" }
        Write-Log 'Scan triggered. Driver downloads (if available) will appear in Windows Update.' 'OK'

        # pnputil update for each broken device using currently staged drivers
        Write-Log 'Attempting pnputil /update-driver for broken devices using staged INFs...'
        foreach ($dev in $errorDevices) {
            $id = $dev.InstanceId
            Write-Log ("  Updating: {0}" -f $dev.FriendlyName)
            $result = & pnputil /update-driver $id 2>&1
            $result | ForEach-Object { Write-Log "    $_" }
        }
    }

    Write-Log 'Update phase complete.' 'OK'
} else {
    Write-Log 'Skipping Phase 4 (SkipUpdate specified).'
}

# Final Summary
Write-Log '  Final device status check' 'HEAD'

$finalErrors = @(Get-PnpDevice -PresentOnly |
    Where-Object { $_.Status -eq 'Error' })

if ($finalErrors.Count -eq 0) {
    Write-Log 'All present devices are healthy.' 'OK'
} else {
    Write-Log ("$($finalErrors.Count) device(s) still in error state:") 'WARN'
    foreach ($dev in $finalErrors) {
        Write-Log ("  [{0}] {1}  -  {2}" -f $dev.Class, $dev.FriendlyName, $dev.Problem) 'WARN'
    }
    Write-Log 'Manual intervention or a hardware-specific driver download may be required.' 'WARN'
}

Write-Log "Log saved to: $LogPath" 'OK'
