# Windows Cleanup Script
#Requires -RunAsAdministrator

# Helpers
function Remove-ItemSilent {
    param([string]$Path)
    if (Test-Path $Path) {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Remove-FolderContents {
    param([string]$Path)
    if (Test-Path $Path) {
        Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-FolderSizeGB {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return 0 }
    $sum = (Get-ChildItem -Path $Path -Recurse -Force -File -ErrorAction SilentlyContinue |
        Measure-Object -Property Length -Sum).Sum
    if ($null -eq $sum) { $sum = 0 }
    return [math]::Round($sum / 1GB, 2)
}

$script:CapabilityAccessManagerHardDeleteThresholdMB = 512

function Register-DeleteOnReboot {
    param([string]$Path)

    $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    $existing = (Get-ItemProperty -Path $regPath -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue).PendingFileRenameOperations
    $nativePath = "\??\$Path"
    $updated = @()

    if ($existing) { $updated += $existing }
    $updated += $nativePath
    $updated += ''

    Set-ItemProperty -Path $regPath -Name 'PendingFileRenameOperations' -Value $updated -Type MultiString -ErrorAction SilentlyContinue
}

function Remove-ItemHard {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $true }

    try { [System.IO.File]::SetAttributes($Path, [System.IO.FileAttributes]::Normal) } catch { }
    try { & attrib.exe -r -s -h "$Path" 2>$null | Out-Null } catch { }
    try { & takeown.exe /f "$Path" /a /d Y 2>$null | Out-Null } catch { }
    try { & icacls.exe "$Path" /inheritance:e /grant 'Administrators:F' /c 2>$null | Out-Null } catch { }

    Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath $Path)) { return $true }

    try {
        cmd.exe /c "del /f /q \"$Path\"" | Out-Null
    } catch { }
    if (-not (Test-Path -LiteralPath $Path)) { return $true }

    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $stream.SetLength(0)
        $stream.Close()
    } catch { }

    Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath $Path)) { return $true }

    Register-DeleteOnReboot -Path $Path
    return $false
}

function Remove-RegistryKeyHard {
    param(
        [string]$RegistryPath,
        [string]$NativePath
    )

    Remove-Item -Path $RegistryPath -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path $RegistryPath) {
        & reg.exe delete $NativePath /f | Out-Null
    }

    return (-not (Test-Path $RegistryPath))
}

function Stop-ServiceIfRunning {
    param([string]$Name)
    $svc = Get-Service -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $svc) { return $false }
    $wasRunning = $svc.Status -eq 'Running'
    if ($wasRunning) {
        Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }
    return $wasRunning
}

function Start-ServiceIfNeeded {
    param(
        [string]$Name,
        [bool]$WasRunning
    )
    if ($WasRunning) {
        Start-Service -Name $Name -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }
}

function Get-UserProfiles {
    Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' -ErrorAction SilentlyContinue |
        ForEach-Object {
            $profile = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            if ($profile.ProfileImagePath -and $_.PSChildName -match '^S-1-5-21-' -and (Test-Path $profile.ProfileImagePath)) {
                [PSCustomObject]@{
                    Sid = $_.PSChildName
                    ProfilePath = $profile.ProfileImagePath
                }
            }
        }
}

function Get-UserProfilePaths {
    Get-UserProfiles | Select-Object -ExpandProperty ProfilePath
}

function Get-CapabilityAccessManagerDbCandidates {
    $roots = @()

    foreach ($profile in Get-UserProfilePaths) {
        $roots += (Join-Path $profile 'AppData\Local')
    }

    $roots += @(
        "$env:SystemRoot\ServiceProfiles\LocalService\AppData\Local",
        "$env:SystemRoot\ServiceProfiles\NetworkService\AppData\Local",
        "$env:SystemRoot\System32\config\systemprofile\AppData\Local",
        "$env:ProgramData\Microsoft",
        "$env:ProgramData\Packages"
    )

    foreach ($root in ($roots | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique)) {
        Get-ChildItem -Path $root -Recurse -Force -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -like 'CapabilityAccessManager.db*' -or
                $_.FullName -match 'CapabilityAccessManager'
            }
    }
}

function Write-LargeMicrosoftFilesReport {
    $microsoftRoot = "$env:ProgramData\Microsoft"
    if (-not (Test-Path $microsoftRoot)) { return }

    Write-Output "`n=== Largest files under ProgramData\\Microsoft (top 20) ==="
    Get-ChildItem -Path $microsoftRoot -Recurse -Force -File -ErrorAction SilentlyContinue |
        Sort-Object Length -Descending |
        Select-Object -First 20 FullName, @{ Name = 'Size (GB)'; Expression = { [math]::Round($_.Length / 1GB, 2) } } |
        Format-Table -AutoSize | Out-String | Write-Output
}

function Remove-WindowsInstallerSecondaryFiles {
    param([string]$InstallerPath = "$env:SystemRoot\Installer")

    if (-not (Test-Path $InstallerPath)) {
        Write-Output "  Windows Installer path not found: $InstallerPath"
        return
    }

    $files = @(Get-ChildItem -Path $InstallerPath -Recurse -Force -File -ErrorAction SilentlyContinue)
    if ($files.Count -lt 2) {
        Write-Output "  Only one (or no) installer file found; skipping duplicate cleanup."
        return
    }

    $duplicateGroups = @()
    $sizeGroups = $files | Group-Object Length | Where-Object { $_.Count -gt 1 }

    foreach ($sizeGroup in $sizeGroups) {
        $hashedItems = @()

        foreach ($file in $sizeGroup.Group) {
            try {
                $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $file.FullName -ErrorAction Stop).Hash
                $hashedItems += [PSCustomObject]@{
                    File = $file
                    Hash = $hash
                }
            } catch {
                Write-Output "  [WARN] Could not hash file, skipping: $($file.FullName)"
            }
        }

        $hashGroups = $hashedItems | Group-Object Hash | Where-Object { $_.Count -gt 1 }
        foreach ($hashGroup in $hashGroups) {
            $duplicateGroups += ,$hashGroup.Group
        }
    }

    if ($duplicateGroups.Count -eq 0) {
        Write-Output "  No duplicate installer files found."
        return
    }

    $groupsProcessed = 0
    $removedFiles = 0
    $scheduledForDelete = 0

    foreach ($group in $duplicateGroups) {
        $groupsProcessed++
        $ordered = $group | Sort-Object @{ Expression = { $_.File.LastWriteTimeUtc }; Descending = $true }, @{ Expression = { $_.File.FullName }; Descending = $false }
        $primary = $ordered[0].File
        $secondaryItems = $ordered | Select-Object -Skip 1

        Write-Output "  Primary (latest) kept: $($primary.FullName)"

        foreach ($secondaryItem in $secondaryItems) {
            $secondary = $secondaryItem.File
            $deleted = Remove-ItemHard -Path $secondary.FullName
            if ($deleted) {
                $removedFiles++
            } else {
                $scheduledForDelete++
                Write-Output "  Secondary file locked; scheduled for deletion on reboot: $($secondary.FullName)"
            }
        }
    }

    Write-Output "  Duplicate groups processed: $groupsProcessed"
    Write-Output "  Secondary duplicate files removed: $removedFiles"
    if ($scheduledForDelete -gt 0) {
        Write-Output "  Secondary duplicate files scheduled for reboot delete: $scheduledForDelete"
    }
}

function Reset-CapabilityAccessManagerStore {
    $backupDir = Join-Path $env:TEMP 'CapabilityAccessManagerBackup'
    $backupStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    New-Item -ItemType Directory -Path $backupDir -Force -ErrorAction SilentlyContinue | Out-Null

    $servicesToRestart = @()
    Get-Service -Name 'camsvc*', 'StateRepository', 'WpnService', 'InstallService', 'AppXSvc' -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.Status -eq 'Running') {
            $servicesToRestart += $_.Name
            Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        }
    }

    $registryResets = 0
    foreach ($profile in Get-UserProfiles) {
        $sid = $profile.Sid
        $profilePath = $profile.ProfilePath
        $hiveWasLoaded = Test-Path "Registry::HKEY_USERS\$sid"
        $loadedByScript = $false

        if (-not $hiveWasLoaded) {
            $ntUserPath = Join-Path $profilePath 'NTUSER.DAT'
            if (Test-Path $ntUserPath) {
                & reg.exe load ("HKU\$sid") $ntUserPath | Out-Null
                $hiveWasLoaded = Test-Path "Registry::HKEY_USERS\$sid"
                $loadedByScript = $hiveWasLoaded
            }
        }

        if ($hiveWasLoaded) {
            $camRoot = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager"
            $nativeCamRoot = "HKU\$sid\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager"
            if (Test-Path $camRoot) {
                $backupFile = Join-Path $backupDir ("CapabilityAccessManager_{0}_{1}.reg" -f $sid, $backupStamp)
                & reg.exe export $nativeCamRoot $backupFile /y | Out-Null

                if (Remove-RegistryKeyHard -RegistryPath $camRoot -NativePath $nativeCamRoot) {
                    $registryResets++
                    Write-Output "  Reset registry store for profile: $profilePath"
                } else {
                    Write-Output "  [WARN] Access denied or key locked for profile: $profilePath"
                }
            }
        }

        if ($loadedByScript) {
            & reg.exe unload ("HKU\$sid") | Out-Null
        }
    }

    $removedFiles = 0
    $scheduledForDelete = 0
    $dbCandidates = @(Get-CapabilityAccessManagerDbCandidates)

    foreach ($item in $dbCandidates) {
        $sizeMB = [math]::Round(($item.Length / 1MB), 1)
        if ($sizeMB -gt $script:CapabilityAccessManagerHardDeleteThresholdMB) {
            Write-Output "  Oversized CapabilityAccessManager DB detected: $($item.FullName) (${sizeMB} MB)"
        }

        $deleted = Remove-ItemHard -Path $item.FullName
        if ($deleted) {
            $removedFiles++
        } else {
            $scheduledForDelete++
            Write-Output "  File is locked or protected; scheduled for deletion on reboot: $($item.FullName)"
        }
    }

    foreach ($serviceName in ($servicesToRestart | Select-Object -Unique)) {
        Start-Service -Name $serviceName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }

    if ($registryResets -gt 0) {
        Write-Output "  Reset CapabilityAccessManager registry data for $registryResets profile(s)."
    } else {
        Write-Output "  No CapabilityAccessManager registry data was found in local user profiles."
    }

    if ($removedFiles -gt 0) {
        Write-Output "  Removed $removedFiles CapabilityAccessManager database file(s), including .db-wal and .db-shm files."
    } elseif ($dbCandidates.Count -eq 0) {
        Write-Output "  No CapabilityAccessManager database files were found in user or system profile locations."
    } else {
        Write-Output "  No CapabilityAccessManager database files were removed immediately."
    }

    if ($scheduledForDelete -gt 0) {
        Write-Output "  Scheduled $scheduledForDelete locked file(s) for deletion on reboot."
    }
}

# Prints a section header and captures free space so Write-SectionResult can diff it
function Write-Section {
    param([string]$Title)
    $script:SectionStart = (Get-PSDrive -Name C).Free
    Write-Output "`n=== $Title ==="
}

function Write-SectionResult {
    $deltaMB = [math]::Round(((Get-PSDrive -Name C).Free - $script:SectionStart) / 1MB, 1)

    if ($deltaMB -gt 0) {
        $script:TotalFreedMB += $deltaMB
        Write-Output "  -> Freed: ${deltaMB} MB"
    } elseif ($deltaMB -lt 0) {
        Write-Output "  -> Net disk usage increased by $([math]::Abs($deltaMB)) MB during this step"
    } else {
        Write-Output "  -> No measurable space reclaimed"
    }
}

# Returns $true if Windows is waiting for a reboot (blocks DISM /ResetBase)
function Test-PendingReboot {
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { return $true }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { return $true }
    $pfro = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' `
        -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    return ($null -ne $pfro)
}

# Space before cleanup
$Before = (Get-PSDrive -Name C).Free
$script:SectionStart = $Before
$script:TotalFreedMB = 0

Write-Output "Starting Windows 11 System Cleanup..."
Write-Output "Free space before: $([math]::Round($Before / 1GB, 2)) GB"

# Windows Temp
Write-Section "Windows Temp"
Remove-FolderContents -Path "$env:SystemRoot\Temp"
Remove-FolderContents -Path "$env:SystemRoot\System32\config\systemprofile\AppData\Local\Temp"
Write-SectionResult

# Prefetch
Write-Section "Prefetch"
Remove-FolderContents -Path "$env:SystemRoot\Prefetch"
Write-SectionResult

# Windows Error Reporting
Write-Section "Windows Error Reporting"
Remove-FolderContents -Path "$env:ProgramData\Microsoft\Windows\WER\ReportQueue"
Remove-FolderContents -Path "$env:ProgramData\Microsoft\Windows\WER\ReportArchive"
Remove-FolderContents -Path "$env:ProgramData\Microsoft\Windows\WER\Temp"
Write-SectionResult

# User temp / app caches / shader cache
Write-Section "User temp and app caches"
foreach ($profile in Get-UserProfilePaths) {
    Remove-FolderContents -Path (Join-Path $profile 'AppData\Local\Temp')
    Remove-FolderContents -Path (Join-Path $profile 'AppData\Local\CrashDumps')
    Remove-FolderContents -Path (Join-Path $profile 'AppData\Local\D3DSCache')
    Remove-FolderContents -Path (Join-Path $profile 'AppData\Local\Microsoft\Windows\INetCache')

    $packagesPath = Join-Path $profile 'AppData\Local\Packages'
    if (Test-Path $packagesPath) {
        Get-ChildItem -Path $packagesPath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            Remove-FolderContents -Path (Join-Path $_.FullName 'AC\Temp')
            Remove-FolderContents -Path (Join-Path $_.FullName 'AC\INetCache')
            Remove-FolderContents -Path (Join-Path $_.FullName 'TempState')
        }
    }
}
Write-SectionResult

# Windows Update download cache 
Write-Section "Windows Update SoftwareDistribution"
$wuauservWasRunning = Stop-ServiceIfRunning -Name 'wuauserv'
$bitsWasRunning = Stop-ServiceIfRunning -Name 'bits'
Remove-FolderContents -Path "$env:SystemRoot\SoftwareDistribution\Download"
Start-ServiceIfNeeded -Name 'wuauserv' -WasRunning $wuauservWasRunning
Start-ServiceIfNeeded -Name 'bits' -WasRunning $bitsWasRunning
Write-SectionResult

# Delivery Optimisation
Write-Section "Delivery Optimisation cache"
$doSvcWasRunning = Stop-ServiceIfRunning -Name 'DoSvc'
Remove-FolderContents -Path "$env:SystemRoot\SoftwareDistribution\DeliveryOptimization"
Remove-FolderContents -Path "$env:SystemDrive\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization"
Remove-FolderContents -Path "$env:ProgramData\Microsoft\Windows\DeliveryOptimization\Logs"
Start-ServiceIfNeeded -Name 'DoSvc' -WasRunning $doSvcWasRunning
Write-SectionResult

# CBS / Windows logs
Write-Section "CBS and Windows logs"
Remove-FolderContents -Path "$env:SystemRoot\Logs\CBS"
Remove-FolderContents -Path "$env:SystemRoot\Logs\DISM"
Remove-FolderContents -Path "$env:SystemRoot\Logs\WindowsUpdate"
Get-ChildItem -Path "$env:SystemRoot\Logs" -Recurse -File -Filter "*.log" -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
    Remove-Item -Force -ErrorAction SilentlyContinue
Write-SectionResult

# DISM component store
Write-Section "DISM component store cleanup"
$dismLog = "$env:TEMP\dism_cleanup.log"
$dismArgs = @('/Online', '/Cleanup-Image', '/StartComponentCleanup')
if (Test-PendingReboot) {
    Write-Output "  [WARN] Pending reboot detected - skipping /ResetBase to avoid 0x800f0806"
    Write-Output "  Running /StartComponentCleanup only (reboot then re-run for full cleanup)..."
} else {
    $dismArgs += '/ResetBase'
}

"DISM arguments: $($dismArgs -join ' ')" | Out-File -FilePath $dismLog -Encoding utf8 -Force
& dism.exe @dismArgs 2>&1 | Tee-Object -FilePath $dismLog -Append | Out-Null
$dismExitCode = $LASTEXITCODE

if ($dismExitCode -eq 0) {
    Write-Output "  DISM succeeded - full log: $dismLog"
} else {
    Write-Output "  DISM failed with exit code $dismExitCode - full log: $dismLog"
}
Write-SectionResult

# Windows Defender
Write-Section "Windows Defender scan history and cache"
Remove-FolderContents -Path "$env:ProgramData\Microsoft\Windows Defender\Scans\History"
Remove-FolderContents -Path "$env:ProgramData\Microsoft\Windows Defender\Scans\mpcache"
# Quarantine: only items older than 30 days (keep recent for review)
Get-ChildItem -Path "$env:ProgramData\Microsoft\Windows Defender\Quarantine" `
    -Recurse -Force -ErrorAction SilentlyContinue |
    Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
    Remove-Item -Force -ErrorAction SilentlyContinue
Write-SectionResult

# System crash dumps
Write-Section "System crash dumps"
Remove-ItemSilent -Path "$env:SystemRoot\MEMORY.DMP"
Remove-FolderContents -Path "$env:SystemRoot\Minidump"
Write-SectionResult

# Windows setup / kernel logs and caches
Write-Section "Windows setup and kernel logs"
Get-ChildItem -Path "$env:SystemRoot\Panther", "$env:SystemRoot\Logs\MoSetup", "$env:SystemRoot\LiveKernelReports" `
    -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-14) } |
    Remove-Item -Force -ErrorAction SilentlyContinue
Remove-FolderContents -Path "$env:SystemRoot\Downloaded Program Files"
Write-SectionResult

# Font cache
Write-Section "Font cache"
$fontCacheWasRunning = Stop-ServiceIfRunning -Name 'FontCache'
Remove-ItemSilent -Path "$env:SystemRoot\ServiceProfiles\LocalService\AppData\Local\FontCache"
Remove-ItemSilent -Path "$env:SystemRoot\ServiceProfiles\LocalService\AppData\Local\FontCache-System"
Get-ChildItem -Path "$env:SystemRoot\System32" -Filter "FNTCACHE.DAT" -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue
Start-ServiceIfNeeded -Name 'FontCache' -WasRunning $fontCacheWasRunning
Write-SectionResult

# IIS logs (if IIS is present)
if (Test-Path "$env:SystemDrive\inetpub\logs\LogFiles") {
    Write-Section "IIS logs (older than 30 days)"
    Get-ChildItem -Path "$env:SystemDrive\inetpub\logs\LogFiles" -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
    Write-SectionResult
}

# Service profile temp folders
Write-Section "Service profile temp folders"
Remove-FolderContents -Path "$env:ProgramData\Temp"
Remove-FolderContents -Path "$env:SystemRoot\ServiceProfiles\LocalService\AppData\Local\Temp"
Remove-FolderContents -Path "$env:SystemRoot\ServiceProfiles\NetworkService\AppData\Local\Temp"
Write-SectionResult

# Windows diagnostic / telemetry
Write-Section "Windows diagnostic staging"
Remove-FolderContents -Path "$env:ProgramData\Microsoft\Diagnosis\ETLLogs\AutoLogger"
Remove-FolderContents -Path "$env:ProgramData\Microsoft\Diagnosis\ETLLogs\ShutdownLogger"
Remove-FolderContents -Path "$env:ProgramData\Microsoft\Windows\WlanReport"
Write-SectionResult

# Thumbnail / icon cache (system)
Write-Section "System thumbnail/icon cache"
$tabletInputWasRunning = Stop-ServiceIfRunning -Name 'TabletInputService'
Remove-FolderContents -Path "$env:SystemRoot\System32\config\systemprofile\AppData\Local\Microsoft\Windows\Explorer"
Start-ServiceIfNeeded -Name 'TabletInputService' -WasRunning $tabletInputWasRunning
Write-SectionResult

# Windows Store staging
# NOTE: PostRebootEventCached.bin is a FILE - use Remove-ItemSilent not Remove-FolderContents
Write-Section "Windows Store staging"
Remove-ItemSilent -Path "$env:SystemRoot\SoftwareDistribution\PostRebootEventCached.bin"
Remove-FolderContents -Path "$env:ProgramData\Microsoft\Windows\AppRepository\Packages"
Write-SectionResult

# Windows Update logs
Write-Section "Windows Update logs"
Remove-FolderContents -Path "$env:ProgramData\USOShared\Logs"
Get-ChildItem -Path "$env:SystemRoot" -File -Filter "WindowsUpdate.log" -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue
Write-SectionResult

# Windows Search index and temp
# Clears the search index database and transaction logs; Windows Search will rebuild it
Write-Section "Windows Search index and temp"
$wSearchWasRunning = Stop-ServiceIfRunning -Name 'WSearch'
$searchPath = "$env:ProgramData\Microsoft\Search\Data\Applications\Windows"
if (Test-Path $searchPath) {
    Get-ChildItem -Path $searchPath -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^(Windows\.edb|edb.*\.log|tmp\.edb|.*\.jrs)$' } |
        Remove-Item -Force -ErrorAction SilentlyContinue
    Write-Output "  Windows Search index files cleared; the index will rebuild automatically."
} else {
    Write-Output "  Windows Search data path not found."
}
Start-ServiceIfNeeded -Name 'WSearch' -WasRunning $wSearchWasRunning
Write-SectionResult

# SCCM / ConfigMgr cache (if present)
if (Test-Path "$env:SystemRoot\ccmcache") {
    Write-Section "SCCM ConfigMgr cache"
    Remove-FolderContents -Path "$env:SystemRoot\ccmcache"
    Write-SectionResult
}

# Azure / MMA / Monitoring agent logs
foreach ($agentPath in @(
    "$env:ProgramData\Microsoft\Windows Azure",
    "$env:ProgramData\Microsoft Monitoring Agent\Agent\Health Service State\Monitoring Host Temporary Files 6\",
    "$env:ProgramData\Microsoft\Windows\OneCollector"
)) {
    if (Test-Path $agentPath) {
        Write-Section "Agent logs: $(Split-Path $agentPath -Leaf)"
        Get-ChildItem -Path $agentPath -Recurse -File -Filter "*.log" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-14) } |
            Remove-Item -Force -ErrorAction SilentlyContinue
        Write-SectionResult
    }
}

# Hyper-V logs (if present)
if (Test-Path "$env:ProgramData\Microsoft\Windows\Hyper-V") {
    Write-Section "Hyper-V logs"
    Get-ChildItem -Path "$env:ProgramData\Microsoft\Windows\Hyper-V" -Recurse -File -Filter "*.log" `
        -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
    Write-SectionResult
}

# Docker dangling data (if Docker is installed)
if (Get-Command docker -ErrorAction SilentlyContinue) {
    Write-Section "Docker dangling images and build cache"
    & docker system prune -f 2>&1 | Write-Output
    Write-SectionResult
}

# CapabilityAccessManager permission store
Write-Section "CapabilityAccessManager permission store"
Reset-CapabilityAccessManagerStore
Write-SectionResult

# Windows Installer duplicate files (keep the newest copy in each duplicate set)
Write-Section "Windows Installer duplicate files"
Remove-WindowsInstallerSecondaryFiles
Write-SectionResult

# Package Cache size warning
$pkgCacheGB = Get-FolderSizeGB -Path "$env:ProgramData\Package Cache"
if ($pkgCacheGB -gt 1) {
    Write-Output "`n=== Package Cache ==="
    Write-Output "  Size: ${pkgCacheGB} GB"
    Write-Output "  [INFO] Not auto-cleaned - removing it breaks VS repair/uninstall."
    Write-Output "  To reclaim space: open Visual Studio Installer and remove unused workloads/versions."
}

# Event log trim (only if very large)
Write-Section "Oversized event log trim (>50,000 entries)"
foreach ($logName in @('Application', 'System', 'Setup')) {
    try {
        $evLog = [System.Diagnostics.EventLog]::new($logName)
        $count = $evLog.Entries.Count
        $evLog.Dispose()
        if ($count -gt 50000) {
            Write-Output "  Clearing $logName ($count entries)..."
            Clear-EventLog -LogName $logName -ErrorAction SilentlyContinue
        }
    } catch { }
}
Write-SectionResult

# Disk Cleanup via CleanMgr
Write-Section "CleanMgr (system files)"
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
$Keys = Get-ChildItem -Path $RegPath -ErrorAction SilentlyContinue
foreach ($Key in $Keys) {
    Set-ItemProperty -Path $Key.PSPath -Name "StateFlags0100" -Value 2 -Type DWord -ErrorAction SilentlyContinue
}
Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:100" -Wait -NoNewWindow -ErrorAction SilentlyContinue
Write-SectionResult

# ProgramData top-level size breakdown
Write-Output "`n=== ProgramData folder size breakdown (top consumers) ==="
Get-ChildItem -Path $env:ProgramData -Directory -Force -ErrorAction SilentlyContinue |
    ForEach-Object {
        $sz = (Get-ChildItem -Path $_.FullName -Recurse -Force -File -ErrorAction SilentlyContinue |
               Measure-Object -Property Length -Sum).Sum
        if ($null -eq $sz) { $sz = 0 }
        [PSCustomObject]@{ Folder = $_.Name; 'Size (GB)' = [math]::Round($sz / 1GB, 2) }
    } |
    Where-Object { $_.'Size (GB)' -gt 0.1 } |
    Sort-Object 'Size (GB)' -Descending |
    Format-Table -AutoSize | Out-String | Write-Output

Write-LargeMicrosoftFilesReport

# Final summary
$After = (Get-PSDrive -Name C).Free
$NetChangeGB = [math]::Round(($After - $Before) / 1GB, 2)
$GrossFreedGB = [math]::Round(($script:TotalFreedMB / 1024), 2)

Write-Output "Cleanup complete. Gross reclaimed across measured steps: $GrossFreedGB GB"
if ($NetChangeGB -ge 0) {
    Write-Output "Net free-space change by end of run: +$NetChangeGB GB"
} else {
    Write-Output "Net free-space change by end of run: $NetChangeGB GB"
    Write-Output "  [INFO] Some Windows services may have recreated cache/log files during the run."
}
Write-Output "Free space after:  $([math]::Round($After / 1GB, 2)) GB"