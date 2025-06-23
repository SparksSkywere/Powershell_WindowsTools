# Ensure script runs as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires administrator privileges. Restarting as administrator..." -ForegroundColor Yellow
    Start-Process PowerShell -Verb RunAs "-File `"$PSCommandPath`""
    exit
}

Write-Host "Starting Windows Optimisation Process..." -ForegroundColor Green
$logPath = "$env:TEMP\WindowsOptimizer.log"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logPath -Append
    Write-Host $Message -ForegroundColor Cyan
}

# Services to disable for better performance
$servicesToDisable = @(
    "Fax", "WerSvc", "Spooler", "SCardSvr", "TabletInputService",
    "WSearch", "HomeGroupListener", "HomeGroupProvider", "WMPNetworkSvc",
    "RemoteRegistry", "RemoteAccess", "SharedAccess", "lfsvc", "MapsBroker",
    "RetailDemo", "dmwappushservice", "TrkWks", "WbioSrvc"
)

Write-Log "Optimizing Windows Services..."
foreach ($service in $servicesToDisable) {
    try {
        $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
        if ($svc -and $svc.StartType -ne 'Disabled') {
            Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
            Set-Service -Name $service -StartupType Disabled -ErrorAction SilentlyContinue
            Write-Log "Disabled service: $service"
        }
    }
    catch {
        Write-Log "Could not disable service: $service"
    }
}

Write-Log "Cleaning Temporary Files..."
# Clean temporary files
$tempPaths = @(
    "$env:TEMP\*",
    "$env:WINDIR\Temp\*",
    "$env:LOCALAPPDATA\Microsoft\Windows\INetCache\*",
    "$env:LOCALAPPDATA\Microsoft\Windows\Temporary Internet Files\*"
)

foreach ($path in $tempPaths) {
    try {
        Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Log "Could not clean: $path"
    }
}

# Run Disk Cleanup
cleanmgr /sagerun:1

Write-Log "Optimizing Visual Effects for Performance..."
# Configure visual effects for best performance
$regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"
if (!(Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
Set-ItemProperty -Path $regPath -Name "VisualFXSetting" -Value 2

# Disable animations
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop\WindowMetrics" -Name "MinAnimate" -Value "0"
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "MenuShowDelay" -Value "0"

Write-Log "Disabling Background Apps..."
# Disable background apps
$backgroundAppsPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications"
if (Test-Path $backgroundAppsPath) {
    Set-ItemProperty -Path $backgroundAppsPath -Name "GlobalUserDisabled" -Value 1
}

Write-Log "Configuring Privacy and Telemetry Settings..."
# Disable telemetry
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" -Name "AllowTelemetry" -Value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0

# Disable location tracking
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location" -Name "Value" -Value "Deny"

# Disable advertising ID
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -Name "Enabled" -Value 0

Write-Log "Optimizing Startup Programs..."
# Disable common startup programs that aren't essential
$startupApps = @("Skype", "Spotify", "Steam", "Discord", "Adobe", "Office")
foreach ($app in $startupApps) {
    Get-WmiObject -Class Win32_StartupCommand | Where-Object { $_.Name -like "*$app*" } | ForEach-Object {
        try {
            $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
            Remove-ItemProperty -Path $regPath -Name $_.Name -ErrorAction SilentlyContinue
            Write-Log "Disabled startup item: $($_.Name)"
        }
        catch {
            Write-Log "Could not disable startup item: $($_.Name)"
        }
    }
}

Write-Log "Optimizing Network Settings..."
# Disable Windows Update P2P sharing
$deliveryOptPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config"
try {
    if (!(Test-Path $deliveryOptPath)) {
        New-Item -Path $deliveryOptPath -Force | Out-Null
    }
    Set-ItemProperty -Path $deliveryOptPath -Name "DODownloadMode" -Value 0 -ErrorAction Stop
    Write-Log "Disabled Windows Update P2P sharing"
}
catch {
    Write-Log "Could not configure Delivery Optimization settings: $($_.Exception.Message)"
}

Write-Log "Configuring Windows Defender for Performance..."
# Configure Windows Defender for better performance
try {
    Set-MpPreference -EnableControlledFolderAccess Disabled -ErrorAction SilentlyContinue
    Set-MpPreference -PUAProtection Disabled -ErrorAction SilentlyContinue
    Set-MpPreference -SubmitSamplesConsent Never -ErrorAction SilentlyContinue
}
catch {
    Write-Log "Could not configure Windows Defender settings"
}

Write-Log "Optimizing System Performance Settings..."
# Disable hibernation to save disk space
powercfg /hibernate off

# Optimize paging file
$cs = Get-WmiObject -Class Win32_ComputerSystem
$ram = [math]::Round($cs.TotalPhysicalMemory / 1GB)
$pagingFileSize = [math]::Max(1024, $ram * 1024 / 2)

$pageFile = Get-WmiObject -Class Win32_PageFileSetting
if ($pageFile) {
    $pageFile.InitialSize = $pagingFileSize
    $pageFile.MaximumSize = $pagingFileSize
    $pageFile.Put() | Out-Null
}

Write-Log "Disabling Unnecessary Windows Features..."
# Disable Windows features that consume resources
$featuresToDisable = @(
    "MediaPlayback", "WorkFolders-Client", "Printing-XPSServices-Features",
    "SMB1Protocol", "MicrosoftWindowsPowerShellV2Root"
)

foreach ($feature in $featuresToDisable) {
    try {
        Disable-WindowsOptionalFeature -Online -FeatureName $feature -NoRestart -ErrorAction SilentlyContinue
        Write-Log "Disabled Windows feature: $feature"
    }
    catch {
        Write-Log "Could not disable feature: $feature"
    }
}

Write-Log "Configuring System Responsiveness..."
# Improve system responsiveness
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" -Name "SystemResponsiveness" -Value 10
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" -Name "NetworkThrottlingIndex" -Value 4294967295

Write-Log "Cleaning Event Logs..."
# Clear event logs to free up space
try {
    $logs = Get-WinEvent -ListLog * -ErrorAction SilentlyContinue | Where-Object { 
        $_.RecordCount -gt 0 -and $_.IsEnabled -eq $true 
    }
    foreach ($log in $logs) {
        try {
            if ($log.LogName -and $log.LogName -notlike "*Analytic*" -and $log.LogName -notlike "*Debug*") {
                [System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.ClearLog($log.LogName)
                Write-Log "Cleared event log: $($log.LogName)"
            }
        }
        catch {
            # Some logs cannot be cleared, continue silently
            Write-Log "Could not clear log: $($log.LogName)"
        }
    }
}
catch {
    Write-Log "Could not enumerate event logs: $($_.Exception.Message)"
}

Write-Log "Optimizing Registry for Performance..."
# Registry optimisations
$regOptimisations = @{
    "HKLM:\SYSTEM\CurrentControlSet\Control\PriorityControl" = @{
        "Win32PrioritySeparation" = 26
    }
    "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" = @{
        "ClearPageFileAtShutdown" = 0
        "DisablePagingExecutive" = 1
    }
    "HKCU:\Control Panel\Mouse" = @{
        "MouseHoverTime" = "10"
    }
}

foreach ($path in $regOptimisations.Keys) {
    if (!(Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    foreach ($name in $regOptimisations[$path].Keys) {
        Set-ItemProperty -Path $path -Name $name -Value $regOptimisations[$path][$name]
    }
}

Write-Log "Optimisation Complete!"
Write-Host "`nWindows Optimisation completed successfully!" -ForegroundColor Green
Write-Host "Log file saved to: $logPath" -ForegroundColor Yellow
Write-Host "A system restart is recommended to apply all changes." -ForegroundColor Yellow

# Prompt for restart
$restart = Read-Host "`nWould you like to restart now? (Y/N)"
if ($restart -eq 'Y' -or $restart -eq 'y') {
    Write-Host "Restarting system in 10 seconds..." -ForegroundColor Red
    Start-Sleep -Seconds 10
    Restart-Computer -Force
}
