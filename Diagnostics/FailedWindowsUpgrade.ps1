#Requires -RunAsAdministrator
function Write-LogSection {
    param ([string]$Title)
    
    Write-Host "`n==================== $Title ====================" -ForegroundColor Cyan
}

function Show-Menu {
    Clear-Host
    Write-Host "==================== Windows 11 Upgrade Diagnostic Tool ====================" -ForegroundColor Cyan
    Write-Host "This tool will help diagnose why your Windows 11 upgrade from 23H2 to 24H2 is failing.`n" -ForegroundColor White
    
    Write-Host "1: Deep Scan" -ForegroundColor Green
    Write-Host "   Perform a comprehensive scan of drivers, services, applications, and system compatibility"
    
    Write-Host "`n2: Diagnose Setup" -ForegroundColor Yellow
    Write-Host "   Analyze setup logs and configuration for upgrade blockers"
    
    Write-Host "`n3: Exit" -ForegroundColor Red
    Write-Host "   Exit the diagnostic tool"
    
    Write-Host "`n=================================="
}

# Basic diagnostic functions (from original script)
function Test-SystemRequirements {
    Write-LogSection "System Requirements Check"
    
    # Check CPU compatibility
    $processor = Get-WmiObject -Class Win32_Processor
    Write-Host "CPU: $($processor.Name)"
    
    # Check RAM
    $totalMemoryInGB = [math]::Round((Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
    Write-Host "RAM: $totalMemoryInGB GB"
    if ($totalMemoryInGB -lt 4) {
        Write-Host "WARNING: Windows 11 requires at least 4 GB RAM" -ForegroundColor Yellow
    }
    
    # Check TPM version
    try {
        $tpm = Get-WmiObject -Namespace "root\CIMV2\Security\MicrosoftTpm" -Class Win32_Tpm -ErrorAction Stop
        Write-Host "TPM Available: Yes"
        Write-Host "TPM Version: $($tpm.ManufacturerVersion)"
        if ([version]$tpm.ManufacturerVersion -lt [version]"2.0") {
            Write-Host "WARNING: Windows 11 requires TPM version 2.0" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "WARNING: TPM not available or accessible" -ForegroundColor Yellow
    }
    
    # Check SecureBoot status
    try {
        $secureBootStatus = Confirm-SecureBootUEFI -ErrorAction Stop
        Write-Host "Secure Boot Enabled: $secureBootStatus"
        if (!$secureBootStatus) {
            Write-Host "WARNING: Windows 11 requires Secure Boot to be enabled" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "WARNING: Could not determine Secure Boot status" -ForegroundColor Yellow
    }
}

function Check-DiskSpace {
    Write-LogSection "Disk Space Check"
    
    # Get system drive
    $systemDrive = $env:SystemDrive
    $drive = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='$systemDrive'"
    $freeSpaceGB = [math]::Round($drive.FreeSpace / 1GB, 2)
    
    Write-Host "System Drive: $systemDrive"
    Write-Host "Free Space: $freeSpaceGB GB"
    
    if ($freeSpaceGB -lt 20) {
        Write-Host "WARNING: Less than 20GB free space available. Windows upgrades typically require at least 20GB free space." -ForegroundColor Yellow
    }
}

function Get-SetupLogs {
    Write-LogSection "Setup Logs Analysis"
    
    $setupLogPaths = @(
        "$env:SystemRoot\Panther",
        "$env:SystemRoot\Logs\MoSetup",
        "$env:SystemRoot\Logs\CBS"
    )
    
    $errorPatterns = @(
        "Error", "Failed", "failure", "cannot", "unable", 
        "not supported", "compatibility", "incompatible", "block", "prevent"
    )
    
    foreach ($logPath in $setupLogPaths) {
        if (Test-Path $logPath) {
            Write-Host "Checking logs in: $logPath" -ForegroundColor DarkCyan
            
            # Get log files, focusing on latest ones
            $logFiles = Get-ChildItem -Path $logPath -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 5
            
            foreach ($logFile in $logFiles) {
                Write-Host "  Analyzing: $($logFile.Name)" -ForegroundColor DarkGray
                $logContent = Get-Content $logFile.FullName -ErrorAction SilentlyContinue
                
                if ($logContent) {
                    $errorLines = $logContent | Where-Object { 
                        $line = $_.ToLower()
                        ($errorPatterns | ForEach-Object { $line -match $_ }) -contains $true
                    }
                    
                    if ($errorLines) {
                        Write-Host "    Found potential issues:" -ForegroundColor Yellow
                        foreach ($line in ($errorLines | Select-Object -First 10)) {
                            Write-Host "      $line" -ForegroundColor DarkYellow
                        }
                    }
                }
            }
            
            # Look specifically for setuperr.log in Panther folder
            if ($logPath -eq "$env:SystemRoot\Panther" -and (Test-Path "$logPath\setuperr.log")) {
                Write-Host "`n  IMPORTANT: Setup Error Log Content from $logPath\setuperr.log" -ForegroundColor Red
                Get-Content "$logPath\setuperr.log" -Tail 20 | ForEach-Object {
                    Write-Host "    $_" -ForegroundColor DarkRed
                }
            }
        }
    }
}

function Check-PendingUpdates {
    Write-LogSection "Windows Update Status"
    
    try {
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateUpdateSearcher()
        $pendingUpdates = $updateSearcher.Search("IsInstalled=0 and Type='Software'").Updates
        
        if ($pendingUpdates.Count -gt 0) {
            Write-Host "Found $($pendingUpdates.Count) pending updates that might need to be installed before feature update:" -ForegroundColor Yellow
            foreach ($update in $pendingUpdates) {
                Write-Host "  $($update.Title)" -ForegroundColor DarkYellow
            }
        }
        else {
            Write-Host "No pending Windows updates found." -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Could not check for Windows updates: $_" -ForegroundColor Red
    }
}

function Check-IncompatibleSoftware {
    Write-LogSection "Checking for Incompatible Software"
    
    $potentialBlockers = @(
        "Antivirus", "VPN", "Firewall", "Virtualization", "Norton", "McAfee", "Kaspersky", 
        "Avast", "AVG", "Bitdefender", "Cisco", "VirtualBox", "VMware", "Hyper-V"
    )
    
    $installedSoftware = Get-WmiObject -Class Win32_Product | Select-Object Name, Vendor
    
    $matches = $installedSoftware | Where-Object { 
        $software = $_
        ($potentialBlockers | ForEach-Object { $software.Name -match $_ }) -contains $true
    }
    
    if ($matches) {
        Write-Host "Found potentially incompatible software that might interfere with Windows upgrades:" -ForegroundColor Yellow
        $matches | ForEach-Object {
            Write-Host "  $($_.Name) - $($_.Vendor)" -ForegroundColor DarkYellow
        }
        Write-Host "`nConsider temporarily disabling or uninstalling these applications before upgrading." -ForegroundColor Yellow
    }
    else {
        Write-Host "No commonly problematic software detected." -ForegroundColor Green
    }
}

function Check-WindowsServicesStatus {
    Write-LogSection "Windows Services Check"
    
    $criticalServices = @(
        @{Name="Windows Update"; ServiceName="wuauserv"},
        @{Name="Background Intelligent Transfer Service"; ServiceName="bits"},
        @{Name="Windows Installer"; ServiceName="msiserver"},
        @{Name="Windows Module Installer"; ServiceName="TrustedInstaller"},
        @{Name="Cryptographic Services"; ServiceName="CryptSvc"}
    )
    
    foreach ($service in $criticalServices) {
        $serviceStatus = Get-Service -Name $service.ServiceName -ErrorAction SilentlyContinue
        
        if ($serviceStatus) {
            Write-Host "$($service.Name) ($($service.ServiceName)): $($serviceStatus.Status)" -ForegroundColor $(
                if ($serviceStatus.Status -eq "Running") { "Green" } else { "Yellow" }
            )
            
            if ($serviceStatus.StartType -eq "Disabled") {
                Write-Host "  WARNING: This service is disabled but may be required for Windows upgrades" -ForegroundColor Red
            }
        }
        else {
            Write-Host "$($service.Name) ($($service.ServiceName)): Not found" -ForegroundColor Red
        }
    }
}

function Get-CurrentWindowsVersion {
    Write-LogSection "Current Windows Version"
    
    $currentBuild = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    $releaseId = $currentBuild.ReleaseId
    if (-not $releaseId) {
        $releaseId = $currentBuild.DisplayVersion
    }
    
    Write-Host "Windows Edition: $(Get-WindowsEdition -Online | Select-Object -ExpandProperty Edition)"
    Write-Host "Version: $releaseId"
    Write-Host "OS Build: $($currentBuild.CurrentBuild).$($currentBuild.UBR)"
}

function Get-UpdateBlockerRecommendations {
    Write-LogSection "Specific Recommendations Based on Analysis"
    
    Write-Host @"
Based on the analysis of your ScanResult.xml, your PC has a driver or service 
that isn't ready for Windows 11 24H2. Here are specific steps to resolve this issue:

1. UPDATE DRIVERS:
   - Visit your PC manufacturer's website to download the latest drivers
   - Or use Windows Update to check for driver updates
   - Focus particularly on display/graphics, network, and storage drivers

2. CHECK FOR CONFLICTING SOFTWARE:
   - Temporarily disable or uninstall security software (antivirus, firewall)
   - Disable any VPN software
   - Remove any system optimization or cleaning tools

3. PREPARE FOR UPGRADE:
   - Backup your important data
   - Disconnect unnecessary peripherals (printers, external drives, etc.)
   - Ensure you have at least 20GB of free space
   - Connect to a reliable power source

4. ALTERNATIVE UPGRADE METHODS:
   - Try using the Media Creation Tool to perform an in-place upgrade
   - Create bootable media and perform a clean installation (saving your files)

5. MICROSOFT WORKAROUNDS:
   - Check the Microsoft Known Issues page for Windows 11 24H2 for possible workarounds
   - Visit the link in your error message: https://go.microsoft.com/fwlink/?LinkId=2280120
"@ -ForegroundColor White
}

function Process-ScanResultXML {
    param(
        [Parameter(Mandatory=$false)]
        [string]$XmlFilePath = $null
    )
    
    Write-LogSection "Compatibility Scan Results Analysis"
    
    if (-not $XmlFilePath) {
        # Try to find ScanResult.xml
        $scanResultPaths = @(
            'C:/$WINDOWS.~BT/Sources/Panther/ScanResult.xml',
            'C:\$WINDOWS.~BT\Sources\Panther\ScanResult.xml',
            "C:\Windows\Panther\ScanResult.xml"
        )
        
        foreach ($path in $scanResultPaths) {
            if (Test-Path -LiteralPath $path) {
                $XmlFilePath = $path
                break
            }
        }
    }
    
    if ($XmlFilePath -and (Test-Path -LiteralPath $XmlFilePath)) {
        Write-Host "Analyzing ScanResult.xml at: $XmlFilePath" -ForegroundColor Green
        
        try {
            [xml]$scanResultXml = Get-Content -LiteralPath $XmlFilePath -ErrorAction Stop
            
            $blockingItems = $scanResultXml.CompatReport.Hardware.HardwareItem | 
                Where-Object { $_.CompatibilityInfo.BlockingType -eq "Hard" }
            
            if ($blockingItems) {
                Write-Host "`nBLOCKING ISSUES FOUND:" -ForegroundColor Red
                
                foreach ($item in $blockingItems) {
                    Write-Host "  Type: $($item.HardwareType)" -ForegroundColor Red
                    Write-Host "  Title: $($item.CompatibilityInfo.Title)" -ForegroundColor Red
                    Write-Host "  Message: $($item.CompatibilityInfo.Message)" -ForegroundColor Red
                    
                    if ($item.Link) {
                        Write-Host "  More info: $($item.Link.Target)" -ForegroundColor Yellow
                    }
                    Write-Host ""
                }
            }
            
            # Check for blocking drivers
            $blockingDrivers = $scanResultXml.CompatReport.DriverPackages.DriverPackage | 
                Where-Object { $_.BlockMigration -eq "True" }
                
            if ($blockingDrivers) {
                Write-Host "`nBLOCKING DRIVERS FOUND:" -ForegroundColor Red
                $blockingDrivers | ForEach-Object {
                    Write-Host "  Driver: $($_.Inf)" -ForegroundColor Red
                }
            } else {
                Write-Host "`nInterestingly, no specific blocking driver packages are marked in the XML." -ForegroundColor Yellow
                Write-Host "This means the issue is likely with a service or a driver compatibility that isn't explicitly marked." -ForegroundColor Yellow
            }
            
            return $true
        }
        catch {
            Write-Host "Error analyzing the XML: $_" -ForegroundColor Red
            return $false
        }
    }
    else {
        Write-Host "ScanResult.xml not found. Run Windows 11 Setup to generate compatibility report." -ForegroundColor Yellow
        return $false
    }
}

# Deep scan functions (new enhanced functionality)
function Perform-DeepDriverScan {
    Write-LogSection "Deep Driver Compatibility Scan"
    
    Write-Host "Scanning all installed drivers for potential compatibility issues..." -ForegroundColor Cyan
    
    # Get all devices
    $allDevices = Get-WmiObject Win32_PnPEntity
    Write-Host "Found $($allDevices.Count) devices installed on your system." -ForegroundColor White
    
    # Get device driver details
    $deviceDriverDetails = @()
    $problemDevices = @()
    $outdatedDrivers = @()
    
    # Analyze devices
    foreach ($device in $allDevices) {
        $deviceName = $device.Caption
        $deviceID = $device.DeviceID
        $status = $device.Status
        $deviceProblems = $device.ConfigManagerErrorCode
        
        # Create device details object
        $deviceDetail = New-Object PSObject -Property @{
            Name = $deviceName
            ID = $deviceID
            Status = $status
            ErrorCode = $deviceProblems
        }
        
        # Check if device has problems
        if ($deviceProblems -ne 0) {
            $problemDevices += $deviceDetail
        }
        
        $deviceDriverDetails += $deviceDetail
    }
    
    # Check for devices with problems
    if ($problemDevices.Count -gt 0) {
        Write-Host "`nDEVICES WITH PROBLEMS DETECTED:" -ForegroundColor Red
        $problemDevices | ForEach-Object {
            Write-Host "  Device: $($_.Name)" -ForegroundColor Red
            Write-Host "  Status: $($_.Status)" -ForegroundColor Red
            Write-Host "  Error Code: $($_.ErrorCode)" -ForegroundColor Red
            
            # Translate error code to message
            $errorMessage = switch ($_.ErrorCode) {
                1 {"The device is not configured correctly."}
                2 {"Windows cannot load the driver for this device."}
                3 {"The driver for this device might be corrupted, or your system may be running low on memory or other resources."}
                4 {"This device is not working properly. One of its drivers or your registry might be corrupted."}
                5 {"The driver for this device needs a resource that Windows cannot manage."}
                6 {"The boot configuration for this device conflicts with other devices."}
                7 {"Cannot filter."}
                8 {"The driver loader for the device is missing."}
                9 {"This device is not working properly because the controlling firmware is reporting the resources for the device incorrectly."}
                10 {"This device cannot start."}
                11 {"This device failed."}
                12 {"This device cannot find enough free resources that it can use."}
                13 {"Windows cannot verify this device's resources."}
                14 {"This device cannot work properly until you restart your computer."}
                15 {"This device is not working properly because there is probably a re-enumeration problem."}
                16 {"Windows cannot identify all the resources this device uses."}
                17 {"This device is asking for an unknown resource type."}
                18 {"Reinstall the drivers for this device."}
                19 {"Failure using the VxD loader."}
                20 {"Your registry might be corrupted."}
                21 {"System failure: Try changing the driver for this device. If that does not work, see your hardware documentation."}
                22 {"This device is disabled."}
                23 {"System failure: Try changing the driver for this device. If that doesn't work, see your hardware documentation."}
                24 {"This device is not present, is not working properly, or does not have all its drivers installed."}
                25 {"Windows is still setting up this device."}
                26 {"Windows is still setting up this device."}
                27 {"This device does not have valid log configuration."}
                28 {"The drivers for this device are not installed."}
                29 {"This device is disabled because the firmware of the device did not give it the required resources."}
                30 {"This device is using an Interrupt Request (IRQ) resource that another device is using."}
                31 {"This device is not working properly because Windows cannot load the drivers required for this device."}
                default {"Unknown error."}
            }
            
            Write-Host "  Problem: $errorMessage" -ForegroundColor Red
            Write-Host ""
        }
        
        Write-Host "These device issues could be preventing your Windows upgrade." -ForegroundColor Yellow
        Write-Host "Consider updating or reinstalling drivers for these devices before attempting the upgrade again." -ForegroundColor Yellow
    }
    else {
        Write-Host "No devices with hardware problems were detected." -ForegroundColor Green
    }
    
    # Now check for driver details using pnputil
    Write-Host "`nAnalyzing driver packages..." -ForegroundColor Cyan
    
    $allDrivers = & pnputil /enum-drivers
    $driverPackages = @()
    $currentDriver = $null
    
    foreach ($line in $allDrivers) {
        if ($line -match "Published name:(.*)") {
            if ($currentDriver -ne $null) {
                $driverPackages += $currentDriver
            }
            $currentDriver = @{
                PublishedName = $Matches[1].Trim()
            }
        }
        elseif ($line -match "Original name:(.*)") {
            $currentDriver.OriginalName = $Matches[1].Trim()
        }
        elseif ($line -match "Provider name:(.*)") {
            $currentDriver.ProviderName = $Matches[1].Trim()
        }
        elseif ($line -match "Class name:(.*)") {
            $currentDriver.ClassName = $Matches[1].Trim()
        }
        elseif ($line -match "Date and version:(.*)") {
            $currentDriver.DateAndVersion = $Matches[1].Trim()
            
            # Extract date from format "mm/dd/yyyy x.x.x.x"
            if ($currentDriver.DateAndVersion -match "(\d{1,2}\/\d{1,2}\/\d{4})") {
                $driverDate = [DateTime]::Parse($Matches[1])
                $currentDriver.Date = $driverDate
                
                # Check if driver is older than 2 years
                if ($driverDate -lt (Get-Date).AddYears(-2)) {
                    $currentDriver.IsOld = $true
                    $outdatedDrivers += $currentDriver
                }
                else {
                    $currentDriver.IsOld = $false
                }
            }
        }
    }
    
    if ($currentDriver -ne $null) {
        $driverPackages += $currentDriver
    }
    
    # Categories often associated with Windows upgrade issues
    $problematicCategories = @(
        "Display", "Graphics", "Video", "Monitor", "Network", "Wireless", "WiFi", "Storage", 
        "RAID", "NVMe", "SSD", "HDD", "USB", "Audio", "Sound", "Security", "Biometric", 
        "Fingerprint", "Camera", "Bluetooth", "Chipset", "Thunderbolt"
    )
    
    $potentialProblemDrivers = $driverPackages | Where-Object {
        $driver = $_
        ($problematicCategories | ForEach-Object { 
            $driver.ClassName -match $_ -or 
            $driver.ProviderName -match $_ -or 
            $driver.OriginalName -match $_
        }) -contains $true
    }
    
    if ($potentialProblemDrivers.Count -gt 0) {
        Write-Host "`nPOTENTIAL PROBLEMATIC DRIVERS IDENTIFIED:" -ForegroundColor Yellow
        $potentialProblemDrivers | ForEach-Object {
            $driver = $_
            $oldWarning = if ($driver.IsOld) { " (OUTDATED)" } else { "" }
            
            Write-Host "  Driver: $($driver.OriginalName)$oldWarning" -ForegroundColor Yellow
            Write-Host "  Provider: $($driver.ProviderName)" -ForegroundColor DarkYellow
            Write-Host "  Class: $($driver.ClassName)" -ForegroundColor DarkYellow
            Write-Host "  Version: $($driver.DateAndVersion)" -ForegroundColor DarkYellow
            Write-Host ""
        }
        
        if ($outdatedDrivers.Count -gt 0) {
            Write-Host "`nOUTDATED DRIVERS (older than 2 years):" -ForegroundColor Red
            foreach ($driver in $outdatedDrivers) {
                Write-Host "  $($driver.OriginalName) - $($driver.DateAndVersion)" -ForegroundColor Red
            }
            Write-Host "`nOutdated drivers are a common cause of Windows upgrade failures. Consider updating them." -ForegroundColor Red
        }
    }
    else {
        Write-Host "No potentially problematic drivers identified in common categories." -ForegroundColor Green
    }
    
    # Check GPU drivers specifically
    Write-Host "`nChecking graphics card drivers (common upgrade blockers)..." -ForegroundColor Cyan
    
    $graphicsDrivers = Get-WmiObject Win32_VideoController
    foreach ($gpu in $graphicsDrivers) {
        Write-Host "  GPU: $($gpu.Name)" -ForegroundColor White
        Write-Host "  Driver Version: $($gpu.DriverVersion)" -ForegroundColor White
        Write-Host "  Driver Date: $([System.Management.ManagementDateTimeConverter]::ToDateTime($gpu.DriverDate))" -ForegroundColor White
        
        # Check for Intel, NVIDIA or AMD to provide specific guidance
        if ($gpu.Name -match "NVIDIA") {
            Write-Host "  Recommendation: Check for updated NVIDIA drivers at https://www.nvidia.com/Download/index.aspx" -ForegroundColor Cyan
        }
        elseif ($gpu.Name -match "AMD|Radeon") {
            Write-Host "  Recommendation: Check for updated AMD drivers at https://www.amd.com/en/support" -ForegroundColor Cyan
        }
        elseif ($gpu.Name -match "Intel") {
            Write-Host "  Recommendation: Check for updated Intel drivers at https://www.intel.com/content/www/us/en/download-center/home.html" -ForegroundColor Cyan
        }
        
        Write-Host ""
    }
    
    # Final recommendations
    Write-Host "`nDRIVER RECOMMENDATIONS:" -ForegroundColor Cyan
    if ($problemDevices.Count -gt 0 -or $outdatedDrivers.Count -gt 0) {
        Write-Host "Your system has potential driver issues that could be blocking the Windows 11 upgrade." -ForegroundColor Yellow
        Write-Host "1. Update drivers for devices showing errors or warnings" -ForegroundColor Yellow
        Write-Host "2. Focus on graphics, network, and storage drivers first" -ForegroundColor Yellow
        Write-Host "3. Visit your PC manufacturer's website for the latest driver packages" -ForegroundColor Yellow
    }
    else {
        Write-Host "Your drivers appear to be in good condition, but some may still be incompatible with Windows 11 24H2." -ForegroundColor White
        Write-Host "Since Microsoft detected a driver or service issue, consider updating critical drivers anyway." -ForegroundColor White
    }
}

function Perform-DeepApplicationScan {
    Write-LogSection "Deep Application Compatibility Scan"
    
    Write-Host "Scanning for applications that might be incompatible with Windows 11..." -ForegroundColor Cyan
    Write-Host "This may take a few minutes..." -ForegroundColor DarkGray
    
    # Get installed applications using different methods for thoroughness
    $apps = @()
    
    # Method 1: Traditional registry method
    Write-Host "Scanning registry for installed applications..." -ForegroundColor DarkGray
    $uninstallKeys = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $uninstallKeys += @(
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
    }
    
    foreach ($key in $uninstallKeys) {
        $apps += Get-ItemProperty $key -ErrorAction SilentlyContinue | 
            Where-Object { $_.DisplayName -and $_.UninstallString } | 
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
    }
    
    Write-Host "Found $($apps.Count) installed applications." -ForegroundColor White
    
    # Known problematic applications
    $knownIssueApps = @(
        @{Name = "VPN"; Pattern = "VPN|Cisco|OpenVPN|NordVPN|ExpressVPN|Pulse|AnyConnect|FortiClient"; Reason = "VPN software often includes network filter drivers that can interfere with Windows upgrades." },
        @{Name = "Antivirus"; Pattern = "Norton|McAfee|Kaspersky|Avast|AVG|Bitdefender|ESET|Trend Micro|Webroot|F-Secure|Sophos"; Reason = "Security software can block Windows system file modifications during upgrade." },
        @{Name = "Disk Encryption"; Pattern = "BitLocker|VeraCrypt|DiskCryptor|TrueCrypt|PGP|Symantec Encryption"; Reason = "Disk encryption can prevent Windows from accessing required drive sectors during upgrade." },
        @{Name = "System Utilities"; Pattern = "CCleaner|Advanced SystemCare|System Mechanic|Wise|TuneUp|Glary|PC Cleaner|Driver Booster"; Reason = "System optimization tools can modify Windows settings that interfere with upgrades." },
        @{Name = "Legacy Software"; Pattern = "2000|2003|XP|Vista|7|2008"; Reason = "Very old software may use deprecated APIs or services not compatible with Windows 11." },
        @{Name = "Virtualization"; Pattern = "VMware|VirtualBox|Hyper-V|Virtual PC|Parallels|QEMU|Xen"; Reason = "Virtualization software uses low-level system drivers that might be incompatible." },
        @{Name = "Audio/Video"; Pattern = "ASIO|Realtek HD|Sound Blaster|NVIDIA HD Audio|AMD High Definition|Dolby|Conexant"; Reason = "Audio drivers can be problematic for Windows 11 upgrades." },
        @{Name = "Remote Access"; Pattern = "TeamViewer|AnyDesk|Remote Desktop|LogMeIn|Chrome Remote|UltraVNC|RealVNC|TightVNC"; Reason = "Remote access tools insert hooks into Windows that might interfere with upgrades." },
        @{Name = "Custom Hardware"; Pattern = "Logitech|Razer|Corsair|SteelSeries|ASUS ROG|Alienware|MSI|Gigabyte"; Reason = "Custom hardware control software often includes drivers that may be incompatible." }
    )
    
    $potentialBlockers = @()
    
    # Check installed apps against known patterns
    foreach ($app in $apps) {
        foreach ($knownIssue in $knownIssueApps) {
            if ($app.DisplayName -match $knownIssue.Pattern) {
                $potentialBlockers += [PSCustomObject]@{
                    Name = $app.DisplayName
                    Version = $app.DisplayVersion
                    Publisher = $app.Publisher
                    Category = $knownIssue.Name
                    Reason = $knownIssue.Reason
                }
                break # Found a match, no need to check other patterns for this app
            }
        }
    }
    
    if ($potentialBlockers.Count -gt 0) {
        Write-Host "`nPOTENTIALLY INCOMPATIBLE APPLICATIONS DETECTED:" -ForegroundColor Yellow
        $potentialBlockers | ForEach-Object {
            Write-Host "  Application: $($_.Name)" -ForegroundColor Yellow
            Write-Host "  Version: $($_.Version)" -ForegroundColor DarkYellow
            Write-Host "  Publisher: $($_.Publisher)" -ForegroundColor DarkYellow
            Write-Host "  Category: $($_.Category)" -ForegroundColor DarkYellow
            Write-Host "  Reason: $($_.Reason)" -ForegroundColor DarkYellow
            Write-Host ""
        }
        
        # Group by category to summarize
        $categories = $potentialBlockers | Group-Object -Property Category | 
            Sort-Object -Property Count -Descending
        
        Write-Host "`nSUMMARY OF POTENTIAL BLOCKERS:" -ForegroundColor Cyan
        foreach ($category in $categories) {
            Write-Host "  $($category.Name) software: $($category.Count) found" -ForegroundColor White
        }
        
        Write-Host "`nRECOMMENDATIONS:" -ForegroundColor Cyan
        Write-Host "1. Temporarily disable or uninstall these applications before upgrading" -ForegroundColor White
        Write-Host "2. For security software, use Safe Mode with Networking for the upgrade" -ForegroundColor White
        Write-Host "3. Check for updated versions of these programs compatible with Windows 11" -ForegroundColor White
    }
    else {
        Write-Host "No commonly problematic applications detected." -ForegroundColor Green
    }
    
    # Check for startup programs that might interfere
    Write-Host "`nChecking startup programs..." -ForegroundColor Cyan
    
    $startupItems = @()
    $startupItems += Get-CimInstance -ClassName Win32_StartupCommand
    
    if ($startupItems.Count -gt 0) {
        Write-Host "Found $($startupItems.Count) startup items." -ForegroundColor White
        Write-Host "Consider disabling non-essential startup items before upgrading." -ForegroundColor Yellow
        
        # Display some of the startup items
        Write-Host "`nKey startup items:" -ForegroundColor DarkYellow
        $startupItems | Where-Object { 
            $_.Command -match "VPN|Security|Antivirus|Firewall|Sync|Cloud|Backup|Monitor" 
        } | ForEach-Object {
            Write-Host "  $($_.Name) - $($_.Location)" -ForegroundColor DarkYellow
        }
    }
}

function Analyze-ServicesCompatibility {
    Write-LogSection "Service Compatibility Analysis"
    
    Write-Host "Analyzing system services for potential upgrade blockers..." -ForegroundColor Cyan
    
    # Services known to potentially interfere with Windows upgrades
    $problematicServices = @(
        @{Name="Windows Update"; ServiceName="wuauserv"; Description="Required for update installation"},
        @{Name="Background Intelligent Transfer"; ServiceName="bits"; Description="Required for downloading updates"},
        @{Name="Windows Installer"; ServiceName="msiserver"; Description="Required for component installation"},
        @{Name="App Readiness"; ServiceName="AppReadiness"; Description="Prepares apps for use the first time a user signs in"},
        @{Name="Windows Update Medic"; ServiceName="WaaSMedicSvc"; Description="Repairs Windows Update components"},
        @{Name="Cryptographic Services"; ServiceName="CryptSvc"; Description="Required for verifying digital signatures"}
    )
    
    # Third-party services that might cause issues
    $thirdPartyServices = Get-Service | Where-Object {
        $_.Status -eq "Running" -and
        -not $_.DisplayName.StartsWith("Windows") -and
        -not $_.DisplayName.StartsWith("Microsoft") -and
        $_.StartType -eq "Automatic"
    }
    
    # Check critical services
    $criticalIssues = @()
    foreach ($service in $problematicServices) {
        $svc = Get-Service -Name $service.ServiceName -ErrorAction SilentlyContinue
        
        if (-not $svc) {
            $criticalIssues += "Service not found: $($service.Name) ($($service.ServiceName))"
            continue
        }
        
        if ($svc.Status -ne "Running") {
            $criticalIssues += "Service not running: $($service.Name) ($($service.ServiceName))"
        }
        
        if ($svc.StartType -eq "Disabled") {
            $criticalIssues += "Service disabled: $($service.Name) ($($service.ServiceName))"
        }
    }
    
    if ($criticalIssues.Count -gt 0) {
        Write-Host "`nCRITICAL SERVICE ISSUES FOUND:" -ForegroundColor Red
        foreach ($issue in $criticalIssues) {
            Write-Host "  $issue" -ForegroundColor Red
        }
        
        Write-Host "`nThese service issues must be fixed before Windows can upgrade successfully." -ForegroundColor Red
        Write-Host "Run 'services.msc' and ensure these services are set to 'Automatic' and are 'Running'" -ForegroundColor Yellow
    }
    else {
        Write-Host "All critical Windows Update services are running correctly." -ForegroundColor Green
    }
    
    # Check for potential problematic third-party services
    $securityRelated = $thirdPartyServices | Where-Object {
        $_.DisplayName -match "Security|Antivirus|Protection|Firewall|VPN|Defense|Guard|Shield"
    }
    
    $virtualizationRelated = $thirdPartyServices | Where-Object {
        $_.DisplayName -match "VM|Virtual|Hyper|Docker|Container"
    }
    
    $networkRelated = $thirdPartyServices | Where-Object {
        $_.DisplayName -match "Network|Adapter|WiFi|Wireless|Ethernet|Broadband|Modem|RAS|VPN|Tunnel"
    }
    
    Write-Host "`nPOTENTIAL PROBLEMATIC THIRD-PARTY SERVICES:" -ForegroundColor Yellow
    
    if ($securityRelated.Count -gt 0) {
        Write-Host "`n  SECURITY SOFTWARE SERVICES:" -ForegroundColor Yellow
        $securityRelated | ForEach-Object {
            Write-Host "    $($_.DisplayName) [$($_.Name)]" -ForegroundColor DarkYellow
        }
        Write-Host "  Security software is a common cause of Windows upgrade failures." -ForegroundColor Yellow
    }
    
    if ($virtualizationRelated.Count -gt 0) {
        Write-Host "`n  VIRTUALIZATION SERVICES:" -ForegroundColor Yellow
        $virtualizationRelated | ForEach-Object {
            Write-Host "    $($_.DisplayName) [$($_.Name)]" -ForegroundColor DarkYellow
        }
        Write-Host "  Virtualization software can conflict with Windows system changes during upgrade." -ForegroundColor Yellow
    }
    
    if ($networkRelated.Count -gt 0) {
        Write-Host "`n  NETWORKING SERVICES:" -ForegroundColor Yellow
        foreach ($svc in $networkRelated | Select-Object -First 5) {
            Write-Host "    $($svc.DisplayName) [$($svc.Name)]" -ForegroundColor DarkYellow
        }
        if ($networkRelated.Count -gt 5) {
            Write-Host "    ... and $($networkRelated.Count - 5) more" -ForegroundColor DarkYellow
        }
        Write-Host "  Custom networking services can interfere with Windows setup networking components." -ForegroundColor Yellow
    }
    
    Write-Host "`nSERVICE RECOMMENDATIONS:" -ForegroundColor Cyan
    Write-Host "1. Temporarily disable third-party security services before upgrading" -ForegroundColor White
    Write-Host "2. Ensure all Microsoft update services are running" -ForegroundColor White
    Write-Host "3. Try performing the upgrade in Safe Mode with Networking if service issues persist" -ForegroundColor White
    Write-Host "4. Run this command before upgrading to reset Windows Update components:" -ForegroundColor White
    Write-Host "   net stop wuauserv && net stop bits && net stop cryptsvc && ren %systemroot%\SoftwareDistribution SoftwareDistribution.old && net start cryptsvc && net start bits && net start wuauserv" -ForegroundColor DarkGray
}

function Export-DiagnosticReport {
    param (
        [string]$ReportPath = "C:\Temp\Windows11_Upgrade_Diagnostic_Report.html"
    )
    
    Write-LogSection "Creating Diagnostic Report"
    
    # Check if destination directory exists and create it if needed
    $reportDirectory = Split-Path -Path $ReportPath -Parent
    if (-not (Test-Path -Path $reportDirectory -PathType Container)) {
        try {
            Write-Host "Creating directory: $reportDirectory" -ForegroundColor Yellow
            New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
            Write-Host "Directory created successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to create directory $reportDirectory. Error: $_" -ForegroundColor Red
            # Fall back to desktop if we can't create the requested directory
            $ReportPath = "$env:USERPROFILE\Desktop\Windows11_Upgrade_Diagnostic_Report.html"
            Write-Host "Falling back to desktop: $ReportPath" -ForegroundColor Yellow
        }
    }
    
    # Create HTML header
    $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows 11 Upgrade Diagnostic Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0078d4; }
        h2 { color: #0078d4; margin-top: 20px; border-bottom: 1px solid #ddd; padding-bottom: 5px; }
        .error { color: #c00; }
        .warning { color: #c50; }
        .success { color: #080; }
        .info { color: #555; }
        .section { margin: 10px 0; padding: 10px; background: #f8f8f8; border-left: 4px solid #0078d4; }
        .recommendation { margin: 10px 0; padding: 10px; background: #e6f0ff; border-left: 4px solid #0078d4; }
    </style>
</head>
<body>
    <h1>Windows 11 Upgrade Diagnostic Report</h1>
    <p>Generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
"@
    
    # Create HTML content sections (these would be populated by capturing output from various functions)
    $htmlContent = "<div class='section'><h2>System Information</h2>"
    $htmlContent += "<p>This report identifies potential issues preventing your Windows 11 23H2 to 24H2 upgrade.</p>"
    $htmlContent += "</div>"
    
    # Create recommendations section
    $htmlRecommendations = "<div class='recommendation'><h2>Recommendations</h2>"
    $htmlRecommendations += "<p class='warning'>Your PC has a driver or service that isn't ready for Windows 11 24H2.</p>"
    $htmlRecommendations += "<ul>"
    $htmlRecommendations += "<li>Update all drivers, especially graphics, network, and storage drivers</li>"
    $htmlRecommendations += "<li>Temporarily disable security software during the upgrade</li>"
    $htmlRecommendations += "<li>Disconnect non-essential peripherals before upgrading</li>"
    $htmlRecommendations += "<li>Try the Media Creation Tool instead of Windows Update</li>"
    $htmlRecommendations += "</ul>"
    $htmlRecommendations += "</div>"
    
    # Create HTML footer
    $htmlFooter = @"
    <p style="margin-top:30px; font-size:12px; color:#777;">
    This report was generated by the Windows 11 Upgrade Diagnostic Tool.<br>
    For more information, visit <a href="https://go.microsoft.com/fwlink/?LinkId=2280120">Microsoft's Windows 11 Compatibility page</a>.
    </p>
</body>
</html>
"@
    
    # Combine all HTML elements
    $htmlReport = $htmlHeader + $htmlContent + $htmlRecommendations + $htmlFooter
    
    # Write HTML to file
    try {
        $htmlReport | Out-File -FilePath $ReportPath -Encoding UTF8 -Force
        Write-Host "Diagnostic report saved to: $ReportPath" -ForegroundColor Green
        Write-Host "Open this HTML file in your browser for a formatted report with recommendations." -ForegroundColor White
        
        # Try to open the report automatically
        if ([System.Environment]::OSVersion.Platform -eq "Win32NT") {
            Start-Process $ReportPath -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Host "Failed to create diagnostic report: $_" -ForegroundColor Red
        
        # Try one more time with a report on the desktop
        try {
            $desktopPath = "$env:USERPROFILE\Desktop\Windows11_Upgrade_Diagnostic_Report.html"
            $htmlReport | Out-File -FilePath $desktopPath -Encoding UTF8 -Force
            Write-Host "Diagnostic report saved to desktop instead: $desktopPath" -ForegroundColor Yellow
            
            if ([System.Environment]::OSVersion.Platform -eq "Win32NT") {
                Start-Process $desktopPath -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Host "Could not save report anywhere. Make sure you have write permissions." -ForegroundColor Red
        }
    }
}

# Main menu execution with the complete loop
function Start-DiagnosticTool {
    $exit = $false
    
    while (-not $exit) {
        Show-Menu
        $choice = Read-Host "Enter your choice (1-3)"
        
        switch ($choice) {
            1 {
                Write-Host "Starting deep scan..." -ForegroundColor Green
                Get-CurrentWindowsVersion
                Test-SystemRequirements
                Check-DiskSpace
                Process-ScanResultXML
                Perform-DeepDriverScan
                Perform-DeepApplicationScan
                Analyze-ServicesCompatibility
                Check-PendingUpdates
                Get-UpdateBlockerRecommendations
                Export-DiagnosticReport
                
                Write-Host "`nDeep scan completed. Review the information above for potential issues." -ForegroundColor Green
                Write-Host "Press Enter to return to the main menu..." -ForegroundColor Cyan
                Read-Host | Out-Null
            }
            2 {
                Write-Host "Starting setup diagnosis..." -ForegroundColor Yellow
                Get-CurrentWindowsVersion
                Test-SystemRequirements
                Check-DiskSpace
                Process-ScanResultXML
                Check-PendingUpdates
                Check-IncompatibleSoftware
                Check-WindowsServicesStatus
                Get-SetupLogs
                Get-UpdateBlockerRecommendations
                
                Write-Host "`nSetup diagnosis completed. Press Enter to return to the main menu..." -ForegroundColor Yellow
                Read-Host | Out-Null
            }
            3 {
                Write-Host "Exiting the diagnostic tool. Goodbye!" -ForegroundColor Red
                $exit = $true
            }
            default {
                Write-Host "Invalid choice. Please enter a number between 1 and 3." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    }
}

# Start the tool
Start-DiagnosticTool
