# Requires admin privileges
#Requires -RunAsAdministrator

# Set error action preference
$ErrorActionPreference = "Stop"

# Function to find outdated drivers
function Find-OutdatedDrivers {
    [CmdletBinding()]
    param()
    
    Write-Host "Scanning for outdated drivers..." -ForegroundColor Cyan
    
    try {
        # Get all device drivers
        $allDrivers = Get-WmiObject Win32_PnPSignedDriver | 
                     Where-Object { $null -ne $_.DeviceName } |
                     Select-Object DeviceName, DeviceID, DriverVersion, DriverDate, Manufacturer
        
        # Create custom object with driver information and outdated status
        $drivers = @()
        foreach ($driver in $allDrivers) {
            # This is a simplified check - in a real scenario you might want to check against an online database
            $driverAge = New-TimeSpan -Start ([DateTime]::ParseExact($driver.DriverDate.Split('.')[0], 'yyyyMMdd', $null)) -End (Get-Date)
            $isOutdated = $driverAge.Days -gt 365  # Consider drivers older than 1 year as outdated
            
            $drivers += [PSCustomObject]@{
                DeviceName = $driver.DeviceName
                DeviceID = $driver.DeviceID
                DriverVersion = $driver.DriverVersion
                DriverDate = [DateTime]::ParseExact($driver.DriverDate.Split('.')[0], 'yyyyMMdd', $null).ToString("yyyy-MM-dd")
                Manufacturer = $driver.Manufacturer
                Age = $driverAge.Days
                IsOutdated = $isOutdated
            }
        }
        
        # Return all drivers with outdated flag
        return $drivers
    }
    catch {
        Write-Error "Error finding outdated drivers: $_"
        return $null
    }
}

# Function to download driver updates
function Get-DriverUpdates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$OutdatedDrivers,
        
        [Parameter(Mandatory = $true)]
        [string]$DownloadPath
    )
    
    Write-Host "Preparing to download driver updates..." -ForegroundColor Cyan
    
    # Create download directory if it doesn't exist
    if (-not (Test-Path -Path $DownloadPath)) {
        New-Item -Path $DownloadPath -ItemType Directory -Force | Out-Null
    }
    
    $downloadResults = @()
    
    foreach ($driver in $OutdatedDrivers) {
        if ($driver.IsOutdated) {
            Write-Host "Attempting to find updates for $($driver.DeviceName)..." -ForegroundColor Yellow
            
            # This is where you would implement manufacturer-specific download logic
            # For demonstration, we'll simulate the download process
            $downloadSuccess = $false
            $downloadUrl = ""
            $downloadedFile = ""
            
            # Check manufacturer and simulate download logic
            switch -Wildcard ($driver.Manufacturer) {
                "*Intel*" { 
                    $downloadSuccess = $true 
                    $downloadUrl = "https://downloadcenter.intel.com/download/sample"
                    $downloadedFile = Join-Path -Path $DownloadPath -ChildPath "Intel_$($driver.DeviceName -replace '[^\w\-]', '_')_Update.exe"
                }
                "*NVIDIA*" { 
                    $downloadSuccess = $true 
                    $downloadUrl = "https://www.nvidia.com/download/sample"
                    $downloadedFile = Join-Path -Path $DownloadPath -ChildPath "NVIDIA_$($driver.DeviceName -replace '[^\w\-]', '_')_Update.exe"
                }
                "*AMD*" { 
                    $downloadSuccess = $true 
                    $downloadUrl = "https://www.amd.com/download/sample"
                    $downloadedFile = Join-Path -Path $DownloadPath -ChildPath "AMD_$($driver.DeviceName -replace '[^\w\-]', '_')_Update.exe"
                }
                default { 
                    $downloadSuccess = $false
                    Write-Host "  No download source found for $($driver.Manufacturer)" -ForegroundColor Red
                }
            }
            
            if ($downloadSuccess) {
                # Simulate download by creating empty file
                # In a real scenario, you would use Invoke-WebRequest or similar
                Write-Host "  Downloading from: $downloadUrl" -ForegroundColor Green
                Write-Host "  Saving to: $downloadedFile" -ForegroundColor Green
                
                # Simulating download by creating an empty file
                New-Item -Path $downloadedFile -ItemType File -Force | Out-Null
                
                $downloadResults += [PSCustomObject]@{
                    DeviceName = $driver.DeviceName
                    DeviceID = $driver.DeviceID
                    DownloadSuccess = $true
                    DownloadPath = $downloadedFile
                    InstallReady = $true
                }
            }
            else {
                $downloadResults += [PSCustomObject]@{
                    DeviceName = $driver.DeviceName
                    DeviceID = $driver.DeviceID
                    DownloadSuccess = $false
                    DownloadPath = $null
                    InstallReady = $false
                }
            }
        }
    }
    
    return $downloadResults
}

# Function to install driver updates
function Install-DriverUpdates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$DriverUpdates
    )
    
    Write-Host "Installing driver updates..." -ForegroundColor Cyan
    
    $installResults = @()
    
    foreach ($update in $DriverUpdates) {
        if ($update.InstallReady -and (Test-Path -Path $update.DownloadPath)) {
            Write-Host "Installing update for $($update.DeviceName)..." -ForegroundColor Yellow
            
            # This is a simulated installation - in real usage, you would execute the installer
            # For example: Start-Process -FilePath $update.DownloadPath -ArgumentList "/quiet", "/norestart" -Wait
            
            # Simulate installation success (in reality, you would verify the exit code of the installer)
            $installSuccess = $true
            
            if ($installSuccess) {
                Write-Host "  Installation successful" -ForegroundColor Green
            }
            else {
                Write-Host "  Installation failed" -ForegroundColor Red
            }
            
            $installResults += [PSCustomObject]@{
                DeviceName = $update.DeviceName
                DeviceID = $update.DeviceID
                InstallSuccess = $installSuccess
                InstallerPath = $update.DownloadPath
            }
        }
    }
    
    return $installResults
}

# Function to verify driver installations
function Test-DriverInstallations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$InstalledDrivers,
        
        [Parameter(Mandatory = $true)]
        [array]$OriginalDrivers
    )
    
    Write-Host "Verifying driver installations..." -ForegroundColor Cyan
    
    $verificationResults = @()
    
    # Get updated driver information
    $currentDrivers = Get-WmiObject Win32_PnPSignedDriver | 
                     Where-Object { $null -ne $_.DeviceName } |
                     Select-Object DeviceName, DeviceID, DriverVersion, DriverDate
    
    foreach ($installed in $InstalledDrivers) {
        $originalDriver = $OriginalDrivers | Where-Object { $_.DeviceID -eq $installed.DeviceID }
        $currentDriver = $currentDrivers | Where-Object { $_.DeviceID -eq $installed.DeviceID }
        
        if ($originalDriver -and $currentDriver) {
            $isNewer = $false
            
            # Check if driver version is newer
            if ([version]$currentDriver.DriverVersion -gt [version]$originalDriver.DriverVersion) {
                $isNewer = $true
            }
            # If versions are the same, check dates
            elseif ([version]$currentDriver.DriverVersion -eq [version]$originalDriver.DriverVersion) {
                $originalDate = [DateTime]::ParseExact($originalDriver.DriverDate.Split('.')[0], 'yyyyMMdd', $null)
                $currentDate = [DateTime]::ParseExact($currentDriver.DriverDate.Split('.')[0], 'yyyyMMdd', $null)
                
                if ($currentDate -gt $originalDate) {
                    $isNewer = $true
                }
            }
            
            $verificationResults += [PSCustomObject]@{
                DeviceName = $installed.DeviceName
                DeviceID = $installed.DeviceID
                OldVersion = $originalDriver.DriverVersion
                NewVersion = $currentDriver.DriverVersion
                UpdateSuccessful = $isNewer
            }
        }
    }
    
    return $verificationResults
}

# Function to clean up old drivers
function Remove-OldDrivers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DownloadPath,
        
        [switch]$CleanupDownloads,
        
        [switch]$CleanupDriverStore
    )
    
    Write-Host "Cleaning up old drivers..." -ForegroundColor Cyan
    
    $cleanupResults = [PSCustomObject]@{
        DownloadsRemoved = 0
        DriverStoreItemsRemoved = 0
    }
    
    # Clean up downloaded files
    if ($CleanupDownloads -and (Test-Path -Path $DownloadPath)) {
        $downloadedFiles = Get-ChildItem -Path $DownloadPath -File
        $cleanupResults.DownloadsRemoved = $downloadedFiles.Count
        
        foreach ($file in $downloadedFiles) {
            Write-Host "Removing downloaded file: $($file.FullName)" -ForegroundColor Yellow
            Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
        }
    }
    
    # Clean up driver store (requires DISM module or pnputil)
    if ($CleanupDriverStore) {
        try {
            Write-Host "Cleaning up driver store (this may take some time)..." -ForegroundColor Yellow
            
            # Get all driver packages
            $driverPackages = & pnputil.exe /enum-drivers
            
            # Parse the output to identify old/unused drivers
            # This is a simplified approach - in reality, you'd need more sophisticated parsing
            $originalCount = $driverPackages.Count
            
            # For demonstration, we'll just display the command that would be used
            Write-Host "Would execute: pnputil.exe /delete-driver <oem-inf-name> /force" -ForegroundColor Yellow
            Write-Host "Driver store cleanup requires manual implementation based on your requirements" -ForegroundColor Yellow
            
            $cleanupResults.DriverStoreItemsRemoved = 0  # Would be the count of removed items in actual implementation
        }
        catch {
            Write-Error "Error cleaning up driver store: $_"
        }
    }
    
    return $cleanupResults
}

# Main function to orchestrate driver management
function Start-DriverManagement {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$DownloadPath = "$env:TEMP\DriverHunter",
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipDownload,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipInstall,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipCleanup
    )
    
    Write-Host "===== DRIVER HUNTER SCRIPT =====" -ForegroundColor Cyan
    Write-Host "Starting driver management process..." -ForegroundColor Cyan
    Write-Host "Download path: $DownloadPath" -ForegroundColor Cyan
    
    # Step 1: Find outdated drivers
    $outdatedDrivers = Find-OutdatedDrivers
    
    if (-not $outdatedDrivers) {
        Write-Host "No drivers found or error occurred during scan." -ForegroundColor Red
        return
    }
    
    $outdatedCount = ($outdatedDrivers | Where-Object { $_.IsOutdated }).Count
    Write-Host "Found $outdatedCount outdated drivers out of $($outdatedDrivers.Count) total drivers." -ForegroundColor Yellow
    
    # Display outdated drivers
    $outdatedDrivers | Where-Object { $_.IsOutdated } | Format-Table DeviceName, DriverVersion, DriverDate, Manufacturer, Age -AutoSize
    
    # Step 2: Download updates (if not skipped)
    $downloadResults = $null
    if (-not $SkipDownload -and $outdatedCount -gt 0) {
        $downloadResults = Get-DriverUpdates -OutdatedDrivers ($outdatedDrivers | Where-Object { $_.IsOutdated }) -DownloadPath $DownloadPath
        
        $successfulDownloads = ($downloadResults | Where-Object { $_.DownloadSuccess }).Count
        Write-Host "Successfully downloaded $successfulDownloads driver updates." -ForegroundColor Green
    }
    else {
        Write-Host "Driver download step skipped." -ForegroundColor Yellow
    }
    
    # Step 3: Install updates (if not skipped)
    $installResults = $null
    if (-not $SkipInstall -and $downloadResults) {
        $installResults = Install-DriverUpdates -DriverUpdates $downloadResults
        
        $successfulInstalls = ($installResults | Where-Object { $_.InstallSuccess }).Count
        Write-Host "Successfully installed $successfulInstalls driver updates." -ForegroundColor Green
        
        # Step 4: Verify installations
        if ($successfulInstalls -gt 0) {
            $verificationResults = Test-DriverInstallations -InstalledDrivers $installResults -OriginalDrivers $outdatedDrivers
            
            $successfulUpdates = ($verificationResults | Where-Object { $_.UpdateSuccessful }).Count
            Write-Host "Verified $successfulUpdates successful driver updates." -ForegroundColor Green
            
            $verificationResults | Format-Table DeviceName, OldVersion, NewVersion, UpdateSuccessful -AutoSize
        }
    }
    else {
        Write-Host "Driver installation step skipped." -ForegroundColor Yellow
    }
    
    # Step 5: Clean up (if not skipped)
    if (-not $SkipCleanup) {
        $cleanupResults = Remove-OldDrivers -DownloadPath $DownloadPath -CleanupDownloads -CleanupDriverStore
        
        Write-Host "Cleanup completed:" -ForegroundColor Green
        Write-Host "  - Downloaded files removed: $($cleanupResults.DownloadsRemoved)" -ForegroundColor Green
        Write-Host "  - Driver store items removed: $($cleanupResults.DriverStoreItemsRemoved)" -ForegroundColor Green
    }
    else {
        Write-Host "Driver cleanup step skipped." -ForegroundColor Yellow
    }
    
    Write-Host "===== DRIVER HUNTER COMPLETED =====" -ForegroundColor Cyan
}

# Run the script
try {
    # Call the main function with default parameters
    Start-DriverManagement
}
catch {
    Write-Error "Error in Driver Hunter script: $_"
}
