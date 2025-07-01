param(
    [Parameter(Mandatory=$false)]
    [string]$OSTPath,
    
    [Parameter(Mandatory=$false)]
    [string]$PSTPath
)

# Function to close any existing Outlook processes
function Close-OutlookProcesses {
    try {
        Write-Host "Checking for existing Outlook processes..." -ForegroundColor Yellow
        $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        
        if ($outlookProcesses) {
            Write-Host "Found $($outlookProcesses.Count) Outlook process(es). Closing them..." -ForegroundColor Yellow
            foreach ($process in $outlookProcesses) {
                try {
                    $process.CloseMainWindow()
                    Start-Sleep -Seconds 2
                    if (-not $process.HasExited) {
                        $process.Kill()
                    }
                    Write-Host "Closed Outlook process (PID: $($process.Id))" -ForegroundColor Green
                }
                catch {
                    Write-Warning "Could not close Outlook process (PID: $($process.Id)): $($_.Exception.Message)"
                }
            }
            # Wait a moment for processes to fully close
            Start-Sleep -Seconds 3
        } else {
            Write-Host "No existing Outlook processes found." -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Error checking for Outlook processes: $($_.Exception.Message)"
    }
}

# Function to check if Outlook is installed
function Test-OutlookInstalled {
    try {
        Write-Host "Checking Outlook installation..." -ForegroundColor Yellow
        
        # First check if Outlook is registered as a COM object
        $outlookProgId = "Outlook.Application"
        $outlookType = [Type]::GetTypeFromProgID($outlookProgId)
        if (-not $outlookType) {
            Write-Host "Outlook COM object not found in registry." -ForegroundColor Red
            return $false
        }
        
        # Create a job to test Outlook with timeout
        $job = Start-Job -ScriptBlock {
            try {
                $outlook = New-Object -ComObject Outlook.Application
                if ($outlook) {
                    $outlook.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
                    return $true
                }
                return $false
            }
            catch {
                return $false
            }
        }
        
        # Wait for job completion with timeout (30 seconds)
        $timeout = 30
        Write-Host "Testing Outlook COM interface (timeout: $timeout seconds)..." -ForegroundColor Yellow
        
        if (Wait-Job -Job $job -Timeout $timeout) {
            $result = Receive-Job -Job $job
            Remove-Job -Job $job -Force
            
            if ($result) {
                Write-Host "Outlook installation verified." -ForegroundColor Green
                return $true
            } else {
                Write-Host "Outlook COM interface test failed." -ForegroundColor Red
                return $false
            }
        } else {
            Write-Host "Outlook COM interface test timed out." -ForegroundColor Red
            Remove-Job -Job $job -Force
            # Close any Outlook processes that might be hanging
            Close-OutlookProcesses
            return $false
        }
    }
    catch {
        Write-Host "Error checking Outlook installation: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to get folder item count recursively
function Get-FolderItemCount {
    param($folder)
    $count = $folder.Items.Count
    foreach ($subfolder in $folder.Folders) {
        $count += Get-FolderItemCount -folder $subfolder
    }
    return $count
}

# Function to copy folder contents recursively
function Copy-FolderContents {
    param(
        $sourceFolder,
        $destinationFolder,
        [ref]$processedItems,
        $totalItems
    )
    
    # Copy items in current folder
    $items = $sourceFolder.Items
    for ($i = 1; $i -le $items.Count; $i++) {
        try {
            $item = $items.Item($i)
            $copiedItem = $item.Copy()
            $copiedItem.Move($destinationFolder)
            $processedItems.Value++
            
            $percentComplete = [math]::Round(($processedItems.Value / $totalItems) * 100, 2)
            $currentStatus = "Processing item $($processedItems.Value) of $totalItems - $($item.Subject)"
            if ($currentStatus.Length -gt 80) {
                $currentStatus = $currentStatus.Substring(0, 77) + "..."
            }
            Write-Progress -Activity "Converting OST to PST" -Status $currentStatus -PercentComplete $percentComplete -CurrentOperation "Copying items from folder: $($sourceFolder.Name)"
        }
        catch {
            Write-Warning "Failed to copy item: $($_.Exception.Message)"
        }
    }
    
    # Process subfolders recursively
    foreach ($subfolder in $sourceFolder.Folders) {
        try {
            $newFolder = $destinationFolder.Folders.Add($subfolder.Name, $subfolder.DefaultItemType)
            Copy-FolderContents -sourceFolder $subfolder -destinationFolder $newFolder -processedItems $processedItems -totalItems $totalItems
        }
        catch {
            Write-Warning "Failed to process folder '$($subfolder.Name)': $($_.Exception.Message)"
        }
    }
}

# Main script execution
try {
    Write-Host "OST to PST Converter" -ForegroundColor Green
    Write-Host "===================" -ForegroundColor Green
    
    # Phase 1: Initial setup
    Write-Progress -Activity "OST to PST Conversion" -Status "Initializing..." -PercentComplete 0
    
    # Close any existing Outlook processes first
    Close-OutlookProcesses
    
    # Check if Outlook is installed
    Write-Progress -Activity "OST to PST Conversion" -Status "Checking Outlook installation..." -PercentComplete 5
    if (-not (Test-OutlookInstalled)) {
        throw "Microsoft Outlook is not installed or not accessible via COM interface."
    }
    
    # Get OST file path if not provided
    if (-not $OSTPath) {
        do {
            $OSTPath = Read-Host "Enter the path to the OST file"
            # Clean up the path - remove surrounding quotes if present
            $OSTPath = $OSTPath.Trim('"').Trim("'").Trim()
            
            Write-Host "Testing path: '$OSTPath'" -ForegroundColor Cyan
            
            if (-not (Test-Path $OSTPath)) {
                Write-Host "File not found. Please enter a valid path." -ForegroundColor Red
                Write-Host "Make sure the file exists and you have access to it." -ForegroundColor Yellow
                $OSTPath = $null
            }
            elseif (-not $OSTPath.EndsWith('.ost', [System.StringComparison]::OrdinalIgnoreCase)) {
                Write-Host "Please specify a valid OST file (must have .ost extension)." -ForegroundColor Red
                $OSTPath = $null
            }
        } while (-not $OSTPath)
    }
    else {
        # Clean up the provided path
        $OSTPath = $OSTPath.Trim('"').Trim("'").Trim()
        Write-Host "Using provided OST path: '$OSTPath'" -ForegroundColor Cyan
    }
    
    # Validate OST file exists
    if (-not (Test-Path $OSTPath)) {
        Write-Host "Tested path: '$OSTPath'" -ForegroundColor Red
        throw "OST file not found: $OSTPath"
    }
    
    # Validate it's an OST file
    if (-not $OSTPath.EndsWith('.ost', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Invalid file type. Please specify an OST file."
    }
    
    # Get PST file path if not provided
    if (-not $PSTPath) {
        $defaultPSTPath = [System.IO.Path]::ChangeExtension($OSTPath, ".pst")
        $PSTPath = Read-Host "Enter the path for the output PST file (default: $defaultPSTPath)"
        if ([string]::IsNullOrWhiteSpace($PSTPath)) {
            $PSTPath = $defaultPSTPath
        }
    }
    
    # Check if PST file already exists
    if (Test-Path $PSTPath) {
        $overwrite = Read-Host "PST file already exists. Overwrite? (y/n)"
        if ($overwrite -ne 'y') {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            exit
        }
        Remove-Item $PSTPath -Force
    }
    
    Write-Host "Starting conversion..." -ForegroundColor Yellow
    Write-Host "Source OST: $OSTPath" -ForegroundColor Cyan
    Write-Host "Target PST: $PSTPath" -ForegroundColor Cyan
    
    # Phase 3: Initialize Outlook
    Write-Progress -Activity "OST to PST Conversion" -Status "Initializing Outlook COM interface..." -PercentComplete 20
    
    # Create Outlook application object
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    # Phase 4: Loading OST file
    Write-Progress -Activity "OST to PST Conversion" -Status "Loading OST file..." -PercentComplete 25
    Write-Host "Adding OST file to Outlook profile..." -ForegroundColor Yellow
    $namespace.AddStore($OSTPath)
    
    # Wait for store to be available with progress indication
    for ($i = 1; $i -le 3; $i++) {
        Write-Progress -Activity "OST to PST Conversion" -Status "Waiting for OST file to load... ($i/3)" -PercentComplete (25 + ($i * 2))
        Start-Sleep -Seconds 1
    }
    
    # Find the added store
    $ostStore = $null
    foreach ($store in $namespace.Stores) {
        if ($store.FilePath -eq $OSTPath) {
            $ostStore = $store
            break
        }
    }
    
    if (-not $ostStore) {
        throw "Could not access OST file. Make sure the file is not corrupted and not currently in use."
    }
    
    # Phase 5: Creating PST file
    Write-Progress -Activity "OST to PST Conversion" -Status "Creating PST file..." -PercentComplete 35
    Write-Host "Creating PST file..." -ForegroundColor Yellow
    $namespace.AddStore($PSTPath)
    
    # Wait for PST to be created with progress indication
    for ($i = 1; $i -le 2; $i++) {
        Write-Progress -Activity "OST to PST Conversion" -Status "Waiting for PST file creation... ($i/2)" -PercentComplete (35 + ($i * 2))
        Start-Sleep -Seconds 1
    }
    
    # Find the PST store
    $pstStore = $null
    foreach ($store in $namespace.Stores) {
        if ($store.FilePath -eq $PSTPath) {
            $pstStore = $store
            break
        }
    }
    
    if (-not $pstStore) {
        throw "Could not create PST file."
    }
    
    # Phase 6: Analyzing data structure
    Write-Progress -Activity "OST to PST Conversion" -Status "Analyzing data structure..." -PercentComplete 40
    
    # Get root folders
    $ostRootFolder = $ostStore.GetRootFolder()
    $pstRootFolder = $pstStore.GetRootFolder()
    
    # Calculate total items for progress tracking
    Write-Host "Calculating total items..." -ForegroundColor Yellow
    $totalItems = Get-FolderItemCount -folder $ostRootFolder
    Write-Host "Total items to process: $totalItems" -ForegroundColor Cyan
    
    # Phase 7: Data conversion (45% to 95%)
    Write-Progress -Activity "OST to PST Conversion" -Status "Starting data conversion..." -PercentComplete 45
    
    # Copy all folders and items
    $processedItems = [ref]0
    Copy-FolderContents -sourceFolder $ostRootFolder -destinationFolder $pstRootFolder -processedItems $processedItems -totalItems $totalItems
    
    # Phase 8: Cleanup
    Write-Progress -Activity "OST to PST Conversion" -Status "Finalizing conversion..." -PercentComplete 95
    
    # Remove OST from profile
    $namespace.RemoveStore($ostRootFolder)
    
    Write-Progress -Activity "OST to PST Conversion" -Status "Conversion completed!" -PercentComplete 100
    Start-Sleep -Seconds 1
    Write-Progress -Activity "OST to PST Conversion" -Completed
    
    Write-Host "Conversion completed successfully!" -ForegroundColor Green
    Write-Host "PST file created: $PSTPath" -ForegroundColor Green
    Write-Host "Total items processed: $($processedItems.Value)" -ForegroundColor Cyan
}
catch {
    Write-Progress -Activity "OST to PST Conversion" -Completed
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    # Cleanup COM objects
    if ($namespace) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
    }
    if ($outlook) {
        $outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
    
    # Force garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "Script execution completed." -ForegroundColor Green
