param(
    [Parameter(Mandatory=$false)]
    [string]$OSTPath,
    
    [Parameter(Mandatory=$false)]
    [string]$PSTPath
)

# Function to check if Outlook is installed
function Test-OutlookInstalled {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        return $true
    }
    catch {
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
    
    # Check if Outlook is installed
    Write-Progress -Activity "OST to PST Conversion" -Status "Checking Outlook installation..." -PercentComplete 5
    if (-not (Test-OutlookInstalled)) {
        throw "Microsoft Outlook is not installed or not accessible via COM interface."
    }
    
    # Get OST file path if not provided
    if (-not $OSTPath) {
        do {
            $OSTPath = Read-Host "Enter the path to the OST file"
            if (-not (Test-Path $OSTPath)) {
                Write-Host "File not found. Please enter a valid path." -ForegroundColor Red
                $OSTPath = $null
            }
        } while (-not $OSTPath)
    }
    
    # Validate OST file exists
    if (-not (Test-Path $OSTPath)) {
        throw "OST file not found: $OSTPath"
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
