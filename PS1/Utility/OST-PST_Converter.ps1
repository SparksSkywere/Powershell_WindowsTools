<#
.SYNOPSIS
    Simple OST to PST Converter

.DESCRIPTION
    Converts any OST file to PST format using Outlook COM interface.
    Simplified approach that just gets the job done.

.PARAMETER OSTPath
    Path to the OST file to convert

.PARAMETER PSTPath
    Path for the output PST file

.EXAMPLE
    .\OST-PST_Converter.ps1 -OSTPath "C:\path\to\file.ost" -PSTPath "C:\path\to\output.pst"
#>

param(
    [string]$OSTPath,
    [string]$PSTPath
)

# Helper function to clean up Outlook processes
function Stop-OutlookProcesses {
    Write-Host "Stopping Outlook processes..." -ForegroundColor Yellow
    Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Host "  Stopping PID: $($_.Id)" -ForegroundColor Cyan
        $_.Kill()
    }
    Start-Sleep -Seconds 3
}

# Helper function to wait for file operations
function Wait-ForOperation {
    param([scriptblock]$Operation, [string]$Description, [int]$TimeoutMinutes = 10)
    
    Write-Host "Waiting for: $Description" -ForegroundColor Cyan
    $timeout = (Get-Date).AddMinutes($TimeoutMinutes)
    
    while ((Get-Date) -lt $timeout) {
        try {
            $result = & $Operation
            if ($result) { return $result }
        } catch {
            # Continue waiting
        }
        Start-Sleep -Seconds 5
        Write-Host "." -NoNewline -ForegroundColor Gray
    }
    Write-Host ""
    throw "Timeout waiting for: $Description"
}

# Main conversion function
function Convert-OSTToPST {
    param([string]$SourceOST, [string]$TargetPST)
    
    Write-Host "`n=== OST to PST Converter ===" -ForegroundColor Green
    Write-Host "Source: $SourceOST" -ForegroundColor Cyan
    Write-Host "Target: $TargetPST" -ForegroundColor Cyan
    
    # Stop any existing Outlook processes
    Stop-OutlookProcesses
    
    # Validate source file
    if (-not (Test-Path $SourceOST)) {
        throw "OST file not found: $SourceOST"
    }
    
    # Create target directory if needed
    $targetDir = Split-Path $TargetPST -Parent
    if (-not (Test-Path $targetDir)) {
        New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
    }
    
    # Remove existing PST if it exists
    if (Test-Path $TargetPST) {
        Remove-Item $TargetPST -Force
        Write-Host "Removed existing PST file" -ForegroundColor Yellow
    }
    
    Write-Host "`nInitializing Outlook..." -ForegroundColor Yellow
    
    # Initialize Outlook COM objects
    $outlook = $null
    $namespace = $null
    $ostStore = $null
    $pstStore = $null
    
    try {
        # Create Outlook application
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        Write-Host "Loading OST file..." -ForegroundColor Yellow
        
        # Try to add OST store - this may fail for orphaned OST files
        $ostLoadSuccess = $false
        try {
            $namespace.AddStore($SourceOST)
            
            # Wait for OST store to be available
            $ostStore = Wait-ForOperation -Description "OST store to load" -TimeoutMinutes 2 -Operation {
                foreach ($store in $namespace.Stores) {
                    if ($store.FilePath -eq $SourceOST) {
                        try {
                            $rootFolder = $store.GetRootFolder()
                            if ($rootFolder) { return $store }
                        } catch { }
                    }
                }
                return $null
            }
            $ostLoadSuccess = $true
        } catch {
            Write-Host "  Standard OST loading failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "  This appears to be an orphaned OST file from a different profile/computer" -ForegroundColor Yellow
        }
        
        # If standard loading failed, try orphaned OST conversion approach
        if (-not $ostLoadSuccess) {
            Write-Host "  Attempting orphaned OST conversion..." -ForegroundColor Cyan
            
            # Method 1: Try to import OST directly using temporary profile
            try {
                Write-Host "  Creating temporary Outlook session..." -ForegroundColor Cyan
                
                # Create a new temporary PST first
                $tempPSTPath = [System.IO.Path]::ChangeExtension($TargetPST, ".temp.pst")
                if (Test-Path $tempPSTPath) { Remove-Item $tempPSTPath -Force }
                
                # Create PST first
                $namespace.AddStoreEx($tempPSTPath, 3) # Unicode PST
                Start-Sleep -Seconds 2
                
                # Try to force-add the OST by copying it to a temp location and renaming
                $tempOSTPath = [System.IO.Path]::ChangeExtension($SourceOST, ".temp.ost")
                Copy-Item $SourceOST $tempOSTPath -Force
                
                try {
                    # Try adding the copied OST
                    $namespace.AddStore($tempOSTPath)
                    Start-Sleep -Seconds 3
                    
                    # Look for the OST store
                    foreach ($store in $namespace.Stores) {
                        if ($store.FilePath -eq $tempOSTPath) {
                            try {
                                $ostStore = $store
                                $ostLoadSuccess = $true
                                Write-Host "  Orphaned OST loaded successfully!" -ForegroundColor Green
                                break
                            } catch {
                                Write-Host "    OST found but not accessible: $($_.Exception.Message)" -ForegroundColor Red
                            }
                        }
                    }
                } finally {
                    # Clean up temp OST
                    if (Test-Path $tempOSTPath) {
                        try { Remove-Item $tempOSTPath -Force } catch { }
                    }
                }
            } catch {
                Write-Host "    Temporary profile method failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        # Final fallback: Use scanpst.exe approach if available
        if (-not $ostLoadSuccess) {
            Write-Host "  Trying OST repair and conversion..." -ForegroundColor Cyan
            
            # Look for scanpst.exe
            $scanpstPaths = @(
                "${env:ProgramFiles}\Microsoft Office\root\Office16\scanpst.exe",
                "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\scanpst.exe",
                "${env:ProgramFiles}\Microsoft Office\Office16\scanpst.exe",
                "${env:ProgramFiles(x86)}\Microsoft Office\Office16\scanpst.exe",
                "${env:ProgramFiles}\Microsoft Office\Office15\scanpst.exe",
                "${env:ProgramFiles(x86)}\Microsoft Office\Office15\scanpst.exe"
            )
            
            $scanpstPath = $null
            foreach ($path in $scanpstPaths) {
                if (Test-Path $path) {
                    $scanpstPath = $path
                    break
                }
            }
            
            if ($scanpstPath) {
                Write-Host "  Found scanpst.exe, attempting repair..." -ForegroundColor Cyan
                
                # Create a copy of the OST for repair
                $repairOSTPath = [System.IO.Path]::ChangeExtension($SourceOST, ".repair.ost")
                Copy-Item $SourceOST $repairOSTPath -Force
                
                try {
                    # Run scanpst.exe to repair the OST
                    $process = Start-Process -FilePath $scanpstPath -ArgumentList "`"$repairOSTPath`" -f" -Wait -PassThru -WindowStyle Hidden
                    
                    if ($process.ExitCode -eq 0) {
                        Write-Host "    OST repair completed, retrying load..." -ForegroundColor Green
                        
                        # Try to load the repaired OST
                        try {
                            $namespace.AddStore($repairOSTPath)
                            Start-Sleep -Seconds 3
                            
                            foreach ($store in $namespace.Stores) {
                                if ($store.FilePath -eq $repairOSTPath) {
                                    try {
                                        $rootFolder = $store.GetRootFolder()
                                        if ($rootFolder) {
                                            $ostStore = $store
                                            $ostLoadSuccess = $true
                                            Write-Host "    Repaired OST loaded successfully!" -ForegroundColor Green
                                            break
                                        }
                                    } catch { }
                                }
                            }
                        } catch {
                            Write-Host "    Repaired OST still not accessible: $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                } finally {
                    # Clean up repair file
                    if (Test-Path $repairOSTPath) {
                        try { Remove-Item $repairOSTPath -Force } catch { }
                    }
                }
            }
        }
        
        if (-not $ostLoadSuccess -or -not $ostStore) {
            throw "Unable to load OST file. This OST may be corrupted, encrypted, or from an incompatible Outlook version. Please try using a specialized OST recovery tool or contact the original user to export the data properly."
        }
        
        Write-Host "`nOST loaded successfully!" -ForegroundColor Green
        Write-Host "Creating PST file..." -ForegroundColor Yellow
        
        # Create PST file using proper Outlook methods
        $pstCreated = $false
        
        try {
            # Method 1: Use AddStoreEx with Unicode PST format
            Write-Host "  Creating Unicode PST file..." -ForegroundColor Cyan
            $namespace.AddStoreEx($TargetPST, 3) # 3 = Unicode PST
            Start-Sleep -Seconds 3
            
            # Verify PST was created and is accessible
            $pstStore = $null
            foreach ($store in $namespace.Stores) {
                if ($store.FilePath -eq $TargetPST) {
                    try {
                        $testRoot = $store.GetRootFolder()
                        if ($testRoot) {
                            $pstStore = $store
                            $pstCreated = $true
                            Write-Host "  PST created and verified!" -ForegroundColor Green
                            break
                        }
                    } catch {
                        Write-Host "    PST created but not accessible: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
        } catch {
            Write-Host "    Unicode PST creation failed: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Method 2: Fallback to ANSI PST if Unicode failed
        if (-not $pstCreated) {
            try {
                Write-Host "  Trying ANSI PST format..." -ForegroundColor Cyan
                # Remove failed file
                if (Test-Path $TargetPST) { Remove-Item $TargetPST -Force }
                
                $namespace.AddStoreEx($TargetPST, [Microsoft.Office.Interop.Outlook.OlStoreType]::olStoreANSI)
                Start-Sleep -Seconds 3
                
                foreach ($store in $namespace.Stores) {
                    if ($store.FilePath -eq $TargetPST) {
                        try {
                            $testRoot = $store.GetRootFolder()
                            if ($testRoot) {
                                $pstStore = $store
                                $pstCreated = $true
                                Write-Host "  ANSI PST created and verified!" -ForegroundColor Green
                                break
                            }
                        } catch {
                            Write-Host "    ANSI PST created but not accessible: $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                }
            } catch {
                Write-Host "    ANSI PST creation failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        # Method 3: Final fallback using default store creation
        if (-not $pstCreated) {
            try {
                Write-Host "  Trying default PST creation..." -ForegroundColor Cyan
                # Remove failed file
                if (Test-Path $TargetPST) { Remove-Item $TargetPST -Force }
                
                # Create empty PST and add it
                $namespace.AddStore($TargetPST)
                Start-Sleep -Seconds 5
                
                foreach ($store in $namespace.Stores) {
                    if ($store.FilePath -eq $TargetPST) {
                        try {
                            $testRoot = $store.GetRootFolder()
                            if ($testRoot) {
                                $pstStore = $store
                                $pstCreated = $true
                                Write-Host "  Default PST created and verified!" -ForegroundColor Green
                                break
                            }
                        } catch {
                            Write-Host "    Default PST not accessible: $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                }
            } catch {
                Write-Host "    Default PST creation failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        if (-not $pstCreated -or -not $pstStore) {
            throw "Failed to create accessible PST file. Please ensure Outlook is properly installed and you have write permissions to the target directory."
        }
        
        Write-Host "PST created successfully!" -ForegroundColor Green
        
        # Get root folders
        $ostRoot = $ostStore.GetRootFolder()
        $pstRoot = $pstStore.GetRootFolder()
        
        Write-Host "`nStarting data copy..." -ForegroundColor Yellow
        
        # Simple recursive copy function
        function Copy-FolderRecursive {
            param($sourceFolder, $destFolder, [ref]$itemCount)
            
            Write-Host "Processing: $($sourceFolder.Name)" -ForegroundColor Cyan
            
            # Copy all items in current folder
            $items = $sourceFolder.Items
            for ($i = 1; $i -le $items.Count; $i++) {
                try {
                    $item = $items.Item($i)
                    $item.Copy().Move($destFolder)
                    $itemCount.Value++
                    
                    if ($itemCount.Value % 50 -eq 0) {
                        Write-Host "  Copied $($itemCount.Value) items..." -ForegroundColor Gray
                    }
                } catch {
                    Write-Host "    Failed to copy item: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            
            # Process subfolders
            foreach ($subfolder in $sourceFolder.Folders) {
                try {
                    $newFolder = $destFolder.Folders.Add($subfolder.Name)
                    Copy-FolderRecursive -sourceFolder $subfolder -destFolder $newFolder -itemCount $itemCount
                } catch {
                    Write-Host "    Failed to process folder '$($subfolder.Name)': $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
        
        # Start copying
        $copiedItems = [ref]0
        Copy-FolderRecursive -sourceFolder $ostRoot -destFolder $pstRoot -itemCount $copiedItems
        
        Write-Host "`n=== Conversion Complete ===" -ForegroundColor Green
        Write-Host "Items copied: $($copiedItems.Value)" -ForegroundColor Cyan
        Write-Host "PST file: $TargetPST" -ForegroundColor Cyan
        
    } catch {
        Write-Host "`nConversion failed: $($_.Exception.Message)" -ForegroundColor Red
        throw
    } finally {
        # Cleanup
        Write-Host "`nCleaning up..." -ForegroundColor Yellow
        
        try {
            if ($ostStore) { $namespace.RemoveStore($ostStore.GetRootFolder()) }
        } catch { }
        
        if ($namespace) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        }
        if ($outlook) {
            $outlook.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Final cleanup of any remaining processes
        Stop-OutlookProcesses
    }
}

# Main execution
try {
    # Get file paths if not provided
    if (-not $OSTPath) {
        do {
            $OSTPath = Read-Host "Enter OST file path"
            $OSTPath = $OSTPath.Trim('"').Trim("'")
            if (-not (Test-Path $OSTPath)) {
                Write-Host "File not found!" -ForegroundColor Red
                $OSTPath = $null
            }
        } while (-not $OSTPath)
    }
    
    if (-not $PSTPath) {
        $PSTPath = [System.IO.Path]::ChangeExtension($OSTPath, ".pst")
        $response = Read-Host "Output PST path (default: $PSTPath)"
        if ($response) { $PSTPath = $response.Trim('"').Trim("'") }
    }
    
    # Check for overwrite
    if (Test-Path $PSTPath) {
        $overwrite = Read-Host "PST exists. Overwrite? (y/n)"
        if ($overwrite -ne 'y') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            exit
        }
    }
    
    # Run conversion
    Convert-OSTToPST -SourceOST $OSTPath -TargetPST $PSTPath
    
    Write-Host "`nConversion successful!" -ForegroundColor Green
    
} catch {
    Write-Host "`nError: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}