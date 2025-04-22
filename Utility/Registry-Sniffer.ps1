# Registry Sniffer - GUI tool to find registry keys containing specified text

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-Console
{
    param ([Switch]$Show,[Switch]$Hide)
    if (-not ("Console.Window" -as [type])) { 

        Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
        '
    }
    if ($Show)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        $null = [Console.Window]::ShowWindow($consolePtr, 5)
    }
    if ($Hide)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        #0 hide
        $null = [Console.Window]::ShowWindow($consolePtr, 0)
    }
}
# End of powershell console hiding
# To show the console change "-hide" to "-show"
show-console -hide

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Registry Sniffer"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Create search input label
$labelSearch = New-Object System.Windows.Forms.Label
$labelSearch.Location = New-Object System.Drawing.Point(10, 20)
$labelSearch.Size = New-Object System.Drawing.Size(120, 20)
$labelSearch.Text = "Search String:"
$form.Controls.Add($labelSearch)

# Create search input textbox
$textBoxSearch = New-Object System.Windows.Forms.TextBox
$textBoxSearch.Location = New-Object System.Drawing.Point(130, 20)
$textBoxSearch.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($textBoxSearch)

# Create search button
$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Location = New-Object System.Drawing.Point(440, 18)
$buttonSearch.Size = New-Object System.Drawing.Size(100, 25)
$buttonSearch.Text = "Search"
$form.Controls.Add($buttonSearch)

# Create status label
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(10, 50)
$labelStatus.Size = New-Object System.Drawing.Size(680, 20)
$labelStatus.Text = "Ready"
$form.Controls.Add($labelStatus)

# Create percentage label
$labelPercentage = New-Object System.Windows.Forms.Label
$labelPercentage.Location = New-Object System.Drawing.Point(700, 50)
$labelPercentage.Size = New-Object System.Drawing.Size(70, 20)
$labelPercentage.Text = "0%"
$labelPercentage.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($labelPercentage)

# Create progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 70)
$progressBar.Size = New-Object System.Drawing.Size(760, 15)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Step = 1
$form.Controls.Add($progressBar)

# Create sorting options groupbox
$groupBoxSort = New-Object System.Windows.Forms.GroupBox
$groupBoxSort.Location = New-Object System.Drawing.Point(550, 10)
$groupBoxSort.Size = New-Object System.Drawing.Size(220, 50)
$groupBoxSort.Text = "Sort By"
$form.Controls.Add($groupBoxSort)

# Create sort by path radio button
$radioPath = New-Object System.Windows.Forms.RadioButton
$radioPath.Location = New-Object System.Drawing.Point(10, 20)
$radioPath.Size = New-Object System.Drawing.Size(100, 20)
$radioPath.Text = "Path (A-Z)"
$radioPath.Checked = $true
$groupBoxSort.Controls.Add($radioPath)

# Create sort by location radio button
$radioLocation = New-Object System.Windows.Forms.RadioButton
$radioLocation.Location = New-Object System.Drawing.Point(110, 20)
$radioLocation.Size = New-Object System.Drawing.Size(100, 20)
$radioLocation.Text = "Location"
$groupBoxSort.Controls.Add($radioLocation)

# Add case-insensitive checkbox
$checkBoxCaseInsensitive = New-Object System.Windows.Forms.CheckBox
$checkBoxCaseInsensitive.Location = New-Object System.Drawing.Point(440, 50)
$checkBoxCaseInsensitive.Size = New-Object System.Drawing.Size(150, 20)
$checkBoxCaseInsensitive.Text = "Case-Insensitive"
$checkBoxCaseInsensitive.Checked = $true  # Default to case-insensitive
$form.Controls.Add($checkBoxCaseInsensitive)

# Create list view
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, 100)
$listView.Size = New-Object System.Drawing.Size(760, 440)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.MultiSelect = $false

# Add columns
$listView.Columns.Add("Registry Path", 450) | Out-Null
$listView.Columns.Add("Hive", 100) | Out-Null
$listView.Columns.Add("Type", 80) | Out-Null
$listView.Columns.Add("Data", 120) | Out-Null

$form.Controls.Add($listView)

# Function to update progress safely
function Update-Progress {
    param (
        [int]$Value,
        [string]$StatusText
    )
    
    try {
        # Set the progress bar value with bounds checking
        $progressBar.Value = [Math]::Min([Math]::Max($Value, 0), 100)
        $labelPercentage.Text = "$($progressBar.Value)%"
        
        if ($StatusText) {
            $labelStatus.Text = $StatusText
        }
        
        # Force UI refresh
        $form.Refresh()
    }
    catch {
        # Silently handle any UI update errors
        # This prevents error cascades
    }
}

# Function to search registry keys
function Search-RegistryKeys {
    param (
        [string]$SearchString,
        [bool]$CaseInsensitive = $true
    )
    
    Update-Progress -Value 0 -StatusText "Searching registry for: $SearchString..."
    
    $results = @()
    
    # Search in common registry hives
    $hives = @(
        @{Name = "HKCU"; Path = "HKCU:\" },
        @{Name = "HKLM"; Path = "HKLM:\" }
    )
    
    # Limit registry crawling depth for performance
    $maxDepth = 5
    $keyCount = 0
    $keyLimit = 10000  # Avoid excessive searching
    
    for ($hiveIndex = 0; $hiveIndex -lt $hives.Count; $hiveIndex++) {
        $hive = $hives[$hiveIndex]
        $baseProgress = ($hiveIndex / $hives.Count) * 90
        
        Update-Progress -Value $baseProgress -StatusText "Searching in $($hive.Name)..."
        
        try {
            # Use a more controlled approach to search the registry
            $stack = New-Object System.Collections.Stack
            $stack.Push(@{Path = $hive.Path; Depth = 0})
            
            while ($stack.Count -gt 0 -and $keyCount -lt $keyLimit) {
                $current = $stack.Pop()
                $currentPath = $current.Path
                $currentDepth = $current.Depth
                
                $keyCount++
                
                # Update progress every 100 keys
                if ($keyCount % 100 -eq 0) {
                    $progressPercent = $baseProgress + (($keyCount / $keyLimit) * (90 / $hives.Count))
                    Update-Progress -Value $progressPercent -StatusText "Searching... (Keys processed: $keyCount)"
                }
                
                try {
                    $key = Get-Item -Path $currentPath -ErrorAction SilentlyContinue
                    
                    # Check the current key name against the search string using correct case sensitivity
                    $keyNameMatches = if ($CaseInsensitive) {
                        $key -and $key.PSChildName -like "*$SearchString*"
                    } else {
                        $key -and $key.PSChildName -clike "*$SearchString*"
                    }
                    
                    if ($keyNameMatches) {
                        $results += [PSCustomObject]@{
                            Path = $key.PSPath.Replace("Microsoft.PowerShell.Core\Registry::", "")
                            Hive = $hive.Name
                            Type = "Key"
                            Data = ""
                        }
                    }
                    
                    # Check properties for the search string
                    $key | Get-ItemProperty -ErrorAction SilentlyContinue | ForEach-Object {
                        foreach ($prop in $_.PSObject.Properties) {
                            if ($prop.Name -notlike "PS*") {
                                $propValue = $prop.Value
                                
                                # Check property name and value using correct case sensitivity
                                $nameMatches = if ($CaseInsensitive) {
                                    $prop.Name -like "*$SearchString*"
                                } else {
                                    $prop.Name -clike "*$SearchString*"
                                }
                                
                                $valueMatches = if ($CaseInsensitive) {
                                    $propValue -is [string] -and $propValue -like "*$SearchString*"
                                } else {
                                    $propValue -is [string] -and $propValue -clike "*$SearchString*"
                                }
                                
                                if ($nameMatches -or $valueMatches) {
                                    $cleanPath = $key.PSPath.Replace("Microsoft.PowerShell.Core\Registry::", "")
                                    $results += [PSCustomObject]@{
                                        Path = "$cleanPath\$($prop.Name)"
                                        Hive = $hive.Name
                                        Type = "Property"
                                        Data = $propValue
                                    }
                                }
                            }
                        }
                    }
                    
                    # Only proceed deeper if we're within the maximum depth
                    if ($currentDepth -lt $maxDepth) {
                        # Get child keys and add them to the stack
                        $key | Get-ChildItem -ErrorAction SilentlyContinue | ForEach-Object {
                            $stack.Push(@{Path = $_.PSPath; Depth = $currentDepth + 1})
                        }
                    }
                }
                catch {
                    # Skip keys that can't be accessed
                    continue
                }
            }
        }
        catch {
            # Handle any unexpected errors during search
            Update-Progress -Value $baseProgress -StatusText "Error searching $($hive.Name): $($_.Exception.Message)"
            Start-Sleep -Seconds 1  # Pause to allow user to see error message
        }
    }
    
    Update-Progress -Value 100 -StatusText "Search complete."
    return $results
}

# Function to display results in the list view
function Display-Results {
    param (
        [array]$Results
    )
    
    try {
        $listView.BeginUpdate()
        $listView.Items.Clear()
        
        # Check if any results were found
        if ($null -eq $Results -or $Results.Count -eq 0) {
            $labelStatus.Text = "No results found for '$($textBoxSearch.Text)'"
            $listView.EndUpdate()
            return
        }
        
        # Sort results based on selected option
        if ($radioPath.Checked) {
            $sortedResults = $Results | Sort-Object -Property Path
        }
        else {
            $sortedResults = $Results | Sort-Object -Property Hive, Path
        }
        
        foreach ($result in $sortedResults) {
            try {
                $item = New-Object System.Windows.Forms.ListViewItem($result.Path)
                $item.SubItems.Add($result.Hive)
                $item.SubItems.Add($result.Type)
                
                # Safely handle data of any type
                $data = ""
                if ($null -ne $result.Data) {
                    # Handle different data types appropriately
                    if ($result.Data -is [string]) {
                        $data = $result.Data
                        # Truncate long strings
                        if ($data.Length -gt 50) {
                            $data = $data.Substring(0, 50) + "..."
                        }
                    }
                    elseif ($result.Data -is [byte[]]) {
                        # Format byte arrays nicely
                        $data = "Binary data: {0} bytes" -f $result.Data.Length
                    }
                    elseif ($result.Data -is [array]) {
                        # Summarize arrays
                        $data = "Array: {0} items" -f $result.Data.Count
                    }
                    else {
                        # Convert other objects to string, safely handle large objects
                        try {
                            $dataStr = "$($result.Data)"
                            if ($dataStr.Length -gt 50) {
                                $data = $dataStr.Substring(0, 50) + "..."
                            }
                            else {
                                $data = $dataStr
                            }
                        }
                        catch {
                            $data = "Complex object: $($result.Data.GetType().Name)"
                        }
                    }
                }
                
                $item.SubItems.Add($data)
                $listView.Items.Add($item)
            }
            catch {
                # Log problematic items but continue processing others
                Write-Host "Error adding item: $($result.Path) - $($_.Exception.Message)"
            }
        }
        
        $labelStatus.Text = "Found $($sortedResults.Count) results for '$($textBoxSearch.Text)'"
        $listView.EndUpdate()
    }
    catch {
        # Final fallback for any UI errors
        $errorMsg = "Error displaying results: $($_.Exception.Message)"
        $labelStatus.Text = $errorMsg
        
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        # Ensure that EndUpdate is called even if an exception occurs
        if ($listView.IsHandleCreated) {
            $listView.EndUpdate()
        }
    }
}

# Add event for search button click
$buttonSearch.Add_Click({
    if ([string]::IsNullOrWhiteSpace($textBoxSearch.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a search string", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Disable the button during search
    $buttonSearch.Enabled = $false
    
    try {
        # Set initial progress
        Update-Progress -Value 0 -StatusText "Starting search for: $($textBoxSearch.Text)"
        
        # Perform the search directly without using jobs
        $results = Search-RegistryKeys -SearchString $textBoxSearch.Text -CaseInsensitive $checkBoxCaseInsensitive.Checked
        
        # Display the results
        Display-Results -Results $results
    }
    catch {
        # Handle any errors during search
        $errorMessage = $_.Exception.Message
        Update-Progress -Value 0 -StatusText "Error: $errorMessage"
        
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred during the search: $errorMessage", 
            "Search Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        # Always re-enable the search button
        $buttonSearch.Enabled = $true
    }
})

# Add event for radio button changes
$radioPath.Add_CheckedChanged({
    if ($listView.Items.Count -gt 0) {
        try {
            # Re-sort the list view based on path
            $listView.BeginUpdate()
            $items = New-Object System.Collections.ArrayList
            
            foreach ($item in $listView.Items) {
                $items.Add($item) | Out-Null
            }
            
            $sortedItems = $items | Sort-Object -Property @{Expression={$_.Text}; Ascending=$true}
            
            $listView.Items.Clear()
            foreach ($item in $sortedItems) {
                $listView.Items.Add($item)
            }
            $listView.EndUpdate()
        }
        catch {
            # Handle any sorting errors
            [System.Windows.Forms.MessageBox]::Show(
                "Error sorting results: $($_.Exception.Message)", 
                "Sorting Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

$radioLocation.Add_CheckedChanged({
    if ($listView.Items.Count -gt 0) {
        try {
            # Re-sort the list view based on hive/location
            $listView.BeginUpdate()
            $items = New-Object System.Collections.ArrayList
            
            foreach ($item in $listView.Items) {
                $items.Add($item) | Out-Null
            }
            
            $sortedItems = $items | Sort-Object -Property @{Expression={$_.SubItems[1].Text}; Ascending=$true}, 
                                                  @{Expression={$_.Text}; Ascending=$true}
            
            $listView.Items.Clear()
            foreach ($item in $sortedItems) {
                $listView.Items.Add($item)
            }
            $listView.EndUpdate()
        }
        catch {
            # Handle any sorting errors
            [System.Windows.Forms.MessageBox]::Show(
                "Error sorting results: $($_.Exception.Message)", 
                "Sorting Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

# Add right-click context menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$copyMenuItem = $contextMenu.Items.Add("Copy Path")
$copyMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        [System.Windows.Forms.Clipboard]::SetText($listView.SelectedItems[0].Text)
    }
})

$openRegEdit = $contextMenu.Items.Add("Open in Registry Editor")
$openRegEdit.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $path = $listView.SelectedItems[0].Text
        Start-Process "regedit.exe" -ArgumentList "/m", $path
    }
})

$listView.ContextMenuStrip = $contextMenu

# Show the form
[void]$form.ShowDialog()
