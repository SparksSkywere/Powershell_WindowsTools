#Requires -RunAsAdministrator

# Add Windows Forms assembly
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
#end of powershell console hiding
#To show the console change "-hide" to "-show"
show-console -hide

# Function to get registry values from a path
function Get-RegistryValues {
    param (
        [string]$RegistryPath,
        [bool]$IncludeSubkeys = $false
    )
    
    try {
        # Check if the path exists
        if (-not (Test-Path -Path $RegistryPath)) {
            return @{
                Success = $false
                Message = "Registry path does not exist."
                Values = @()
            }
        }

        $results = @()

        # Get direct values from the specified key
        $values = Get-ItemProperty -Path $RegistryPath -ErrorAction SilentlyContinue
        if ($values) {
            foreach ($property in ($values.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' })) {
                $type = Get-RegistryValueKind -Path $RegistryPath -Name $property.Name
                
                $results += [PSCustomObject]@{
                    KeyPath = $RegistryPath
                    ValueName = $property.Name
                    ValueData = $property.Value
                    ValueType = $type
                }
            }
        }

        # Recursively get values from subkeys if requested
        if ($IncludeSubkeys) {
            $subkeys = Get-ChildItem -Path $RegistryPath -ErrorAction SilentlyContinue
            foreach ($subkey in $subkeys) {
                $subResults = Get-RegistryValues -RegistryPath $subkey.PSPath -IncludeSubkeys $true
                if ($subResults.Success) {
                    $results += $subResults.Values
                }
            }
        }

        return @{
            Success = $true
            Message = "Retrieved $($results.Count) registry values."
            Values = $results
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Error: $($_.Exception.Message)"
            Values = @()
        }
    }
}

# Function to determine registry value type
function Get-RegistryValueKind {
    param (
        [string]$Path,
        [string]$Name
    )
    
    try {
        # Get the key
        $key = Get-Item -LiteralPath $Path
        
        # Get the value kind for the specified name
        $valueKind = $key.GetValueKind($Name)
        
        # Convert from .NET registry value kind to string type
        switch ($valueKind) {
            "String" { return "String" }
            "ExpandString" { return "ExpandString" }
            "Binary" { return "Binary" }
            "DWord" { return "DWord" }
            "QWord" { return "QWord" }
            "MultiString" { return "MultiString" }
            default { return "Unknown" }
        }
    }
    catch {
        return "Unknown"
    }
}

# Function to convert registry path from PowerShell format to standard format
function Convert-RegistryPath {
    param (
        [string]$Path
    )
    
    # Replace PowerShell drive names with standard registry path roots
    $path = $Path -replace "HKLM:", "HKLM\" -replace "HKCU:", "HKCU\" `
                 -replace "HKCR:", "HKCR\" -replace "HKU:", "HKU\" `
                 -replace "HKCC:", "HKCC\" -replace "Microsoft.PowerShell.Core\\Registry::", ""
    
    return $path
}

# Function to export registry settings to a file
function Export-RegistrySettings {
    param (
        [Array]$RegistrySettings,
        [string]$OutputPath
    )
    
    # Create a structured format that can be imported by GPO-Creator
    $exportObject = @{
        Type = "RegistrySettings"
        Settings = $RegistrySettings | ForEach-Object {
            @{
                KeyPath = (Convert-RegistryPath -Path $_.KeyPath)
                ValueName = $_.ValueName
                ValueData = $_.ValueData
                ValueType = $_.ValueType
            }
        }
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    # Export as JSON for easy import
    $exportObject | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputPath
    
    # Also create human-readable text version
    $textOutputPath = [System.IO.Path]::ChangeExtension($OutputPath, "txt")
    $content = @"
Registry Settings Export
Generated on: $(Get-Date)

$($RegistrySettings.Count) settings exported:

$(($RegistrySettings | ForEach-Object { 
    "Key: $($_.KeyPath)`nName: $($_.ValueName)`nType: $($_.ValueType)`nValue: $($_.ValueData)`n" 
}) -join "`n")

This information can be used for creating GPOs.
"@
    $content | Out-File -FilePath $textOutputPath
    
    return $OutputPath
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Registry Settings Extractor"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Create Group Box for Registry Path
$groupBoxPath = New-Object System.Windows.Forms.GroupBox
$groupBoxPath.Text = "Registry Path"
$groupBoxPath.Location = New-Object System.Drawing.Point(20, 20)
$groupBoxPath.Size = New-Object System.Drawing.Size(850, 100)
$form.Controls.Add($groupBoxPath)

# Create controls for Registry Path
$labelRoot = New-Object System.Windows.Forms.Label
$labelRoot.Text = "Registry Root:"
$labelRoot.Location = New-Object System.Drawing.Point(10, 30)
$labelRoot.AutoSize = $true
$groupBoxPath.Controls.Add($labelRoot)

$comboRoot = New-Object System.Windows.Forms.ComboBox
$comboRoot.Items.Add("HKLM:")
$comboRoot.Items.Add("HKCU:")
$comboRoot.Items.Add("HKCR:")
$comboRoot.Items.Add("HKU:")
$comboRoot.Items.Add("HKCC:")
$comboRoot.SelectedIndex = 0
$comboRoot.Location = New-Object System.Drawing.Point(120, 27)
$comboRoot.Size = New-Object System.Drawing.Size(100, 20)
$comboRoot.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$groupBoxPath.Controls.Add($comboRoot)

$labelPath = New-Object System.Windows.Forms.Label
$labelPath.Text = "Path:"
$labelPath.Location = New-Object System.Drawing.Point(230, 30)
$labelPath.AutoSize = $true
$groupBoxPath.Controls.Add($labelPath)

$textPath = New-Object System.Windows.Forms.TextBox
$textPath.Location = New-Object System.Drawing.Point(270, 27)
$textPath.Size = New-Object System.Drawing.Size(450, 20)
$textPath.Text = "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies"
$groupBoxPath.Controls.Add($textPath)

$checkBoxRecursive = New-Object System.Windows.Forms.CheckBox
$checkBoxRecursive.Text = "Include Subkeys"
$checkBoxRecursive.Location = New-Object System.Drawing.Point(120, 60)
$checkBoxRecursive.AutoSize = $true
$groupBoxPath.Controls.Add($checkBoxRecursive)

$buttonGetValues = New-Object System.Windows.Forms.Button
$buttonGetValues.Text = "Get Registry Values"
$buttonGetValues.Location = New-Object System.Drawing.Point(730, 27)
$buttonGetValues.Size = New-Object System.Drawing.Size(110, 50)
$groupBoxPath.Controls.Add($buttonGetValues)

# Create ListView for registry values
$listViewRegistry = New-Object System.Windows.Forms.ListView
$listViewRegistry.Location = New-Object System.Drawing.Point(20, 130)
$listViewRegistry.Size = New-Object System.Drawing.Size(850, 400)
$listViewRegistry.View = [System.Windows.Forms.View]::Details
$listViewRegistry.FullRowSelect = $true
$listViewRegistry.MultiSelect = $true
$listViewRegistry.Columns.Add("Registry Key", 300)
$listViewRegistry.Columns.Add("Value Name", 150)
$listViewRegistry.Columns.Add("Type", 80)
$listViewRegistry.Columns.Add("Value Data", 300)
$form.Controls.Add($listViewRegistry)

# Create Group Box for Actions
$groupBoxActions = New-Object System.Windows.Forms.GroupBox
$groupBoxActions.Text = "Actions"
$groupBoxActions.Location = New-Object System.Drawing.Point(20, 540)
$groupBoxActions.Size = New-Object System.Drawing.Size(850, 110)
$form.Controls.Add($groupBoxActions)

# Create controls for Actions
$buttonExport = New-Object System.Windows.Forms.Button
$buttonExport.Text = "Export Selected"
$buttonExport.Location = New-Object System.Drawing.Point(10, 30)
$buttonExport.Size = New-Object System.Drawing.Size(120, 30)
$buttonExport.Enabled = $false
$groupBoxActions.Controls.Add($buttonExport)

$buttonExportAll = New-Object System.Windows.Forms.Button
$buttonExportAll.Text = "Export All"
$buttonExportAll.Location = New-Object System.Drawing.Point(140, 30)
$buttonExportAll.Size = New-Object System.Drawing.Size(120, 30)
$buttonExportAll.Enabled = $false
$groupBoxActions.Controls.Add($buttonExportAll)

$buttonClearSelection = New-Object System.Windows.Forms.Button
$buttonClearSelection.Text = "Clear Selection"
$buttonClearSelection.Location = New-Object System.Drawing.Point(270, 30)
$buttonClearSelection.Size = New-Object System.Drawing.Size(120, 30)
$groupBoxActions.Controls.Add($buttonClearSelection)

$buttonCommonSettings = New-Object System.Windows.Forms.Button
$buttonCommonSettings.Text = "Common Settings"
$buttonCommonSettings.Location = New-Object System.Drawing.Point(400, 30)
$buttonCommonSettings.Size = New-Object System.Drawing.Size(140, 30)
$groupBoxActions.Controls.Add($buttonCommonSettings)

$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = "Ready to extract registry settings..."
$labelStatus.Location = New-Object System.Drawing.Point(10, 70)
$labelStatus.Size = New-Object System.Drawing.Size(830, 30)
$groupBoxActions.Controls.Add($labelStatus)

# Add event handler for Get Values button
$buttonGetValues.Add_Click({
    $listViewRegistry.Items.Clear()
    
    # Build registry path
    $root = $comboRoot.SelectedItem.ToString()
    $path = $textPath.Text.Trim()
    $registryPath = Join-Path -Path $root -ChildPath $path
    
    # Get registry values
    $result = Get-RegistryValues -RegistryPath $registryPath -IncludeSubkeys $checkBoxRecursive.Checked
    
    if ($result.Success) {
        # Populate ListView with values
        foreach ($value in $result.Values) {
            $item = New-Object System.Windows.Forms.ListViewItem($value.KeyPath)
            $item.SubItems.Add($value.ValueName)
            $item.SubItems.Add($value.ValueType)
            
            # Format the value data based on type for better display
            $displayValue = $value.ValueData
            if ($value.ValueType -eq "Binary" -and $value.ValueData -is [byte[]]) {
                $displayValue = ($value.ValueData | ForEach-Object { $_.ToString("X2") }) -join " "
            }
            
            $item.SubItems.Add($displayValue)
            $item.Tag = $value  # Store the original value object
            $listViewRegistry.Items.Add($item)
        }
        
        $labelStatus.Text = $result.Message
        
        # Enable export buttons if values were found
        if ($result.Values.Count -gt 0) {
            $buttonExportAll.Enabled = $true
        } else {
            $buttonExportAll.Enabled = $false
        }
    } else {
        $labelStatus.Text = $result.Message
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Add event handler for ListView selection change
$listViewRegistry.Add_SelectedIndexChanged({
    if ($listViewRegistry.SelectedItems.Count -gt 0) {
        $buttonExport.Enabled = $true
    } else {
        $buttonExport.Enabled = $false
    }
})

# Add event handler for Clear Selection button
$buttonClearSelection.Add_Click({
    $listViewRegistry.SelectedItems.Clear()
    $buttonExport.Enabled = $false
})

# Add event handler for Export Selected button
$buttonExport.Add_Click({
    if ($listViewRegistry.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No items selected.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Collect selected registry settings
    $selectedSettings = @()
    foreach ($item in $listViewRegistry.SelectedItems) {
        $selectedSettings += $item.Tag
    }
    
    # Create a SaveFileDialog to choose the export location
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "GPO Creator Files (*.gpodata)|*.gpodata|All files (*.*)|*.*"
    $saveDialog.Title = "Save Registry Settings"
    $saveDialog.FileName = "RegistrySettings_$(Get-Date -Format 'yyyyMMdd_HHmmss').gpodata"
    $saveDialog.DefaultExt = "gpodata"
    
    # Show the dialog and process the result
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $exportPath = Export-RegistrySettings -RegistrySettings $selectedSettings -OutputPath $saveDialog.FileName
        [System.Windows.Forms.MessageBox]::Show("Registry settings exported to:`n$exportPath`n`nA text version has also been created.", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Add event handler for Export All button
$buttonExportAll.Add_Click({
    if ($listViewRegistry.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No registry settings to export.", "Empty List", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Collect all registry settings
    $allSettings = @()
    foreach ($item in $listViewRegistry.Items) {
        $allSettings += $item.Tag
    }
    
    # Create a SaveFileDialog to choose the export location
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "GPO Creator Files (*.gpodata)|*.gpodata|All files (*.*)|*.*"
    $saveDialog.Title = "Save Registry Settings"
    $saveDialog.FileName = "RegistrySettings_$(Get-Date -Format 'yyyyMMdd_HHmmss').gpodata"
    $saveDialog.DefaultExt = "gpodata"
    
    # Show the dialog and process the result
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $exportPath = Export-RegistrySettings -RegistrySettings $allSettings -OutputPath $saveDialog.FileName
        [System.Windows.Forms.MessageBox]::Show("Registry settings exported to:`n$exportPath`n`nA text version has also been created.", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Add event handler for Common Settings button
$buttonCommonSettings.Add_Click({
    # Create a form to display common registry paths
    $commonForm = New-Object System.Windows.Forms.Form
    $commonForm.Text = "Common Registry Settings"
    $commonForm.Size = New-Object System.Drawing.Size(600, 400)
    $commonForm.StartPosition = "CenterParent"
    $commonForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $commonForm.MaximizeBox = $false
    $commonForm.MinimizeBox = $false
    
    # Create a ListView for common registry paths
    $listViewCommon = New-Object System.Windows.Forms.ListView
    $listViewCommon.Location = New-Object System.Drawing.Point(10, 10)
    $listViewCommon.Size = New-Object System.Drawing.Size(565, 300)
    $listViewCommon.View = [System.Windows.Forms.View]::Details
    $listViewCommon.FullRowSelect = $true
    $listViewCommon.Columns.Add("Description", 300)
    $listViewCommon.Columns.Add("Registry Path", 250)
    $commonForm.Controls.Add($listViewCommon)
    
    # Add common registry paths
    $commonPaths = @(
        @{Description = "Windows Update Settings"; Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"},
        @{Description = "Internet Explorer Settings"; Path = "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer"},
        @{Description = "Windows Defender Settings"; Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender"},
        @{Description = "Windows Firewall Settings"; Path = "HKLM:\SOFTWARE\Policies\Microsoft\WindowsFirewall"},
        @{Description = "Active Desktop and Screen Saver"; Path = "HKCU:\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop"},
        @{Description = "Start Menu Settings"; Path = "HKCU:\SOFTWARE\Policies\Microsoft\Windows\Explorer"},
        @{Description = "Office Common Settings"; Path = "HKCU:\SOFTWARE\Policies\Microsoft\Office\Common"},
        @{Description = "Edge Browser Settings"; Path = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"},
        @{Description = "Chrome Browser Settings"; Path = "HKLM:\SOFTWARE\Policies\Google\Chrome"},
        @{Description = "Firefox Browser Settings"; Path = "HKLM:\SOFTWARE\Policies\Mozilla\Firefox"},
        @{Description = "Remote Desktop Settings"; Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"},
        @{Description = "Windows Firewall Advanced Settings"; Path = "HKLM:\SOFTWARE\Policies\Microsoft\WindowsFirewall\FirewallRules"},
        @{Description = "PowerShell Execution Policy"; Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell"},
        @{Description = "Windows Update for Business"; Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"},
        @{Description = "AppLocker Settings"; Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\SrpV2"}
    )
    
    foreach ($path in $commonPaths) {
        $item = New-Object System.Windows.Forms.ListViewItem($path.Description)
        $item.SubItems.Add($path.Path)
        $item.Tag = $path
        $listViewCommon.Items.Add($item)
    }
    
    # Add a button to select the path
    $buttonSelect = New-Object System.Windows.Forms.Button
    $buttonSelect.Text = "Select Path"
    $buttonSelect.Location = New-Object System.Drawing.Point(240, 320)
    $buttonSelect.Size = New-Object System.Drawing.Size(100, 30)
    $commonForm.Controls.Add($buttonSelect)
    
    # Add event handler for the select button
    $buttonSelect.Add_Click({
        if ($listViewCommon.SelectedItems.Count -gt 0) {
            $selectedPath = $listViewCommon.SelectedItems[0].Tag
            $pathParts = $selectedPath.Path -split ':'
            if ($pathParts.Count -gt 1) {
                # Set the root and path in the main form
                $rootPart = $pathParts[0] + ":"
                $remainingPath = $pathParts[1]
                if ($remainingPath.StartsWith('\')) {
                    $remainingPath = $remainingPath.Substring(1)
                }
                
                # Update the main form controls
                for ($i = 0; $i -lt $comboRoot.Items.Count; $i++) {
                    if ($comboRoot.Items[$i] -eq $rootPart) {
                        $comboRoot.SelectedIndex = $i
                        break
                    }
                }
                
                $textPath.Text = $remainingPath
            }
            
            $commonForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $commonForm.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a registry path.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    
    # Show the form as a dialog
    $commonForm.ShowDialog() | Out-Null
})

# Show the form
$form.ShowDialog()