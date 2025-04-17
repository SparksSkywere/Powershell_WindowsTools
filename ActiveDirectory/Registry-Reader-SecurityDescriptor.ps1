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

# Function to get event log registry paths and security descriptors
function Get-EventLogSecurityDescriptor {
    param (
        [string]$LogName
    )
    
    try {
        # Event logs are stored in the registry under this key
        $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\$LogName"
        
        # Get the CustomSD value if it exists
        if (Test-Path -Path $registryPath) {
            $sd = (Get-ItemProperty -Path $registryPath -Name "CustomSD" -ErrorAction SilentlyContinue).CustomSD
            
            # If CustomSD doesn't exist, fall back to default which might be inherited
            if (-not $sd) {
                $sd = "Default security descriptor (inherited)"
            }
            
            return @{
                RegistryPath = $registryPath
                SecurityDescriptor = $sd
            }
        } 
        else {
            return @{
                RegistryPath = "Path not found"
                SecurityDescriptor = "N/A"
            }
        }
    }
    catch {
        return @{
            RegistryPath = "Error"
            SecurityDescriptor = $_.Exception.Message
        }
    }
}

# Function to export information to a file
function Export-EventLogInfo {
    param (
        [string]$LogName,
        [string]$RegistryPath,
        [string]$SecurityDescriptor,
        [string]$OutputPath
    )
    
    # Create a structured format that can be imported by GPO-Creator
    $exportObject = @{
        LogName = $LogName
        RegistryPath = $RegistryPath
        SecurityDescriptor = $SecurityDescriptor
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Type = "EventLogSecurity"
    }
    
    # Export as JSON for easy import
    $exportObject | ConvertTo-Json | Out-File -FilePath $OutputPath
    
    # Also create human-readable text version
    $textOutputPath = [System.IO.Path]::ChangeExtension($OutputPath, "txt")
    $content = @"
Event Log Information: $LogName
Generated on: $(Get-Date)

Registry Path: $RegistryPath
Security Descriptor: $SecurityDescriptor

This information can be used for creating GPOs to manage event log permissions.
"@
    $content | Out-File -FilePath $textOutputPath
    
    return $OutputPath
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Event Log Security Descriptor Reader"
$form.Size = New-Object System.Drawing.Size(800, 500)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Create labels
$labelLog = New-Object System.Windows.Forms.Label
$labelLog.Text = "Select Event Log:"
$labelLog.Location = New-Object System.Drawing.Point(20, 20)
$labelLog.AutoSize = $true
$form.Controls.Add($labelLog)

$labelPath = New-Object System.Windows.Forms.Label
$labelPath.Text = "Registry Path:"
$labelPath.Location = New-Object System.Drawing.Point(20, 80)
$labelPath.AutoSize = $true
$form.Controls.Add($labelPath)

$labelSD = New-Object System.Windows.Forms.Label
$labelSD.Text = "Security Descriptor:"
$labelSD.Location = New-Object System.Drawing.Point(20, 160)
$labelSD.AutoSize = $true
$form.Controls.Add($labelSD)

# Create dropdown for event logs
$comboLogs = New-Object System.Windows.Forms.ComboBox
$comboLogs.Location = New-Object System.Drawing.Point(20, 40)
$comboLogs.Size = New-Object System.Drawing.Size(350, 20)
$comboLogs.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$form.Controls.Add($comboLogs)

# Add button to add custom event log
$buttonAddCustomLog = New-Object System.Windows.Forms.Button
$buttonAddCustomLog.Text = "Add Custom Log"
$buttonAddCustomLog.Location = New-Object System.Drawing.Point(230, 70)
$buttonAddCustomLog.Size = New-Object System.Drawing.Size(140, 23)
$form.Controls.Add($buttonAddCustomLog)

# Add available event logs to the dropdown
$eventLogs = Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog" -ErrorAction SilentlyContinue | ForEach-Object { $_.PSChildName }
if ($eventLogs) {
    foreach ($log in $eventLogs) {
        $comboLogs.Items.Add($log)
    }
    if ($comboLogs.Items.Count -gt 0) {
        $comboLogs.SelectedIndex = 0
    }
} else {
    $comboLogs.Items.Add("<No Event Logs Found>")
    $comboLogs.SelectedIndex = 0
}

# Create textboxes for displaying information
$textBoxPath = New-Object System.Windows.Forms.TextBox
$textBoxPath.Location = New-Object System.Drawing.Point(20, 100)
$textBoxPath.Size = New-Object System.Drawing.Size(740, 40)
$textBoxPath.Multiline = $true
$textBoxPath.ReadOnly = $true
$form.Controls.Add($textBoxPath)

$textBoxSD = New-Object System.Windows.Forms.TextBox
$textBoxSD.Location = New-Object System.Drawing.Point(20, 180)
$textBoxSD.Size = New-Object System.Drawing.Size(740, 220)
$textBoxSD.Multiline = $true
$textBoxSD.ScrollBars = "Vertical"
$textBoxSD.ReadOnly = $true
$form.Controls.Add($textBoxSD)

# Create buttons
$buttonGet = New-Object System.Windows.Forms.Button
$buttonGet.Location = New-Object System.Drawing.Point(400, 40)
$buttonGet.Size = New-Object System.Drawing.Size(150, 23)
$buttonGet.Text = "Get Security Descriptor"
$form.Controls.Add($buttonGet)

$buttonExport = New-Object System.Windows.Forms.Button
$buttonExport.Location = New-Object System.Drawing.Point(580, 40)
$buttonExport.Size = New-Object System.Drawing.Size(150, 23)
$buttonExport.Text = "Export Info"
$buttonExport.Enabled = $false
$form.Controls.Add($buttonExport)

# Add event handler for the Get button
$buttonGet.Add_Click({
    $selectedLog = $comboLogs.SelectedItem
    
    if ($selectedLog) {
        $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\$selectedLog"
        
        if (Test-Path -Path $registryPath) {
            $result = Get-EventLogSecurityDescriptor -LogName $selectedLog
            $textBoxPath.Text = $result.RegistryPath
            $textBoxSD.Text = $result.SecurityDescriptor
        } else {
            $textBoxPath.Text = "Custom Event Log (Not present on this system)"
            $textBoxSD.Text = "O:BAG:SYD:(A;;0xf0007;;;SY)(A;;0x7;;;BA)(A;;0x7;;;SO)(A;;0x3;;;IU)(A;;0x3;;;SU)(A;;0x3;;;S-1-5-3)(A;;0x3;;;S-1-5-33)(A;;0x1;;;S-1-5-32-573)"
        }
        
        $buttonExport.Enabled = $true
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please select an Event Log.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# Add event handler for the Export button
$buttonExport.Add_Click({
    $selectedLog = $comboLogs.SelectedItem
    
    if ($selectedLog) {
        $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\$selectedLog"
        
        if (Test-Path -Path $registryPath) {
            $result = Get-EventLogSecurityDescriptor -LogName $selectedLog
            $securityDescriptor = $result.SecurityDescriptor
            $regPath = $result.RegistryPath
        } else {
            $securityDescriptor = $textBoxSD.Text
            $regPath = "SYSTEM\CurrentControlSet\Services\EventLog\$selectedLog (Custom Event Log)"
        }
        
        # Create a SaveFileDialog to choose the export location
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "GPO Creator Files (*.gpodata)|*.gpodata|All files (*.*)|*.*"
        $saveDialog.Title = "Save Event Log Information"
        $saveDialog.FileName = "EventLog_$selectedLog`_$(Get-Date -Format 'yyyyMMdd_HHmmss').gpodata"
        $saveDialog.DefaultExt = "gpodata"
        
        # Show the dialog and process the result
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportPath = Export-EventLogInfo -LogName $selectedLog -RegistryPath $regPath -SecurityDescriptor $securityDescriptor -OutputPath $saveDialog.FileName
            [System.Windows.Forms.MessageBox]::Show("Information exported to:`n$exportPath`n`nA text version has also been created.", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
})

# Add event handler for the Custom Log button
$buttonAddCustomLog.Add_Click({
    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = "Add Custom Event Log"
    $inputForm.Size = New-Object System.Drawing.Size(400, 150)
    $inputForm.StartPosition = "CenterParent"
    $inputForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $inputForm.MaximizeBox = $false
    $inputForm.MinimizeBox = $false
    
    $inputLabel = New-Object System.Windows.Forms.Label
    $inputLabel.Text = "Enter Custom Event Log Name:"
    $inputLabel.Location = New-Object System.Drawing.Point(10, 20)
    $inputLabel.Size = New-Object System.Drawing.Size(370, 20)
    $inputForm.Controls.Add($inputLabel)
    
    $inputTextBox = New-Object System.Windows.Forms.TextBox
    $inputTextBox.Location = New-Object System.Drawing.Point(10, 45)
    $inputTextBox.Size = New-Object System.Drawing.Size(370, 20)
    $inputForm.Controls.Add($inputTextBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(110, 80)
    $okButton.Size = New-Object System.Drawing.Size(80, 23)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $inputForm.Controls.Add($okButton)
    $inputForm.AcceptButton = $okButton
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(200, 80)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 23)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $inputForm.Controls.Add($cancelButton)
    $inputForm.CancelButton = $cancelButton
    
    # Show the form as a dialog
    $result = $inputForm.ShowDialog()
    
    # Process the input
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $customLogName = $inputTextBox.Text.Trim()
        if ($customLogName) {
            # Check if the log name already exists in the combo box
            $exists = $false
            foreach ($item in $comboLogs.Items) {
                if ($item -eq $customLogName) {
                    $exists = $true
                    break
                }
            }
            
            if (-not $exists) {
                # Remove placeholder if it exists
                if ($comboLogs.Items.Count -eq 1 -and $comboLogs.Items[0] -eq "<No Event Logs Found>") {
                    $comboLogs.Items.Clear()
                }
                
                # Add the custom log name and select it
                $comboLogs.Items.Add($customLogName)
            }
            
            # Select the item
            $comboLogs.SelectedItem = $customLogName
            
            # Set the text boxes
            $textBoxPath.Text = "Custom Event Log (Not present on this system)"
            $textBoxSD.Text = "Default security descriptor template for a custom Event Log. Modify as needed."
            $buttonExport.Enabled = $true
        }
    }
})

# Show the form
$form.ShowDialog()
