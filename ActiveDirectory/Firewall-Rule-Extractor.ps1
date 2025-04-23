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

# Function to get firewall rules
function Get-LocalFirewallRules {
    param (
        [string]$DisplayName = "*",
        [string]$Direction = "*",
        [string]$Action = "*",
        [string]$Protocol = "*"
    )
    
    try {
        # Get all firewall rules based on the filter criteria
        $rules = Get-NetFirewallRule -DisplayName $DisplayName -Direction $Direction -Action $Action -ErrorAction Stop
        
        # Process each rule to get detailed information
        $results = @()
        foreach ($rule in $rules) {
            try {
                # Get rule details
                $ports = Get-NetFirewallPortFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
                $addresses = Get-NetFirewallAddressFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
                $app = Get-NetFirewallApplicationFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
                
                # Filter by protocol if specified
                if ($Protocol -ne "*" -and $ports.Protocol -ne $Protocol) {
                    continue
                }
                
                # Create custom object with rule information
                $result = [PSCustomObject]@{
                    Name = $rule.Name
                    DisplayName = $rule.DisplayName
                    Description = $rule.Description
                    Direction = $rule.Direction
                    Action = $rule.Action
                    Enabled = $rule.Enabled
                    Protocol = $ports.Protocol
                    LocalPorts = $ports.LocalPort -join ","
                    RemotePorts = $ports.RemotePort -join ","
                    LocalAddresses = $addresses.LocalAddress -join ","
                    RemoteAddresses = $addresses.RemoteAddress -join ","
                    Program = $app.Program
                    PackageFamilyName = $app.PackageFamilyName
                }
                
                $results += $result
            }
            catch {
                # Skip rules that cause errors
                continue
            }
        }
        
        return $results
    }
    catch {
        return @([PSCustomObject]@{
            Name = "Error"
            DisplayName = "Error: $($_.Exception.Message)"
        })
    }
}

# Function to export firewall rule to a file
function Export-FirewallRule {
    param (
        [PSCustomObject]$Rule,
        [string]$OutputPath
    )
    
    # Create a structured format that can be imported by GPO-Creator
    $exportObject = @{
        Type = "FirewallRule"
        Name = $Rule.Name
        DisplayName = $Rule.DisplayName
        Description = $Rule.Description
        Direction = $Rule.Direction
        Action = $Rule.Action
        Protocol = $Rule.Protocol
        LocalPorts = $Rule.LocalPorts
        RemotePorts = $Rule.RemotePorts
        LocalAddresses = $Rule.LocalAddresses
        RemoteAddresses = $Rule.RemoteAddresses
        Program = $Rule.Program
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    # Export as JSON for easy import
    $exportObject | ConvertTo-Json | Out-File -FilePath $OutputPath
    
    # Also create human-readable text version
    $textOutputPath = [System.IO.Path]::ChangeExtension($OutputPath, "txt")
    $content = @"
Firewall Rule Information
Generated on: $(Get-Date)

Rule Name: $($Rule.Name)
Display Name: $($Rule.DisplayName)
Description: $($Rule.Description)
Direction: $($Rule.Direction)
Action: $($Rule.Action)
Enabled: $($Rule.Enabled)
Protocol: $($Rule.Protocol)
Local Ports: $($Rule.LocalPorts)
Remote Ports: $($Rule.RemotePorts)
Local Addresses: $($Rule.LocalAddresses)
Remote Addresses: $($Rule.RemoteAddresses)
Program: $($Rule.Program)

This information can be used for creating GPOs to configure firewall rules.
"@
    $content | Out-File -FilePath $textOutputPath
    
    return $OutputPath
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Firewall Rule Extractor"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Create filter controls
$groupBoxFilter = New-Object System.Windows.Forms.GroupBox
$groupBoxFilter.Text = "Filters"
$groupBoxFilter.Location = New-Object System.Drawing.Point(20, 20)
$groupBoxFilter.Size = New-Object System.Drawing.Size(940, 100)
$form.Controls.Add($groupBoxFilter)

$labelDisplayName = New-Object System.Windows.Forms.Label
$labelDisplayName.Text = "Display Name:"
$labelDisplayName.Location = New-Object System.Drawing.Point(10, 25)
$labelDisplayName.AutoSize = $true
$groupBoxFilter.Controls.Add($labelDisplayName)

$textDisplayName = New-Object System.Windows.Forms.TextBox
$textDisplayName.Location = New-Object System.Drawing.Point(100, 22)
$textDisplayName.Size = New-Object System.Drawing.Size(300, 20)
$textDisplayName.Text = "*"
$groupBoxFilter.Controls.Add($textDisplayName)

$labelDirection = New-Object System.Windows.Forms.Label
$labelDirection.Text = "Direction:"
$labelDirection.Location = New-Object System.Drawing.Point(430, 25)
$labelDirection.AutoSize = $true
$groupBoxFilter.Controls.Add($labelDirection)

$comboDirection = New-Object System.Windows.Forms.ComboBox
$comboDirection.Items.Add("*")
$comboDirection.Items.Add("Inbound")
$comboDirection.Items.Add("Outbound")
$comboDirection.SelectedIndex = 0
$comboDirection.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboDirection.Location = New-Object System.Drawing.Point(500, 22)
$comboDirection.Size = New-Object System.Drawing.Size(150, 20)
$groupBoxFilter.Controls.Add($comboDirection)

$labelAction = New-Object System.Windows.Forms.Label
$labelAction.Text = "Action:"
$labelAction.Location = New-Object System.Drawing.Point(680, 25)
$labelAction.AutoSize = $true
$groupBoxFilter.Controls.Add($labelAction)

$comboAction = New-Object System.Windows.Forms.ComboBox
$comboAction.Items.Add("*")
$comboAction.Items.Add("Allow")
$comboAction.Items.Add("Block")
$comboAction.SelectedIndex = 0
$comboAction.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboAction.Location = New-Object System.Drawing.Point(740, 22)
$comboAction.Size = New-Object System.Drawing.Size(150, 20)
$groupBoxFilter.Controls.Add($comboAction)

$labelProtocol = New-Object System.Windows.Forms.Label
$labelProtocol.Text = "Protocol:"
$labelProtocol.Location = New-Object System.Drawing.Point(10, 60)
$labelProtocol.AutoSize = $true
$groupBoxFilter.Controls.Add($labelProtocol)

$comboProtocol = New-Object System.Windows.Forms.ComboBox
$comboProtocol.Items.Add("*")
$comboProtocol.Items.Add("TCP")
$comboProtocol.Items.Add("UDP")
$comboProtocol.Items.Add("ICMP")
$comboProtocol.SelectedIndex = 0
$comboProtocol.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboProtocol.Location = New-Object System.Drawing.Point(100, 57)
$comboProtocol.Size = New-Object System.Drawing.Size(150, 20)
$groupBoxFilter.Controls.Add($comboProtocol)

$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Text = "Search Rules"
$buttonSearch.Location = New-Object System.Drawing.Point(600, 55)
$buttonSearch.Size = New-Object System.Drawing.Size(150, 30)
$groupBoxFilter.Controls.Add($buttonSearch)

$buttonExport = New-Object System.Windows.Forms.Button
$buttonExport.Text = "Export Selected Rule"
$buttonExport.Location = New-Object System.Drawing.Point(760, 55)
$buttonExport.Size = New-Object System.Drawing.Size(150, 30)
$buttonExport.Enabled = $false
$groupBoxFilter.Controls.Add($buttonExport)

# Create a ListView for displaying firewall rules
$listViewRules = New-Object System.Windows.Forms.ListView
$listViewRules.Location = New-Object System.Drawing.Point(20, 130)
$listViewRules.Size = New-Object System.Drawing.Size(940, 250)
$listViewRules.View = [System.Windows.Forms.View]::Details
$listViewRules.FullRowSelect = $true
$listViewRules.Columns.Add("Display Name", 300)
$listViewRules.Columns.Add("Direction", 80)
$listViewRules.Columns.Add("Action", 80)
$listViewRules.Columns.Add("Protocol", 80)
$listViewRules.Columns.Add("Local Ports", 150)
$listViewRules.Columns.Add("Enabled", 80)
$form.Controls.Add($listViewRules)

# Create Group Box for rule details
$groupBoxDetails = New-Object System.Windows.Forms.GroupBox
$groupBoxDetails.Text = "Rule Details"
$groupBoxDetails.Location = New-Object System.Drawing.Point(20, 390)
$groupBoxDetails.Size = New-Object System.Drawing.Size(940, 260)
$form.Controls.Add($groupBoxDetails)

# Create labels and text boxes for rule details
$labelRuleName = New-Object System.Windows.Forms.Label
$labelRuleName.Text = "Rule Name:"
$labelRuleName.Location = New-Object System.Drawing.Point(10, 30)
$labelRuleName.AutoSize = $true
$groupBoxDetails.Controls.Add($labelRuleName)

$textRuleName = New-Object System.Windows.Forms.TextBox
$textRuleName.Location = New-Object System.Drawing.Point(110, 27)
$textRuleName.Size = New-Object System.Drawing.Size(350, 20)
$textRuleName.ReadOnly = $true
$groupBoxDetails.Controls.Add($textRuleName)

$labelDetailDisplayName = New-Object System.Windows.Forms.Label
$labelDetailDisplayName.Text = "Display Name:"
$labelDetailDisplayName.Location = New-Object System.Drawing.Point(10, 60)
$labelDetailDisplayName.AutoSize = $true
$groupBoxDetails.Controls.Add($labelDetailDisplayName)

$textDetailDisplayName = New-Object System.Windows.Forms.TextBox
$textDetailDisplayName.Location = New-Object System.Drawing.Point(110, 57)
$textDetailDisplayName.Size = New-Object System.Drawing.Size(350, 20)
$textDetailDisplayName.ReadOnly = $true
$groupBoxDetails.Controls.Add($textDetailDisplayName)

$labelDetailDirection = New-Object System.Windows.Forms.Label
$labelDetailDirection.Text = "Direction:"
$labelDetailDirection.Location = New-Object System.Drawing.Point(10, 90)
$labelDetailDirection.AutoSize = $true
$groupBoxDetails.Controls.Add($labelDetailDirection)

$textDetailDirection = New-Object System.Windows.Forms.TextBox
$textDetailDirection.Location = New-Object System.Drawing.Point(110, 87)
$textDetailDirection.Size = New-Object System.Drawing.Size(150, 20)
$textDetailDirection.ReadOnly = $true
$groupBoxDetails.Controls.Add($textDetailDirection)

$labelDetailAction = New-Object System.Windows.Forms.Label
$labelDetailAction.Text = "Action:"
$labelDetailAction.Location = New-Object System.Drawing.Point(280, 90)
$labelDetailAction.AutoSize = $true
$groupBoxDetails.Controls.Add($labelDetailAction)

$textDetailAction = New-Object System.Windows.Forms.TextBox
$textDetailAction.Location = New-Object System.Drawing.Point(350, 87)
$textDetailAction.Size = New-Object System.Drawing.Size(110, 20)
$textDetailAction.ReadOnly = $true
$groupBoxDetails.Controls.Add($textDetailAction)

$labelDetailProtocol = New-Object System.Windows.Forms.Label
$labelDetailProtocol.Text = "Protocol:"
$labelDetailProtocol.Location = New-Object System.Drawing.Point(480, 90)
$labelDetailProtocol.AutoSize = $true
$groupBoxDetails.Controls.Add($labelDetailProtocol)

$textDetailProtocol = New-Object System.Windows.Forms.TextBox
$textDetailProtocol.Location = New-Object System.Drawing.Point(550, 87)
$textDetailProtocol.Size = New-Object System.Drawing.Size(100, 20)
$textDetailProtocol.ReadOnly = $true
$groupBoxDetails.Controls.Add($textDetailProtocol)

$labelDetailLocalPorts = New-Object System.Windows.Forms.Label
$labelDetailLocalPorts.Text = "Local Ports:"
$labelDetailLocalPorts.Location = New-Object System.Drawing.Point(10, 120)
$labelDetailLocalPorts.AutoSize = $true
$groupBoxDetails.Controls.Add($labelDetailLocalPorts)

$textDetailLocalPorts = New-Object System.Windows.Forms.TextBox
$textDetailLocalPorts.Location = New-Object System.Drawing.Point(110, 117)
$textDetailLocalPorts.Size = New-Object System.Drawing.Size(350, 20)
$textDetailLocalPorts.ReadOnly = $true
$groupBoxDetails.Controls.Add($textDetailLocalPorts)

$labelDetailRemotePorts = New-Object System.Windows.Forms.Label
$labelDetailRemotePorts.Text = "Remote Ports:"
$labelDetailRemotePorts.Location = New-Object System.Drawing.Point(480, 120)
$labelDetailRemotePorts.AutoSize = $true
$groupBoxDetails.Controls.Add($labelDetailRemotePorts)

$textDetailRemotePorts = New-Object System.Windows.Forms.TextBox
$textDetailRemotePorts.Location = New-Object System.Drawing.Point(570, 117)
$textDetailRemotePorts.Size = New-Object System.Drawing.Size(350, 20)
$textDetailRemotePorts.ReadOnly = $true
$groupBoxDetails.Controls.Add($textDetailRemotePorts)

$labelDetailLocalAddresses = New-Object System.Windows.Forms.Label
$labelDetailLocalAddresses.Text = "Local Addresses:"
$labelDetailLocalAddresses.Location = New-Object System.Drawing.Point(10, 150)
$labelDetailLocalAddresses.AutoSize = $true
$groupBoxDetails.Controls.Add($labelDetailLocalAddresses)

$textDetailLocalAddresses = New-Object System.Windows.Forms.TextBox
$textDetailLocalAddresses.Location = New-Object System.Drawing.Point(110, 147)
$textDetailLocalAddresses.Size = New-Object System.Drawing.Size(350, 20)
$textDetailLocalAddresses.ReadOnly = $true
$groupBoxDetails.Controls.Add($textDetailLocalAddresses)

$labelDetailRemoteAddresses = New-Object System.Windows.Forms.Label
$labelDetailRemoteAddresses.Text = "Remote Addresses:"
$labelDetailRemoteAddresses.Location = New-Object System.Drawing.Point(480, 150)
$labelDetailRemoteAddresses.AutoSize = $true
$groupBoxDetails.Controls.Add($labelDetailRemoteAddresses)

$textDetailRemoteAddresses = New-Object System.Windows.Forms.TextBox
$textDetailRemoteAddresses.Location = New-Object System.Drawing.Point(570, 147)
$textDetailRemoteAddresses.Size = New-Object System.Drawing.Size(350, 20)
$textDetailRemoteAddresses.ReadOnly = $true
$groupBoxDetails.Controls.Add($textDetailRemoteAddresses)

$labelDetailProgram = New-Object System.Windows.Forms.Label
$labelDetailProgram.Text = "Program:"
$labelDetailProgram.Location = New-Object System.Drawing.Point(10, 180)
$labelDetailProgram.AutoSize = $true
$groupBoxDetails.Controls.Add($labelDetailProgram)

$textDetailProgram = New-Object System.Windows.Forms.TextBox
$textDetailProgram.Location = New-Object System.Drawing.Point(110, 177)
$textDetailProgram.Size = New-Object System.Drawing.Size(810, 20)
$textDetailProgram.ReadOnly = $true
$groupBoxDetails.Controls.Add($textDetailProgram)

$labelDetailDescription = New-Object System.Windows.Forms.Label
$labelDetailDescription.Text = "Description:"
$labelDetailDescription.Location = New-Object System.Drawing.Point(10, 210)
$labelDetailDescription.AutoSize = $true
$groupBoxDetails.Controls.Add($labelDetailDescription)

$textDetailDescription = New-Object System.Windows.Forms.TextBox
$textDetailDescription.Location = New-Object System.Drawing.Point(110, 207)
$textDetailDescription.Size = New-Object System.Drawing.Size(810, 40)
$textDetailDescription.Multiline = $true
$textDetailDescription.ReadOnly = $true
$groupBoxDetails.Controls.Add($textDetailDescription)

# Add event handler for the Search button
$buttonSearch.Add_Click({
    $listViewRules.Items.Clear()
    $buttonExport.Enabled = $false
    $displayNameFilter = $textDisplayName.Text
    $directionFilter = $comboDirection.SelectedItem.ToString()
    $actionFilter = $comboAction.SelectedItem.ToString()
    $protocolFilter = $comboProtocol.SelectedItem.ToString()
    
    # Create a progress dialog
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Searching..."
    $progressForm.Size = New-Object System.Drawing.Size(300, 100)
    $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $progressForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $progressForm.ControlBox = $false
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false
    
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(10, 20)
    $progressLabel.Size = New-Object System.Drawing.Size(280, 20)
    $progressLabel.Text = "Searching for firewall rules..."
    $progressForm.Controls.Add($progressLabel)
    
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 40)
    $progressBar.Size = New-Object System.Drawing.Size(270, 20)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progressForm.Controls.Add($progressBar)
    
    # Show the progress form in a non-blocking way
    $progressForm.Show()
    $progressForm.Refresh()
    
    # Get firewall rules
    $rules = Get-LocalFirewallRules -DisplayName $displayNameFilter -Direction $directionFilter -Action $actionFilter -Protocol $protocolFilter
    
    # Close progress form
    $progressForm.Close()
    
    # Process and display rules
    foreach ($rule in $rules) {
        $item = New-Object System.Windows.Forms.ListViewItem($rule.DisplayName)
        $item.SubItems.Add($rule.Direction)
        $item.SubItems.Add($rule.Action)
        $item.SubItems.Add($rule.Protocol)
        $item.SubItems.Add($rule.LocalPorts)
        $item.SubItems.Add($rule.Enabled)
        $item.Tag = $rule
        $listViewRules.Items.Add($item)
    }
    
    # Update status
    if ($rules.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No firewall rules found matching the specified criteria.", "No Results", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        [System.Windows.Forms.MessageBox]::Show("Found $($rules.Count) firewall rules matching the criteria.", "Search Results", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
}))

# Add event handler for the ListView selection change
$listViewRules.Add_SelectedIndexChanged({
    if ($listViewRules.SelectedItems.Count -gt 0) {
        $selectedRule = $listViewRules.SelectedItems[0].Tag
        
        # Update the detail fields with the selected rule information
        $textRuleName.Text = $selectedRule.Name
        $textDetailDisplayName.Text = $selectedRule.DisplayName
        $textDetailDirection.Text = $selectedRule.Direction
        $textDetailAction.Text = $selectedRule.Action
        $textDetailProtocol.Text = $selectedRule.Protocol
        $textDetailLocalPorts.Text = $selectedRule.LocalPorts
        $textDetailRemotePorts.Text = $selectedRule.RemotePorts
        $textDetailLocalAddresses.Text = $selectedRule.LocalAddresses
        $textDetailRemoteAddresses.Text = $selectedRule.RemoteAddresses
        $textDetailProgram.Text = $selectedRule.Program
        $textDetailDescription.Text = $selectedRule.Description
        
        # Enable the export button
        $buttonExport.Enabled = $true
    } else {
        # Clear the detail fields if no rule is selected
        $textRuleName.Clear()
        $textDetailDisplayName.Clear()
        $textDetailDirection.Clear()
        $textDetailAction.Clear()
        $textDetailProtocol.Clear()
        $textDetailLocalPorts.Clear()
        $textDetailRemotePorts.Clear()
        $textDetailLocalAddresses.Clear()
        $textDetailRemoteAddresses.Clear()
        $textDetailProgram.Clear()
        $textDetailDescription.Clear()
        
        # Disable the export button
        $buttonExport.Enabled = $false
    }
}))

# Add event handler for the Export button
$buttonExport.Add_Click({
    if ($listViewRules.SelectedItems.Count -gt 0) {
        $selectedRule = $listViewRules.SelectedItems[0].Tag
        
        # Create a SaveFileDialog to choose the export location
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "GPO Creator Files (*.gpodata)|*.gpodata|All files (*.*)|*.*"
        $saveDialog.Title = "Save Firewall Rule Information"
        $saveDialog.FileName = "FirewallRule_$($selectedRule.Name -replace '\W', '_')_$(Get-Date -Format 'yyyyMMdd_HHmmss').gpodata"
        $saveDialog.DefaultExt = "gpodata"
        
        # Show the dialog and process the result
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportPath = Export-FirewallRule -Rule $selectedRule -OutputPath $saveDialog.FileName
            [System.Windows.Forms.MessageBox]::Show("Information exported to:`n$exportPath`n`nA text version has also been created.", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
}))

# Show the form
$form.ShowDialog()
