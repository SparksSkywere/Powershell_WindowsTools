# Add Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Printer Driver Manager"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::DPI
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)

# Create a ListView to display printer drivers
$listView = New-Object System.Windows.Forms.ListView
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Location = New-Object System.Drawing.Point(10, 10)
$listView.Size = New-Object System.Drawing.Size(860, 200)
$listView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Add columns to the ListView with proportional widths
$listView.Columns.Add("Name", 300) | Out-Null
$listView.Columns.Add("Environment", 180) | Out-Null
$listView.Columns.Add("Version", 120) | Out-Null
$listView.Columns.Add("Manufacturer", 260) | Out-Null

# Create label for printer list
$printerListLabel = New-Object System.Windows.Forms.Label
$printerListLabel.Location = New-Object System.Drawing.Point(10, 220)
$printerListLabel.Size = New-Object System.Drawing.Size(860, 20)
$printerListLabel.Text = "Printers using selected driver:"
$printerListLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$printerListLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Create a ListView to display printers associated with driver
$printerListView = New-Object System.Windows.Forms.ListView
$printerListView.View = [System.Windows.Forms.View]::Details
$printerListView.FullRowSelect = $true
$printerListView.GridLines = $true
$printerListView.Location = New-Object System.Drawing.Point(10, 245)
$printerListView.Size = New-Object System.Drawing.Size(860, 170)
$printerListView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Add columns to the ListView
$printerListView.Columns.Add("Printer Name", 350) | Out-Null
$printerListView.Columns.Add("Port", 400) | Out-Null
$printerListView.Columns.Add("Shared", 110) | Out-Null

# Create a refresh button
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(10, 430)
$refreshButton.Size = New-Object System.Drawing.Size(120, 30)
$refreshButton.Text = "Refresh List"
$refreshButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

# Create an add driver button
$addButton = New-Object System.Windows.Forms.Button
$addButton.Location = New-Object System.Drawing.Point(140, 430)
$addButton.Size = New-Object System.Drawing.Size(120, 30)
$addButton.Text = "Add Driver"

# Create a delete driver button
$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Location = New-Object System.Drawing.Point(270, 430)
$deleteButton.Size = New-Object System.Drawing.Size(120, 30)
$deleteButton.Text = "Delete Driver"

# Create printer management buttons
$addPrinterButton = New-Object System.Windows.Forms.Button
$addPrinterButton.Location = New-Object System.Drawing.Point(410, 430)
$addPrinterButton.Size = New-Object System.Drawing.Size(120, 30)
$addPrinterButton.Text = "Add Printer"

$deletePrinterButton = New-Object System.Windows.Forms.Button
$deletePrinterButton.Location = New-Object System.Drawing.Point(540, 430)
$deletePrinterButton.Size = New-Object System.Drawing.Size(120, 30)
$deletePrinterButton.Text = "Delete Printer"

# Create a status bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusBar.Items.Add($statusLabel) | Out-Null

# Create context menu for right-click options
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$addMenuItem = $contextMenu.Items.Add("Add Driver")
$deleteMenuItem = $contextMenu.Items.Add("Delete Driver")
$refreshMenuItem = $contextMenu.Items.Add("Refresh List")
$showPrintersMenuItem = $contextMenu.Items.Add("Show Associated Printers")
$contextMenu.Items.Add("-") | Out-Null
$listView.ContextMenuStrip = $contextMenu

# Create context menu for printer list
$printerContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$addPrinterMenuItem = $printerContextMenu.Items.Add("Add New Printer")
$deletePrinterMenuItem = $printerContextMenu.Items.Add("Delete Printer")
$printerListView.ContextMenuStrip = $printerContextMenu

# Function to resize columns to fit content and available width
function Resize-ListViewColumns {
    param (
        [System.Windows.Forms.ListView]$ListView
    )
    
    # First pass: Size to content
    $ListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)
    
    # Calculate total width of all columns
    $totalColumnWidth = 0
    for ($i = 0; $i -lt $ListView.Columns.Count; $i++) {
        $totalColumnWidth += $ListView.Columns[$i].Width
    }
    
    # Second pass: Scale proportionally if needed
    $listViewWidth = $ListView.ClientSize.Width
    if ($totalColumnWidth -lt $listViewWidth) {
        $ratio = $listViewWidth / $totalColumnWidth
        for ($i = 0; $i -lt $ListView.Columns.Count; $i++) {
            $ListView.Columns[$i].Width = [int]($ListView.Columns[$i].Width * $ratio)
        }
    }
}

# Function to get all printer drivers
function Get-PrinterDriversList {
    $statusLabel.Text = "Loading printer drivers..."
    $form.Refresh()
    
    # Clear existing items
    $listView.Items.Clear()
    
    try {
        $drivers = Get-PrinterDriver | Sort-Object Name
        
        foreach ($driver in $drivers) {
            $item = New-Object System.Windows.Forms.ListViewItem($driver.Name)
            $item.SubItems.Add($driver.PrinterEnvironment)
            $item.SubItems.Add($driver.MajorVersion)
            $item.SubItems.Add($driver.Manufacturer)
            $listView.Items.Add($item)
        }
        
        $statusLabel.Text = "Found $($drivers.Count) printer drivers."
        
        # Auto-resize columns to fit content
        Resize-ListViewColumns -ListView $listView
    }
    catch {
        $statusLabel.Text = "Error: $_"
        [System.Windows.Forms.MessageBox]::Show("Failed to retrieve printer drivers: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to add a new printer driver (calls pnputil for driver installation)
function Add-PrinterDriverUI {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Driver Files (*.inf)|*.inf|All files (*.*)|*.*"
    $openFileDialog.Title = "Select Printer Driver INF file"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $statusLabel.Text = "Installing printer driver..."
            $form.Refresh()
            
            $infPath = $openFileDialog.FileName
            $result = pnputil.exe -i -a $infPath
            
            [System.Windows.Forms.MessageBox]::Show("Driver installation initiated. Check the output:`n`n$result", "Driver Installation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # Refresh the list
            Get-PrinterDriversList
        }
        catch {
            $statusLabel.Text = "Error: $_"
            [System.Windows.Forms.MessageBox]::Show("Failed to install printer driver: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Function to delete a printer driver
function Remove-PrinterDriverUI {
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a printer driver to delete.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $selectedDriver = $listView.SelectedItems[0].Text
    
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to delete the driver '$selectedDriver'?",
        "Confirm Deletion",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            $statusLabel.Text = "Removing printer driver '$selectedDriver'..."
            $form.Refresh()
            
            Remove-PrinterDriver -Name $selectedDriver -ErrorAction Stop
            
            $statusLabel.Text = "Driver '$selectedDriver' removed successfully."
            Get-PrinterDriversList
        }
        catch {
            $statusLabel.Text = "Error: $_"
            [System.Windows.Forms.MessageBox]::Show("Failed to remove printer driver: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Function to show printers using the selected driver
function Show-PrintersForDriver {
    if ($listView.SelectedItems.Count -eq 0) {
        $printerListView.Items.Clear()
        $statusLabel.Text = "Please select a driver to view associated printers."
        return
    }
    
    $selectedDriver = $listView.SelectedItems[0].Text
    $printerListView.Items.Clear()
    
    try {
        $printers = Get-Printer | Where-Object { $_.DriverName -eq $selectedDriver } | Sort-Object Name
        
        foreach ($printer in $printers) {
            $item = New-Object System.Windows.Forms.ListViewItem($printer.Name)
            $item.SubItems.Add($printer.PortName)
            $item.SubItems.Add($printer.Shared.ToString())
            $printerListView.Items.Add($item)
        }
        
        $statusLabel.Text = "Found $($printers.Count) printers using driver '$selectedDriver'."
        
        # Auto-resize columns to fit content
        Resize-ListViewColumns -ListView $printerListView
    }
    catch {
        $statusLabel.Text = "Error: $_"
        [System.Windows.Forms.MessageBox]::Show("Failed to retrieve printers: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to add a new printer
function Add-PrinterUI {
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a driver to use for the new printer.", "No Driver Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $selectedDriver = $listView.SelectedItems[0].Text
    
    # Create a form for printer details
    $printerForm = New-Object System.Windows.Forms.Form
    $printerForm.Text = "Add New Printer"
    $printerForm.Size = New-Object System.Drawing.Size(400, 300)
    $printerForm.StartPosition = "CenterScreen"
    $printerForm.FormBorderStyle = "FixedDialog"
    $printerForm.MaximizeBox = $false
    $printerForm.MinimizeBox = $false
    
    # Printer Name
    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Location = New-Object System.Drawing.Point(20, 20)
    $nameLabel.Size = New-Object System.Drawing.Size(150, 20)
    $nameLabel.Text = "Printer Name:"
    
    $nameTextBox = New-Object System.Windows.Forms.TextBox
    $nameTextBox.Location = New-Object System.Drawing.Point(180, 20)
    $nameTextBox.Size = New-Object System.Drawing.Size(180, 20)
    
    # IP Address / Port
    $ipLabel = New-Object System.Windows.Forms.Label
    $ipLabel.Location = New-Object System.Drawing.Point(20, 50)
    $ipLabel.Size = New-Object System.Drawing.Size(150, 20)
    $ipLabel.Text = "IP Address:"
    
    $ipTextBox = New-Object System.Windows.Forms.TextBox
    $ipTextBox.Location = New-Object System.Drawing.Point(180, 50)
    $ipTextBox.Size = New-Object System.Drawing.Size(180, 20)
    
    # Port Type
    $portTypeLabel = New-Object System.Windows.Forms.Label
    $portTypeLabel.Location = New-Object System.Drawing.Point(20, 80)
    $portTypeLabel.Size = New-Object System.Drawing.Size(150, 20)
    $portTypeLabel.Text = "Port Type:"
    
    $portTypeComboBox = New-Object System.Windows.Forms.ComboBox
    $portTypeComboBox.Location = New-Object System.Drawing.Point(180, 80)
    $portTypeComboBox.Size = New-Object System.Drawing.Size(180, 20)
    $portTypeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $portTypeComboBox.Items.Add("Standard TCP/IP") | Out-Null
    $portTypeComboBox.Items.Add("WSD") | Out-Null
    $portTypeComboBox.SelectedIndex = 0
    
    # Shared checkbox
    $sharedCheckbox = New-Object System.Windows.Forms.CheckBox
    $sharedCheckbox.Location = New-Object System.Drawing.Point(20, 110)
    $sharedCheckbox.Size = New-Object System.Drawing.Size(150, 20)
    $sharedCheckbox.Text = "Share this printer"
    
    # Share name
    $shareNameLabel = New-Object System.Windows.Forms.Label
    $shareNameLabel.Location = New-Object System.Drawing.Point(20, 140)
    $shareNameLabel.Size = New-Object System.Drawing.Size(150, 20)
    $shareNameLabel.Text = "Share Name:"
    
    $shareNameTextBox = New-Object System.Windows.Forms.TextBox
    $shareNameTextBox.Location = New-Object System.Drawing.Point(180, 140)
    $shareNameTextBox.Size = New-Object System.Drawing.Size(180, 20)
    $shareNameTextBox.Enabled = $false
    
    # Driver info
    $driverLabel = New-Object System.Windows.Forms.Label
    $driverLabel.Location = New-Object System.Drawing.Point(20, 170)
    $driverLabel.Size = New-Object System.Drawing.Size(340, 20)
    $driverLabel.Text = "Using driver: $selectedDriver"
    
    # OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(80, 210)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.Text = "OK"
    
    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(200, 210)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.Text = "Cancel"
    
    # Event handlers
    $sharedCheckbox.Add_CheckedChanged({
        $shareNameTextBox.Enabled = $sharedCheckbox.Checked
    })
    
    $okButton.Add_Click({
        $printerName = $nameTextBox.Text.Trim()
        $ipAddress = $ipTextBox.Text.Trim()
        $portType = $portTypeComboBox.SelectedItem
        $shared = $sharedCheckbox.Checked
        $shareName = $shareNameTextBox.Text.Trim()
        
        if ([string]::IsNullOrEmpty($printerName) -or [string]::IsNullOrEmpty($ipAddress)) {
            [System.Windows.Forms.MessageBox]::Show("Printer name and IP address are required.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ($shared -and [string]::IsNullOrEmpty($shareName)) {
            $shareName = $printerName
            $shareNameTextBox.Text = $shareName
        }
        
        try {
            # Create port name with prefix "IP_"
            $portName = "IP_" + $ipAddress.Replace(".", "_")
            
            # Check if port already exists, create if it doesn't
            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $ipAddress -ErrorAction Stop
            }
            
            # Create the printer
            $newPrinter = Add-Printer -Name $printerName -DriverName $selectedDriver -PortName $portName -ErrorAction Stop
            
            # Share the printer if selected
            if ($shared) {
                Set-Printer -Name $printerName -Shared $true -ShareName $shareName -ErrorAction Stop
            }
            
            $statusLabel.Text = "Printer '$printerName' created successfully."
            [System.Windows.Forms.MessageBox]::Show("Printer '$printerName' created successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # Refresh the printer list
            Show-PrintersForDriver
            $printerForm.Close()
        }
        catch {
            $statusLabel.Text = "Error: $_"
            [System.Windows.Forms.MessageBox]::Show("Failed to create printer: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    
    $cancelButton.Add_Click({
        $printerForm.Close()
    })
    
    # Add controls to the form
    $printerForm.Controls.Add($nameLabel)
    $printerForm.Controls.Add($nameTextBox)
    $printerForm.Controls.Add($ipLabel)
    $printerForm.Controls.Add($ipTextBox)
    $printerForm.Controls.Add($portTypeLabel)
    $printerForm.Controls.Add($portTypeComboBox)
    $printerForm.Controls.Add($sharedCheckbox)
    $printerForm.Controls.Add($shareNameLabel)
    $printerForm.Controls.Add($shareNameTextBox)
    $printerForm.Controls.Add($driverLabel)
    $printerForm.Controls.Add($okButton)
    $printerForm.Controls.Add($cancelButton)
    
    # Show the form
    $printerForm.ShowDialog() | Out-Null
}

# Function to delete a printer
function Remove-PrinterUI {
    if ($printerListView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a printer to delete.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $selectedPrinter = $printerListView.SelectedItems[0].Text
    
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to delete the printer '$selectedPrinter'?",
        "Confirm Deletion",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            $statusLabel.Text = "Removing printer '$selectedPrinter'..."
            $form.Refresh()
            
            Remove-Printer -Name $selectedPrinter -ErrorAction Stop
            
            $statusLabel.Text = "Printer '$selectedPrinter' removed successfully."
            Show-PrintersForDriver
        }
        catch {
            $statusLabel.Text = "Error: $_"
            [System.Windows.Forms.MessageBox]::Show("Failed to remove printer: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Event handlers
$refreshButton.Add_Click({ 
    Get-PrinterDriversList 
})

$addButton.Add_Click({
    Add-PrinterDriverUI
})

$deleteButton.Add_Click({
    Remove-PrinterDriverUI
})

$addMenuItem.Add_Click({
    Add-PrinterDriverUI
})

$deleteMenuItem.Add_Click({
    Remove-PrinterDriverUI
})

$refreshMenuItem.Add_Click({
    Get-PrinterDriversList
})

$showPrintersMenuItem.Add_Click({
    Show-PrintersForDriver
})

$addPrinterButton.Add_Click({
    Add-PrinterUI
})

$deletePrinterButton.Add_Click({
    Remove-PrinterUI
})

$addPrinterMenuItem.Add_Click({
    Add-PrinterUI
})

$deletePrinterMenuItem.Add_Click({
    Remove-PrinterUI
})

$listView.Add_SelectedIndexChanged({
    Show-PrintersForDriver
})

# Event handler for form resize
$form.Add_Resize({
    # Only call resize if the window state is normal (not minimized or maximized)
    if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) {
        # Adjust button positions based on form width
        if ($form.Width -ge 800) {
            $refreshButton.Location = New-Object System.Drawing.Point(10, 430)
            $addButton.Location = New-Object System.Drawing.Point(140, 430)
            $deleteButton.Location = New-Object System.Drawing.Point(270, 430)
            $addPrinterButton.Location = New-Object System.Drawing.Point(410, 430)
            $deletePrinterButton.Location = New-Object System.Drawing.Point(540, 430)
        }
        
        # Resize list view columns to fit new form width
        Resize-ListViewColumns -ListView $listView
        Resize-ListViewColumns -ListView $printerListView
    }
})

# Add controls to the form
$form.Controls.Add($listView)
$form.Controls.Add($printerListLabel)
$form.Controls.Add($printerListView)
$form.Controls.Add($refreshButton)
$form.Controls.Add($addButton)
$form.Controls.Add($deleteButton)
$form.Controls.Add($addPrinterButton)
$form.Controls.Add($deletePrinterButton)
$form.Controls.Add($statusBar)

# Load the driver list when the form loads
$form.Add_Shown({
    Get-PrinterDriversList
})

# Show the form
$form.ShowDialog() | Out-Null
