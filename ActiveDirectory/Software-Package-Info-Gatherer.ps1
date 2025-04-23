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

# Function to get software package information from MSI file
function Get-MsiInformation {
    param (
        [string]$FilePath
    )
    try {
        if (-not (Test-Path $FilePath)) {
            throw "File not found: $FilePath"
        }

        if (-not ($FilePath -match '\.msi$')) {
            throw "File must be an MSI package"
        }

        # Use Windows Installer to extract information from the MSI
        $windowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        $database = $windowsInstaller.GetType().InvokeMember(
            "OpenDatabase", 
            "InvokeMethod", 
            $null, 
            $windowsInstaller, 
            @($FilePath, 0)
        )

        # Get properties from the MSI
        $query = "SELECT * FROM Property"
        $view = $database.GetType().InvokeMember(
            "OpenView",
            "InvokeMethod",
            $null,
            $database,
            @($query)
        )

        $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null)

        $properties = @{}
        $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)

        while ($record -ne $null) {
            $propertyName = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
            $propertyValue = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 2)
            $properties[$propertyName] = $propertyValue
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
        }

        # Close view and clean up COM objects
        $view.GetType().InvokeMember("Close", "InvokeMethod", $null, $view, $null)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($view) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($database) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstaller) | Out-Null

        return @{
            FilePath = $FilePath
            FileName = (Split-Path -Path $FilePath -Leaf)
            ProductName = $properties.ProductName
            ProductCode = $properties.ProductCode
            ProductVersion = $properties.ProductVersion
            Manufacturer = $properties.Manufacturer
            Properties = $properties
        }
    }
    catch {
        return @{
            FilePath = $FilePath
            Error = $_.Exception.Message
        }
    }
    finally {
        # Make sure we clean up COM objects
        if ($null -ne $view) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($view) }
        if ($null -ne $database) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($database) }
        if ($null -ne $windowsInstaller) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstaller) }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Function to export MSI information to a file
function Export-SoftwareInfo {
    param (
        [hashtable]$PackageInfo,
        [string]$OutputPath,
        [string]$DeploymentType = "Assigned"
    )
    
    # Create a structured format that can be imported by GPO-Creator
    $exportObject = @{
        Type = "SoftwareInstallation"
        PackagePath = $PackageInfo.FilePath
        PackageName = $PackageInfo.FileName
        ProductName = $PackageInfo.ProductName
        ProductCode = $PackageInfo.ProductCode
        ProductVersion = $PackageInfo.ProductVersion
        Manufacturer = $PackageInfo.Manufacturer
        DeploymentType = $DeploymentType
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    # Export as JSON for easy import
    $exportObject | ConvertTo-Json | Out-File -FilePath $OutputPath
    
    # Also create human-readable text version
    $textOutputPath = [System.IO.Path]::ChangeExtension($OutputPath, "txt")
    $content = @"
Software Package Information
Generated on: $(Get-Date)

Package Path: $($PackageInfo.FilePath)
Product Name: $($PackageInfo.ProductName)
Product Version: $($PackageInfo.ProductVersion)
Manufacturer: $($PackageInfo.Manufacturer)
Product Code: $($PackageInfo.ProductCode)
Deployment Type: $DeploymentType

This information can be used for creating GPOs to deploy software.
"@
    $content | Out-File -FilePath $textOutputPath
    
    return $OutputPath
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Software Package Information Gatherer"
$form.Size = New-Object System.Drawing.Size(800, 650) # Increased height to accommodate new controls
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Create a panel to hold all controls and enable scrolling
$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = [System.Windows.Forms.DockStyle]::Fill
$panel.AutoScroll = $true
$form.Controls.Add($panel)

# Create labels
$labelPackagePath = New-Object System.Windows.Forms.Label
$labelPackagePath.Text = "MSI Package Path:"
$labelPackagePath.Location = New-Object System.Drawing.Point(20, 20)
$labelPackagePath.AutoSize = $true
$panel.Controls.Add($labelPackagePath)

$textPackagePath = New-Object System.Windows.Forms.TextBox
$textPackagePath.Location = New-Object System.Drawing.Point(20, 40)
$textPackagePath.Size = New-Object System.Drawing.Size(650, 20)
$panel.Controls.Add($textPackagePath)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "Browse..."
$buttonBrowse.Location = New-Object System.Drawing.Point(680, 39)
$buttonBrowse.Size = New-Object System.Drawing.Size(80, 23)
$panel.Controls.Add($buttonBrowse)

$labelDeploymentType = New-Object System.Windows.Forms.Label
$labelDeploymentType.Text = "Deployment Type:"
$labelDeploymentType.Location = New-Object System.Drawing.Point(20, 70)
$labelDeploymentType.AutoSize = $true
$panel.Controls.Add($labelDeploymentType)

$comboDeploymentType = New-Object System.Windows.Forms.ComboBox
$comboDeploymentType.Items.Add("Published")
$comboDeploymentType.Items.Add("Assigned")
$comboDeploymentType.SelectedIndex = 1
$comboDeploymentType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboDeploymentType.Location = New-Object System.Drawing.Point(120, 70)
$comboDeploymentType.Size = New-Object System.Drawing.Size(200, 20)
$panel.Controls.Add($comboDeploymentType)

$buttonAnalyze = New-Object System.Windows.Forms.Button
$buttonAnalyze.Text = "Analyze Package"
$buttonAnalyze.Location = New-Object System.Drawing.Point(350, 70)
$buttonAnalyze.Size = New-Object System.Drawing.Size(120, 23)
$panel.Controls.Add($buttonAnalyze)

$buttonExport = New-Object System.Windows.Forms.Button
$buttonExport.Text = "Export Info"
$buttonExport.Location = New-Object System.Drawing.Point(480, 70)
$buttonExport.Size = New-Object System.Drawing.Size(120, 23)
$buttonExport.Enabled = $false
$panel.Controls.Add($buttonExport)

# Create Group Box for package details
$groupBoxDetails = New-Object System.Windows.Forms.GroupBox
$groupBoxDetails.Text = "Package Details"
$groupBoxDetails.Location = New-Object System.Drawing.Point(20, 110)
$groupBoxDetails.Size = New-Object System.Drawing.Size(740, 200)
$panel.Controls.Add($groupBoxDetails)

# Create labels and text boxes for package details
$labelProductName = New-Object System.Windows.Forms.Label
$labelProductName.Text = "Product Name:"
$labelProductName.Location = New-Object System.Drawing.Point(10, 30)
$labelProductName.AutoSize = $true
$groupBoxDetails.Controls.Add($labelProductName)

$textProductName = New-Object System.Windows.Forms.TextBox
$textProductName.Location = New-Object System.Drawing.Point(110, 30)
$textProductName.Size = New-Object System.Drawing.Size(300, 20)
$textProductName.ReadOnly = $true
$groupBoxDetails.Controls.Add($textProductName)

$labelProductVersion = New-Object System.Windows.Forms.Label
$labelProductVersion.Text = "Product Version:"
$labelProductVersion.Location = New-Object System.Drawing.Point(10, 60)
$labelProductVersion.AutoSize = $true
$groupBoxDetails.Controls.Add($labelProductVersion)

$textProductVersion = New-Object System.Windows.Forms.TextBox
$textProductVersion.Location = New-Object System.Drawing.Point(110, 60)
$textProductVersion.Size = New-Object System.Drawing.Size(300, 20)
$textProductVersion.ReadOnly = $true
$groupBoxDetails.Controls.Add($textProductVersion)

$labelManufacturer = New-Object System.Windows.Forms.Label
$labelManufacturer.Text = "Manufacturer:"
$labelManufacturer.Location = New-Object System.Drawing.Point(10, 90)
$labelManufacturer.AutoSize = $true
$groupBoxDetails.Controls.Add($labelManufacturer)

$textManufacturer = New-Object System.Windows.Forms.TextBox
$textManufacturer.Location = New-Object System.Drawing.Point(110, 90)
$textManufacturer.Size = New-Object System.Drawing.Size(300, 20)
$textManufacturer.ReadOnly = $true
$groupBoxDetails.Controls.Add($textManufacturer)

$labelProductCode = New-Object System.Windows.Forms.Label
$labelProductCode.Text = "Product Code:"
$labelProductCode.Location = New-Object System.Drawing.Point(10, 120)
$labelProductCode.AutoSize = $true
$groupBoxDetails.Controls.Add($labelProductCode)

$textProductCode = New-Object System.Windows.Forms.TextBox
$textProductCode.Location = New-Object System.Drawing.Point(110, 120)
$textProductCode.Size = New-Object System.Drawing.Size(600, 20)
$textProductCode.ReadOnly = $true
$groupBoxDetails.Controls.Add($textProductCode)

# Create a GroupBox for all MSI properties
$groupBoxProperties = New-Object System.Windows.Forms.GroupBox
$groupBoxProperties.Text = "All MSI Properties"
$groupBoxProperties.Location = New-Object System.Drawing.Point(20, 320)
$groupBoxProperties.Size = New-Object System.Drawing.Size(740, 230)
$panel.Controls.Add($groupBoxProperties)

# Create a ListView for displaying all MSI properties
$listViewProperties = New-Object System.Windows.Forms.ListView
$listViewProperties.Location = New-Object System.Drawing.Point(10, 20)
$listViewProperties.Size = New-Object System.Drawing.Size(720, 200)
$listViewProperties.View = [System.Windows.Forms.View]::Details
$listViewProperties.FullRowSelect = $true
$listViewProperties.Columns.Add("Property", 200)
$listViewProperties.Columns.Add("Value", 500)
$groupBoxProperties.Controls.Add($listViewProperties)

# Create a ListView for multiple MSI packages
$groupBoxMultiPackage = New-Object System.Windows.Forms.GroupBox
$groupBoxMultiPackage.Text = "Multiple Package Management"
$groupBoxMultiPackage.Location = New-Object System.Drawing.Point(20, 550)
$groupBoxMultiPackage.Size = New-Object System.Drawing.Size(740, 80)
$panel.Controls.Add($groupBoxMultiPackage)

$listViewPackages = New-Object System.Windows.Forms.ListView
$listViewPackages.Location = New-Object System.Drawing.Point(10, 20)
$listViewPackages.Size = New-Object System.Drawing.Size(600, 50)
$listViewPackages.View = [System.Windows.Forms.View]::Details
$listViewPackages.FullRowSelect = $true
$listViewPackages.Columns.Add("Package Path", 300)
$listViewPackages.Columns.Add("Product Name", 150)
$listViewPackages.Columns.Add("Version", 80)
$listViewPackages.Columns.Add("Deployment", 80)
$groupBoxMultiPackage.Controls.Add($listViewPackages)

$buttonAddToList = New-Object System.Windows.Forms.Button
$buttonAddToList.Text = "Add to List"
$buttonAddToList.Location = New-Object System.Drawing.Point(620, 20)
$buttonAddToList.Size = New-Object System.Drawing.Size(110, 23)
$buttonAddToList.Enabled = $false
$groupBoxMultiPackage.Controls.Add($buttonAddToList)

$buttonBatchExport = New-Object System.Windows.Forms.Button
$buttonBatchExport.Text = "Batch Export"
$buttonBatchExport.Location = New-Object System.Drawing.Point(620, 45)
$buttonBatchExport.Size = New-Object System.Drawing.Size(110, 23)
$buttonBatchExport.Enabled = $false
$groupBoxMultiPackage.Controls.Add($buttonBatchExport)

# Add event handler for the Browse button
$buttonBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "MSI Files (*.msi)|*.msi|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select MSI Package"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textPackagePath.Text = $openFileDialog.FileName
        $buttonAnalyze.PerformClick()
    }
})

# Add event handler for the Analyze button
$buttonAnalyze.Add_Click({
    $packagePath = $textPackagePath.Text.Trim()
    
    if (-not $packagePath -or -not (Test-Path $packagePath)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid MSI package file.", "Invalid File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Clear previous data
    $textProductName.Clear()
    $textProductVersion.Clear()
    $textManufacturer.Clear()
    $textProductCode.Clear()
    $listViewProperties.Items.Clear()
    
    # Get package information
    $packageInfo = Get-MsiInformation -FilePath $packagePath
    
    if ($packageInfo.ContainsKey("Error")) {
        [System.Windows.Forms.MessageBox]::Show("Error analyzing package: $($packageInfo.Error)", "Analysis Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Update form with package information
    $textProductName.Text = $packageInfo.ProductName
    $textProductVersion.Text = $packageInfo.ProductVersion
    $textManufacturer.Text = $packageInfo.Manufacturer
    $textProductCode.Text = $packageInfo.ProductCode
    
    # Add all properties to the ListView
    foreach ($prop in $packageInfo.Properties.GetEnumerator() | Sort-Object -Property Name) {
        $item = New-Object System.Windows.Forms.ListViewItem($prop.Name)
        $item.SubItems.Add($prop.Value)
        $listViewProperties.Items.Add($item)
    }
    
    # Store package info in the form's Tag property for later use
    $form.Tag = $packageInfo
    
    # Enable export buttons
    $buttonExport.Enabled = $true
    $buttonAddToList.Enabled = $true
})

# Add event handler for the Export button
$buttonExport.Add_Click({
    $packageInfo = $form.Tag
    $deploymentType = $comboDeploymentType.SelectedItem.ToString()
    
    if ($null -eq $packageInfo) {
        [System.Windows.Forms.MessageBox]::Show("Please analyze a package first.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Create a SaveFileDialog to choose the export location
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "GPO Creator Files (*.gpodata)|*.gpodata|All files (*.*)|*.*"
    $saveDialog.Title = "Save Software Package Information"
    $saveDialog.FileName = "Software_$($packageInfo.ProductName -replace '\W', '_')_$(Get-Date -Format 'yyyyMMdd_HHmmss').gpodata"
    $saveDialog.DefaultExt = "gpodata"
    
    # Show the dialog and process the result
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $exportPath = Export-SoftwareInfo -PackageInfo $packageInfo -OutputPath $saveDialog.FileName -DeploymentType $deploymentType
        [System.Windows.Forms.MessageBox]::Show("Information exported to:`n$exportPath`n`nA text version has also been created.", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Add event handler for Add to List button
$buttonAddToList.Add_Click({
    $packageInfo = $form.Tag
    $deploymentType = $comboDeploymentType.SelectedItem.ToString()
    
    if ($null -eq $packageInfo) {
        [System.Windows.Forms.MessageBox]::Show("Please analyze a package first.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Check if package already exists in the list
    $exists = $false
    foreach ($item in $listViewPackages.Items) {
        if ($item.Text -eq $packageInfo.FilePath) {
            $exists = $true
            break
        }
    }
    
    if ($exists) {
        [System.Windows.Forms.MessageBox]::Show("This package is already in the list.", "Duplicate", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    # Add to ListView
    $item = New-Object System.Windows.Forms.ListViewItem($packageInfo.FilePath)
    $item.SubItems.Add($packageInfo.ProductName)
    $item.SubItems.Add($packageInfo.ProductVersion)
    $item.SubItems.Add($deploymentType)
    $item.Tag = @{
        PackageInfo = $packageInfo
        DeploymentType = $deploymentType
    }
    $listViewPackages.Items.Add($item)
    
    # Enable batch export if there's at least one item
    if ($listViewPackages.Items.Count -gt 0) {
        $buttonBatchExport.Enabled = $true
    }
})

# Add event handler for Batch Export button
$buttonBatchExport.Add_Click({
    if ($listViewPackages.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No packages to export.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Create a folder browser dialog
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowserDialog.Description = "Select folder to save package information files"
    
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputFolder = $folderBrowserDialog.SelectedPath
        $exportedFiles = @()
        
        foreach ($item in $listViewPackages.Items) {
            $packageData = $item.Tag
            $packageInfo = $packageData.PackageInfo
            $deploymentType = $packageData.DeploymentType
            
            # Generate unique filename
            $fileName = "Software_$($packageInfo.ProductName -replace '\W', '_')_$(Get-Date -Format 'yyyyMMdd_HHmmss').gpodata"
            $outputPath = Join-Path -Path $outputFolder -ChildPath $fileName
            
            # Export package info
            $exportPath = Export-SoftwareInfo -PackageInfo $packageInfo -OutputPath $outputPath -DeploymentType $deploymentType
            $exportedFiles += $exportPath
        }
        
        [System.Windows.Forms.MessageBox]::Show("Exported $($exportedFiles.Count) package information files to:`n$outputFolder", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Show the form
$form.ShowDialog()
