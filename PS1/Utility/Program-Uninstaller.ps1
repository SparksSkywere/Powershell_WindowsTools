# Add Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Program Uninstaller"
$form.Size = New-Object System.Drawing.Size(1200, 700)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::DPI
$form.MinimumSize = New-Object System.Drawing.Size(1000, 600)

# Create a ListView to display installed programs
$listView = New-Object System.Windows.Forms.ListView
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Location = New-Object System.Drawing.Point(10, 50)
$listView.Size = New-Object System.Drawing.Size(1160, 400)
$listView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

# Add columns to the ListView
$listView.Columns.Add("Program Name", 250) | Out-Null
$listView.Columns.Add("Version", 100) | Out-Null
$listView.Columns.Add("Publisher", 180) | Out-Null
$listView.Columns.Add("Install Date", 100) | Out-Null
$listView.Columns.Add("Install Location", 280) | Out-Null
$listView.Columns.Add("Size (MB)", 80) | Out-Null
$listView.Columns.Add("Uninstall String", 200) | Out-Null

# Create search label and textbox
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Location = New-Object System.Drawing.Point(10, 20)
$searchLabel.Size = New-Object System.Drawing.Size(80, 20)
$searchLabel.Text = "Search:"
$searchLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$searchTextBox = New-Object System.Windows.Forms.TextBox
$searchTextBox.Location = New-Object System.Drawing.Point(90, 18)
$searchTextBox.Size = New-Object System.Drawing.Size(200, 25)

# Create buttons
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(300, 17)
$refreshButton.Size = New-Object System.Drawing.Size(100, 28)
$refreshButton.Text = "Refresh List"
$refreshButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$uninstallButton = New-Object System.Windows.Forms.Button
$uninstallButton.Location = New-Object System.Drawing.Point(10, 460)
$uninstallButton.Size = New-Object System.Drawing.Size(120, 30)
$uninstallButton.Text = "Uninstall"
$uninstallButton.BackColor = [System.Drawing.Color]::LightCoral
$uninstallButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(140, 460)
$exportButton.Size = New-Object System.Drawing.Size(120, 30)
$exportButton.Text = "Export List"
$exportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left

$detailsButton = New-Object System.Windows.Forms.Button
$detailsButton.Location = New-Object System.Drawing.Point(270, 460)
$detailsButton.Size = New-Object System.Drawing.Size(120, 30)
$detailsButton.Text = "Show Details"
$detailsButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left

# Create filter options
$filterLabel = New-Object System.Windows.Forms.Label
$filterLabel.Location = New-Object System.Drawing.Point(450, 20)
$filterLabel.Size = New-Object System.Drawing.Size(80, 20)
$filterLabel.Text = "Show:"
$filterLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$showAllRadio = New-Object System.Windows.Forms.RadioButton
$showAllRadio.Location = New-Object System.Drawing.Point(530, 18)
$showAllRadio.Size = New-Object System.Drawing.Size(80, 25)
$showAllRadio.Text = "All"
$showAllRadio.Checked = $true

$showInstalledRadio = New-Object System.Windows.Forms.RadioButton
$showInstalledRadio.Location = New-Object System.Drawing.Point(610, 18)
$showInstalledRadio.Size = New-Object System.Drawing.Size(110, 25)
$showInstalledRadio.Text = "User Installed"

# Create a status bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusBar.Items.Add($statusLabel) | Out-Null

# Create context menu for right-click options
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$uninstallMenuItem = $contextMenu.Items.Add("Uninstall Program")
$detailsMenuItem = $contextMenu.Items.Add("Show Details")
$contextMenu.Items.Add("-") | Out-Null
$openLocationMenuItem = $contextMenu.Items.Add("Open Install Location")
$copyPathMenuItem = $contextMenu.Items.Add("Copy Install Path")
$contextMenu.Items.Add("-") | Out-Null
$searchOnlineMenuItem = $contextMenu.Items.Add("Search Online")
$refreshMenuItem = $contextMenu.Items.Add("Refresh List")
$listView.ContextMenuStrip = $contextMenu

# Global variable to store all programs
$global:allPrograms = @()

# Function to resize columns to fit content and available width
function Resize-ListViewColumns {
    param (
        [System.Windows.Forms.ListView]$ListView
    )
    
    # Auto-resize columns to content first
    $ListView.AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
    
    # Calculate total width and adjust if needed
    $totalWidth = 0
    for ($i = 0; $i -lt $ListView.Columns.Count; $i++) {
        $totalWidth += $ListView.Columns[$i].Width
    }
    
    $availableWidth = $ListView.ClientSize.Width
    if ($totalWidth -lt $availableWidth) {
        $ratio = $availableWidth / $totalWidth
        for ($i = 0; $i -lt $ListView.Columns.Count; $i++) {
            $ListView.Columns[$i].Width = [int]($ListView.Columns[$i].Width * $ratio)
        }
    }
}

# Function to get installed programs from registry
function Get-InstalledPrograms {
    param (
        [bool]$UserInstalledOnly = $false
    )
    
    $statusLabel.Text = "Loading installed programs..."
    $form.Refresh()
    
    $programs = @()
    
    # Registry paths to check
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    try {
        foreach ($path in $registryPaths) {
            Get-ItemProperty $path -ErrorAction SilentlyContinue | ForEach-Object {
                $displayName = $_.DisplayName
                if ($displayName -and $displayName.Trim() -ne "") {
                    # Skip system components if filtering for user-installed only
                    if ($UserInstalledOnly) {
                        if ($_.SystemComponent -eq 1 -or 
                            $_.ParentKeyName -or 
                            $_.WindowsInstaller -eq 1 -or
                            $displayName -match "^(Microsoft|Windows|KB\d+|Security Update|Update for)" -or
                            $_.Publisher -match "^Microsoft") {
                            return
                        }
                    }
                    
                    # Parse install date
                    $installDate = ""
                    if ($_.InstallDate) {
                        try {
                            $dateStr = $_.InstallDate.ToString()
                            if ($dateStr.Length -eq 8) {
                                $year = $dateStr.Substring(0, 4)
                                $month = $dateStr.Substring(4, 2)
                                $day = $dateStr.Substring(6, 2)
                                $installDate = "$month/$day/$year"
                            }
                        }
                        catch {
                            $installDate = "Unknown"
                        }
                    }
                    
                    # Parse size
                    $sizeInMB = ""
                    if ($_.EstimatedSize) {
                        $sizeInMB = [math]::Round($_.EstimatedSize / 1024, 2)
                    }
                    
                    # Get install location
                    $installLocation = $_.InstallLocation
                    if (-not $installLocation) {
                        $installLocation = $_.InstallSource
                    }
                    
                    $program = [PSCustomObject]@{
                        Name = $displayName
                        Version = $_.DisplayVersion
                        Publisher = $_.Publisher
                        InstallDate = $installDate
                        InstallLocation = $installLocation
                        SizeMB = $sizeInMB
                        UninstallString = $_.UninstallString
                        QuietUninstallString = $_.QuietUninstallString
                        ModifyPath = $_.ModifyPath
                        HelpLink = $_.HelpLink
                        URLInfoAbout = $_.URLInfoAbout
                        Contact = $_.Contact
                        Comments = $_.Comments
                        RegistryKey = $_.PSPath
                    }
                    
                    $programs += $program
                }
            }
        }
        
        # Remove duplicates and sort
        $programs = $programs | Sort-Object Name -Unique
        $statusLabel.Text = "Found $($programs.Count) installed programs."
        
        return $programs
    }
    catch {
        $statusLabel.Text = "Error: $_"
        [System.Windows.Forms.MessageBox]::Show("Failed to retrieve installed programs: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return @()
    }
}

# Function to populate the ListView
function Update-ProgramList {
    param (
        [array]$Programs = $global:allPrograms,
        [string]$SearchFilter = ""
    )
    
    $listView.Items.Clear()
    
    $filteredPrograms = $Programs
    
    # Apply search filter
    if ($SearchFilter -and $SearchFilter.Trim() -ne "") {
        $filteredPrograms = $Programs | Where-Object { 
            $_.Name -like "*$SearchFilter*" -or 
            $_.Publisher -like "*$SearchFilter*" -or
            $_.Version -like "*$SearchFilter*"
        }
    }
    
    foreach ($program in $filteredPrograms) {
        $item = New-Object System.Windows.Forms.ListViewItem($program.Name)
        
        # Add null checks for all SubItems to prevent null reference exceptions
        $item.SubItems.Add($(if ($program.Version) { $program.Version } else { "" }))
        $item.SubItems.Add($(if ($program.Publisher) { $program.Publisher } else { "" }))
        $item.SubItems.Add($(if ($program.InstallDate) { $program.InstallDate } else { "" }))
        $item.SubItems.Add($(if ($program.InstallLocation) { $program.InstallLocation } else { "" }))
        $item.SubItems.Add($(if ($program.SizeMB) { $program.SizeMB } else { "" }))
        $item.SubItems.Add($(if ($program.UninstallString) { $program.UninstallString } else { "" }))
        
        $item.Tag = $program
        $listView.Items.Add($item)
    }
    
    Resize-ListViewColumns -ListView $listView
    $statusLabel.Text = "Displaying $($filteredPrograms.Count) programs."
}

# Function to refresh the program list
function Refresh-ProgramList {
    $userInstalledOnly = $showInstalledRadio.Checked
    $global:allPrograms = Get-InstalledPrograms -UserInstalledOnly $userInstalledOnly
    Update-ProgramList -Programs $global:allPrograms -SearchFilter $searchTextBox.Text
}

# Function to uninstall selected program
function Uninstall-SelectedProgram {
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a program to uninstall.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $selectedProgram = $listView.SelectedItems[0].Tag
    
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to uninstall '$($selectedProgram.Name)'?`n`nThis action cannot be undone.",
        "Confirm Uninstall",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            $uninstallString = $selectedProgram.UninstallString
            if (-not $uninstallString) {
                [System.Windows.Forms.MessageBox]::Show("No uninstall information found for this program.", "Cannot Uninstall", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            $statusLabel.Text = "Uninstalling '$($selectedProgram.Name)'..."
            $form.Refresh()
            
            # Parse and execute uninstall command
            if ($uninstallString -match '^"([^"]+)"(.*)$') {
                $executable = $matches[1]
                $arguments = $matches[2].Trim()
                Start-Process -FilePath $executable -ArgumentList $arguments -Wait
            }
            else {
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $uninstallString -Wait
            }
            
            $statusLabel.Text = "Uninstall process completed for '$($selectedProgram.Name)'."
            
            # Refresh the list after uninstall
            Start-Sleep -Seconds 2
            Refresh-ProgramList
        }
        catch {
            $statusLabel.Text = "Error: $_"
            [System.Windows.Forms.MessageBox]::Show("Failed to uninstall program: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Function to show program details
function Show-ProgramDetails {
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a program to view details.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $selectedProgram = $listView.SelectedItems[0].Tag
    
    # Create details form
    $detailsForm = New-Object System.Windows.Forms.Form
    $detailsForm.Text = "Program Details - $($selectedProgram.Name)"
    $detailsForm.Size = New-Object System.Drawing.Size(600, 500)
    $detailsForm.StartPosition = "CenterParent"
    $detailsForm.FormBorderStyle = "FixedDialog"
    $detailsForm.MaximizeBox = $false
    
    # Create text box for details
    $detailsTextBox = New-Object System.Windows.Forms.TextBox
    $detailsTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $detailsTextBox.Size = New-Object System.Drawing.Size(560, 420)
    $detailsTextBox.Multiline = $true
    $detailsTextBox.ScrollBars = "Vertical"
    $detailsTextBox.ReadOnly = $true
    $detailsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    
    # Build details text
    $details = @"
Program Name: $($selectedProgram.Name)
Version: $($selectedProgram.Version)
Publisher: $($selectedProgram.Publisher)
Install Date: $($selectedProgram.InstallDate)
Install Location: $($selectedProgram.InstallLocation)
Size: $($selectedProgram.SizeMB) MB
Contact: $($selectedProgram.Contact)
Help Link: $($selectedProgram.HelpLink)
Info URL: $($selectedProgram.URLInfoAbout)
Comments: $($selectedProgram.Comments)

Uninstall String:
$($selectedProgram.UninstallString)

Quiet Uninstall String:
$($selectedProgram.QuietUninstallString)

Modify Path:
$($selectedProgram.ModifyPath)

Registry Key:
$($selectedProgram.RegistryKey)
"@
    
    $detailsTextBox.Text = $details
    
    # Close button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(250, 440)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Text = "Close"
    $closeButton.Add_Click({ $detailsForm.Close() })
    
    $detailsForm.Controls.Add($detailsTextBox)
    $detailsForm.Controls.Add($closeButton)
    $detailsForm.ShowDialog() | Out-Null
}

# Function to export program list
function Export-ProgramList {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt"
    $saveFileDialog.Title = "Export Program List"
    $saveFileDialog.FileName = "InstalledPrograms_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $global:allPrograms | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Program list exported successfully to:`n$($saveFileDialog.FileName)", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to export program list: $_", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Event handlers
$refreshButton.Add_Click({ Refresh-ProgramList })
$uninstallButton.Add_Click({ Uninstall-SelectedProgram })
$detailsButton.Add_Click({ Show-ProgramDetails })
$exportButton.Add_Click({ Export-ProgramList })

$showAllRadio.Add_CheckedChanged({ if ($showAllRadio.Checked) { Refresh-ProgramList } })
$showInstalledRadio.Add_CheckedChanged({ if ($showInstalledRadio.Checked) { Refresh-ProgramList } })

$searchTextBox.Add_TextChanged({
    Update-ProgramList -Programs $global:allPrograms -SearchFilter $searchTextBox.Text
})

# Context menu event handlers
$uninstallMenuItem.Add_Click({ Uninstall-SelectedProgram })
$detailsMenuItem.Add_Click({ Show-ProgramDetails })
$refreshMenuItem.Add_Click({ Refresh-ProgramList })

$openLocationMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $location = $listView.SelectedItems[0].Tag.InstallLocation
        if ($location -and (Test-Path $location)) {
            Start-Process "explorer.exe" -ArgumentList $location
        } else {
            [System.Windows.Forms.MessageBox]::Show("Install location not found or not accessible.", "Location Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
})

$copyPathMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $location = $listView.SelectedItems[0].Tag.InstallLocation
        if ($location) {
            [System.Windows.Forms.Clipboard]::SetText($location)
            $statusLabel.Text = "Install path copied to clipboard."
        }
    }
})

$searchOnlineMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $programName = $listView.SelectedItems[0].Tag.Name
        $searchUrl = "https://www.google.com/search?q=" + [System.Uri]::EscapeDataString($programName)
        Start-Process $searchUrl
    }
})

# Form resize event
$form.Add_Resize({
    if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) {
        Resize-ListViewColumns -ListView $listView
    }
})

# Add controls to the form
$form.Controls.Add($searchLabel)
$form.Controls.Add($searchTextBox)
$form.Controls.Add($refreshButton)
$form.Controls.Add($filterLabel)
$form.Controls.Add($showAllRadio)
$form.Controls.Add($showInstalledRadio)
$form.Controls.Add($listView)
$form.Controls.Add($uninstallButton)
$form.Controls.Add($exportButton)
$form.Controls.Add($detailsButton)
$form.Controls.Add($statusBar)

# Load the program list when the form loads
$form.Add_Shown({
    Refresh-ProgramList
})

# Show the form
$form.ShowDialog() | Out-Null
