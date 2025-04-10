# Load windows forms module
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Terminal Hide
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

# Debug mode toggle
$debug = $true  # Change this to $false to turn off debugging logs

# Function to write debug logs if debugging is enabled
function Write-DebugLog {
    param($message)
    if ($debug) {
        Write-Host "Debug: $message"
    }
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Network Profiles Manager"
$form.Width = 700
$form.Height = 750

# Set program icon from imageres.dll
$iconPath = "$($env:SystemRoot)\System32\imageres.dll"
$iconIndex = 109

try {
    if (Test-Path $iconPath) {
        $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
        $form.Icon = $icon
    } else {
        Write-DebugLog "imageres.dll not found at: $iconPath"
    }
} catch {
    Write-DebugLog "Failed to load icon: $_"
}

# Create main panel
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($mainPanel)

# Create button panel
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Height = 40
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$mainPanel.Controls.Add($buttonPanel)

# Create ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Dock = [System.Windows.Forms.DockStyle]::Fill
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.CheckBoxes = $true
$listView.MultiSelect = $true
$mainPanel.Controls.Add($listView)

# Add columns to ListView
$listView.Columns.Add("Profile ID", 300) | Out-Null
$listView.Columns.Add("Profile Name", 400) | Out-Null

# Track last selected index for shift-click selection
$script:lastSelectedIndex = -1

# Handle ListView item selection for shift-click functionality
$listView.Add_ItemSelectionChanged({
    $currentIndex = $_.ItemIndex
    
    # If shift key is pressed and we have a previous selection
    if ([System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Shift -and $script:lastSelectedIndex -ne -1) {
        $start = [Math]::Min($script:lastSelectedIndex, $currentIndex)
        $end = [Math]::Max($script:lastSelectedIndex, $currentIndex)
        
        # Select all items between last selected and current
        for ($i = $start; $i -le $end; $i++) {
            $listView.Items[$i].Selected = $true
        }
    }
    else {
        # Update last selected index
        $script:lastSelectedIndex = $currentIndex
    }
})

# Variables for sorting
$script:sortColumn = 0
$script:sortAscending = $true

# Add our custom ListView comparer class with proper assembly references
Add-Type @"
using System;
using System.Collections;
using System.Windows.Forms;

public class ListViewColumnSorter : IComparer
{
    public int ColumnToSort;
    public SortOrder OrderOfSort;

    public ListViewColumnSorter()
    {
        ColumnToSort = 0;
        OrderOfSort = SortOrder.Ascending;
    }

    public int Compare(object x, object y)
    {
        ListViewItem listviewX = (ListViewItem)x;
        ListViewItem listviewY = (ListViewItem)y;

        int compareResult = String.Compare(
            listviewX.SubItems[ColumnToSort].Text,
            listviewY.SubItems[ColumnToSort].Text);

        if (OrderOfSort == SortOrder.Ascending)
        {
            return compareResult;
        }
        else
        {
            return -compareResult;
        }
    }
}
"@ -ReferencedAssemblies ('System.Windows.Forms')

# Create and assign the comparer object
$script:lvwColumnSorter = New-Object ListViewColumnSorter
$listView.ListViewItemSorter = $script:lvwColumnSorter

# Add column click event for sorting
$listView.Add_ColumnClick({
    $clickedColumn = $_.Column
    
    # If clicked the same column, reverse the sort order
    if ($script:lvwColumnSorter.ColumnToSort -eq $clickedColumn) {
        # Reverse the current sort direction
        if ($script:lvwColumnSorter.OrderOfSort -eq [System.Windows.Forms.SortOrder]::Ascending) {
            $script:lvwColumnSorter.OrderOfSort = [System.Windows.Forms.SortOrder]::Descending
            $script:sortAscending = $false
        } else {
            $script:lvwColumnSorter.OrderOfSort = [System.Windows.Forms.SortOrder]::Ascending
            $script:sortAscending = $true
        }
    } else {
        # Set the column number that is to be sorted; default to ascending sort
        $script:lvwColumnSorter.ColumnToSort = $clickedColumn
        $script:lvwColumnSorter.OrderOfSort = [System.Windows.Forms.SortOrder]::Ascending
        $script:sortColumn = $clickedColumn
        $script:sortAscending = $true
    }
    
    Write-DebugLog "Sorting by column $($script:lvwColumnSorter.ColumnToSort), Ascending: $($script:sortAscending)"
    
    # Perform the sort with these new sort options
    $listView.Sort()
})

# Context menu for right-click actions
$contextMenu = New-Object System.Windows.Forms.ContextMenu

# Refresh button in context menu
$refreshMenuItem = New-Object System.Windows.Forms.MenuItem
$refreshMenuItem.Text = "Refresh"
$refreshMenuItem.Add_Click({
    Update-NetworkProfilesList
})
$contextMenu.MenuItems.Add($refreshMenuItem)

# Rename Profile menu item
$renameMenuItem = New-Object System.Windows.Forms.MenuItem
$renameMenuItem.Text = "Rename Profile"
$renameMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -eq 1) {
        Rename-SelectedProfile
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select exactly one profile to rename.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$contextMenu.MenuItems.Add($renameMenuItem)

# Delete Profiles menu item (combines both selection methods)
$deleteMenuItem = New-Object System.Windows.Forms.MenuItem
$deleteMenuItem.Text = "Delete Profiles"
$deleteMenuItem.Add_Click({
    Delete-SelectedProfiles
})
$contextMenu.MenuItems.Add($deleteMenuItem)

# Add right-click event handler
$listView.Add_MouseUp({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $contextMenu.Show($listView, $_.Location)
    }
})

# Create buttons
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh"
$refreshButton.Width = 100
$refreshButton.Location = New-Object System.Drawing.Point(10, 10)
$refreshButton.Add_Click({ Update-NetworkProfilesList })
$buttonPanel.Controls.Add($refreshButton)

$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Text = "Select All"
$selectAllButton.Width = 100
$selectAllButton.Location = New-Object System.Drawing.Point(120, 10)
$selectAllButton.Add_Click({ 
    foreach ($item in $listView.Items) {
        $item.Checked = $true
    }
})
$buttonPanel.Controls.Add($selectAllButton)

$unselectAllButton = New-Object System.Windows.Forms.Button
$unselectAllButton.Text = "Unselect All"
$unselectAllButton.Width = 100
$unselectAllButton.Location = New-Object System.Drawing.Point(230, 10)
$unselectAllButton.Add_Click({ 
    foreach ($item in $listView.Items) {
        $item.Checked = $false
    }
})
$buttonPanel.Controls.Add($unselectAllButton)

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "Delete Profiles"
$deleteButton.Width = 120
$deleteButton.Location = New-Object System.Drawing.Point(340, 10)
$deleteButton.Add_Click({
    Delete-SelectedProfiles
})
$buttonPanel.Controls.Add($deleteButton)

$renameButton = New-Object System.Windows.Forms.Button
$renameButton.Text = "Rename"
$renameButton.Width = 100
$renameButton.Location = New-Object System.Drawing.Point(470, 10)
$renameButton.Add_Click({
    if ($listView.SelectedItems.Count -eq 1) {
        Rename-SelectedProfile
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select exactly one profile to rename.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$buttonPanel.Controls.Add($renameButton)

# Function to rename selected profile
function Rename-SelectedProfile {
    if ($listView.SelectedItems.Count -ne 1) {
        return
    }
    
    $selectedItem = $listView.SelectedItems[0]
    $profileId = $selectedItem.SubItems[0].Text
    $currentName = $selectedItem.SubItems[1].Text
    
    # Create input box for new name
    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = "Rename Network Profile"
    $inputForm.Size = New-Object System.Drawing.Size(400, 150)
    $inputForm.StartPosition = "CenterParent"
    $inputForm.FormBorderStyle = "FixedDialog"
    $inputForm.MaximizeBox = $false
    $inputForm.MinimizeBox = $false
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter new name for profile:"
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(380, 20)
    $inputForm.Controls.Add($label)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 50)
    $textBox.Size = New-Object System.Drawing.Size(365, 20)
    $textBox.Text = $currentName
    $inputForm.Controls.Add($textBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(210, 80)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = "OK"
    $okButton.Add_Click({
        $inputForm.Tag = $textBox.Text
        $inputForm.Close()
    })
    $inputForm.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(300, 80)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({ $inputForm.Close() })
    $inputForm.Controls.Add($cancelButton)
    
    $inputForm.AcceptButton = $okButton
    $inputForm.CancelButton = $cancelButton
    
    # Show dialog and get result
    $inputForm.ShowDialog() | Out-Null
    $newName = $inputForm.Tag
    
    # If name is provided and changed, update the profile
    if ($newName -and $newName -ne $currentName) {
        try {
            $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles\$profileId"
            Set-ItemProperty -Path $registryPath -Name "ProfileName" -Value $newName
            Write-DebugLog "Renamed profile $profileId from '$currentName' to '$newName'"
            
            # Update the ListView item directly for immediate feedback
            $selectedItem.SubItems[1].Text = $newName
            
            # Sort the ListView again to maintain sort order
            $listView.Sort()
            
            [System.Windows.Forms.MessageBox]::Show("Profile renamed successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            Write-DebugLog "Failed to rename profile $profileId`: $_"
            [System.Windows.Forms.MessageBox]::Show("Failed to rename profile: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Function to delete selected profiles
function Delete-SelectedProfiles {
    # Collect items to delete from both checked and selected items
    $checkedItems = @($listView.Items | Where-Object { $_.Checked })
    $selectedItems = @($listView.SelectedItems)
    
    # Combine both lists and remove duplicates
    $itemsToDelete = @()
    $uniqueIDs = @{}
    
    foreach ($item in $checkedItems) {
        $id = $item.SubItems[0].Text
        if (-not $uniqueIDs.ContainsKey($id)) {
            $uniqueIDs[$id] = $true
            $itemsToDelete += $item
        }
    }
    
    foreach ($item in $selectedItems) {
        $id = $item.SubItems[0].Text
        if (-not $uniqueIDs.ContainsKey($id)) {
            $uniqueIDs[$id] = $true
            $itemsToDelete += $item
        }
    }
    
    if ($itemsToDelete.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No profiles selected or checked. Please select profiles using checkboxes or by clicking on them.", 
            "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete $($itemsToDelete.Count) profile(s)?", 
        "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $successCount = 0
        $failCount = 0
        
        foreach ($item in $itemsToDelete) {
            $profileId = $item.SubItems[0].Text
            try {
                Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles\$profileId" -Force -ErrorAction Stop
                Write-DebugLog "Deleted profile: $profileId"
                $successCount++
            } catch {
                Write-DebugLog "Failed to delete profile $profileId`: $_"
                $failCount++
            }
        }
        
        # Clear the ListView's sort to prevent indexing issues
        $script:lvwColumnSorter.ColumnToSort = 0
        $script:lvwColumnSorter.OrderOfSort = [System.Windows.Forms.SortOrder]::Ascending
        
        # Reset last selected index since items are being refreshed
        $script:lastSelectedIndex = -1
        
        # Perform a complete refresh to ensure registry state is properly reflected
        Update-NetworkProfilesList
        
        $message = "$successCount profiles were successfully deleted."
        if ($failCount -gt 0) {
            $message += " $failCount profiles could not be deleted."
        }
        
        [System.Windows.Forms.MessageBox]::Show($message, "Delete Results", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
}

# Function to update network profiles list
function Update-NetworkProfilesList {
    # Store currently selected items to restore selection after refresh
    $selectedIds = @()
    foreach ($selectedItem in $listView.SelectedItems) {
        $selectedIds += $selectedItem.SubItems[0].Text
    }
    
    # Store checked items to restore checked state after refresh
    $checkedIds = @()
    foreach ($item in $listView.Items | Where-Object { $_.Checked }) {
        $checkedIds += $item.SubItems[0].Text
    }
    
    # Clear and repopulate the ListView
    $listView.BeginUpdate()
    $listView.Items.Clear()
    
    try {
        $profiles = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles" -ErrorAction Stop
        foreach ($profile in $profiles) {
            try {
                $profileName = (Get-ItemProperty -Path $profile.PSPath -Name ProfileName).ProfileName
                $item = $listView.Items.Add($profile.PSChildName)
                $item.SubItems.Add($profileName)
                
                # Restore checked state
                if ($checkedIds -contains $profile.PSChildName) {
                    $item.Checked = $true
                }
                
                # Restore selection
                if ($selectedIds -contains $profile.PSChildName) {
                    $item.Selected = $true
                }
            } catch {
                Write-DebugLog "Error reading profile name for: $($profile.PSChildName) - $_"
            }
        }
    } catch {
        Write-DebugLog "Error retrieving profiles: $_"
        [System.Windows.Forms.MessageBox]::Show("Failed to retrieve network profiles: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    
    $listView.EndUpdate()
    
    # Apply current sort settings
    $script:lvwColumnSorter.ColumnToSort = $script:sortColumn
    if ($script:sortAscending) {
        $script:lvwColumnSorter.OrderOfSort = [System.Windows.Forms.SortOrder]::Ascending
    } else {
        $script:lvwColumnSorter.OrderOfSort = [System.Windows.Forms.SortOrder]::Descending
    }
    $listView.Sort()
}

# Add ConsoleManager class for hiding console window if needed
if (-not ([System.Management.Automation.PSTypeName]'ConsoleManager').Type) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class ConsoleManager {
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();
        
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        
        public const int SW_HIDE = 0;
        public const int SW_SHOW = 5;
    }
"@
}

# Load event to populate profiles on startup
$form.Add_Load({
    Update-NetworkProfilesList
})

# Ensure console is visible if debugging
if (-not $debug) {
    [ConsoleManager]::ShowWindow([ConsoleManager]::GetConsoleWindow(), [ConsoleManager]::SW_HIDE)
}

# Show the form
$form.ShowDialog() | Out-Null
