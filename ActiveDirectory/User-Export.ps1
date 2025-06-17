#Requires -RunAsAdministrator
#Requires -Modules ActiveDirectory

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

# Debug mode toggle
$debug = $true  # Change this to $false to turn off debugging logs

# Function to write debug logs if debugging is enabled
function Write-DebugLog {
    param($message)
    if ($debug) {
        Write-Host "Debug: $message"
    }
}

# Import required modules
if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-DebugLog "ActiveDirectory module imported successfully"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to import ActiveDirectory module. Make sure RSAT tools are installed.", "Module Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}

# Function to get all available domains
function Get-ADDomainList {
    try {
        $forest = Get-ADForest
        return $forest.Domains
    }
    catch {
        Write-DebugLog "Failed to get domain list: $_"
        return @((Get-ADDomain).DNSRoot)
    }
}

# Function to get all OUs in domain
function Get-ADOUList {
    param([string]$Domain)
    
    try {
        if ($Domain) {
            $ous = Get-ADOrganizationalUnit -Filter * -Server $Domain | Select-Object Name, DistinguishedName | Sort-Object Name
        } else {
            $ous = Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName | Sort-Object Name
        }
        return $ous
    }
    catch {
        Write-DebugLog "Failed to get OU list: $_"
        return @()
    }
}

# Function to get groups by type
function Get-ADGroupsByType {
    param(
        [string]$Domain = "",
        [string]$GroupType = "All"
    )
    
    try {
        $params = @{
            Properties = @("Name", "GroupCategory", "GroupScope", "Description")
        }
        
        if ($Domain) {
            $params.Server = $Domain
        }
        
        switch ($GroupType) {
            "Security" {
                $params.Filter = "GroupCategory -eq 'Security'"
            }
            "Distribution" {
                $params.Filter = "GroupCategory -eq 'Distribution'"
            }
            "DomainLocal" {
                $params.Filter = "GroupScope -eq 'DomainLocal'"
            }
            "Global" {
                $params.Filter = "GroupScope -eq 'Global'"
            }
            "Universal" {
                $params.Filter = "GroupScope -eq 'Universal'"
            }
            default {
                $params.Filter = "*"
            }
        }
        
        $groups = Get-ADGroup @params | Sort-Object Name
        return $groups
    }
    catch {
        Write-DebugLog "Failed to get groups: $_"
        return @()
    }
}

# Modified function to export users with group filtering
function Export-ADUsersToCSV {
    param(
        [string]$OutputPath,
        [string[]]$Properties,
        [string]$Filter = "*",
        [string]$SearchBase = "",
        [string]$Domain = "",
        [string]$GroupFilter = "",
        [string]$GroupFilterType = "MemberOf"
    )
    
    try {
        $params = @{
            Filter = $Filter
            Properties = $Properties
        }
        
        if ($SearchBase) {
            $params.SearchBase = $SearchBase
        }
        
        if ($Domain) {
            $params.Server = $Domain
        }
        
        Write-DebugLog "Exporting users with filter: $Filter"
        Write-DebugLog "Selected properties: $($Properties -join ', ')"
        
        $users = Get-ADUser @params
        
        # Apply group filtering if specified
        if ($GroupFilter -and $GroupFilter -ne "All Users" -and $GroupFilter -ne "(No Group Filter)") {
            Write-DebugLog "Applying group filter: $GroupFilter (Type: $GroupFilterType)"
            
            if ($GroupFilterType -eq "MemberOf") {
                # Filter users who are members of the selected group
                $groupMembers = Get-ADGroupMember -Identity $GroupFilter -Server $Domain -ErrorAction SilentlyContinue | Where-Object { $_.objectClass -eq 'user' }
                $memberSamAccountNames = $groupMembers | ForEach-Object { $_.SamAccountName }
                $users = $users | Where-Object { $_.SamAccountName -in $memberSamAccountNames }
            }
            elseif ($GroupFilterType -eq "NotMemberOf") {
                # Filter users who are NOT members of the selected group
                $groupMembers = Get-ADGroupMember -Identity $GroupFilter -Server $Domain -ErrorAction SilentlyContinue | Where-Object { $_.objectClass -eq 'user' }
                $memberSamAccountNames = $groupMembers | ForEach-Object { $_.SamAccountName }
                $users = $users | Where-Object { $_.SamAccountName -notin $memberSamAccountNames }
            }
        }
        
        if ($users) {
            # Select only the requested properties
            $selectedUsers = $users | Select-Object $Properties
            $selectedUsers | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            return @{
                Success = $true
                Message = "Exported $($users.Count) users to $OutputPath"
                Count = $users.Count
            }
        } else {
            return @{
                Success = $false
                Message = "No users found matching the criteria"
                Count = 0
            }
        }
    }
    catch {
        Write-DebugLog "Export failed: $_"
        return @{
            Success = $false
            Message = "Export failed: $($_.Exception.Message)"
            Count = 0
        }
    }
}

# Function to export groups to CSV
function Export-ADGroupsToCSV {
    param(
        [string]$OutputPath,
        [string[]]$Properties,
        [string]$Filter = "*",
        [string]$SearchBase = "",
        [string]$Domain = ""
    )
    
    try {
        $params = @{
            Filter = $Filter
            Properties = $Properties
        }
        
        if ($SearchBase) {
            $params.SearchBase = $SearchBase
        }
        
        if ($Domain) {
            $params.Server = $Domain
        }
        
        Write-DebugLog "Exporting groups with filter: $Filter"
        Write-DebugLog "Selected properties: $($Properties -join ', ')"
        
        $groups = Get-ADGroup @params
        
        if ($groups) {
            # Select only the requested properties
            $selectedGroups = $groups | Select-Object $Properties
            $selectedGroups | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            return @{
                Success = $true
                Message = "Exported $($groups.Count) groups to $OutputPath"
                Count = $groups.Count
            }
        } else {
            return @{
                Success = $false
                Message = "No groups found matching the criteria"
                Count = 0
            }
        }
    }
    catch {
        Write-DebugLog "Export failed: $_"
        return @{
            Success = $false
            Message = "Export failed: $($_.Exception.Message)"
            Count = 0
        }
    }
}

# Function to export computers to CSV
function Export-ADComputersToCSV {
    param(
        [string]$OutputPath,
        [string[]]$Properties,
        [string]$Filter = "*",
        [string]$SearchBase = "",
        [string]$Domain = ""
    )
    
    try {
        $params = @{
            Filter = $Filter
            Properties = $Properties
        }
        
        if ($SearchBase) {
            $params.SearchBase = $SearchBase
        }
        
        if ($Domain) {
            $params.Server = $Domain
        }
        
        Write-DebugLog "Exporting computers with filter: $Filter"
        Write-DebugLog "Selected properties: $($Properties -join ', ')"
        
        $computers = Get-ADComputer @params
        
        if ($computers) {
            # Select only the requested properties
            $selectedComputers = $computers | Select-Object $Properties
            $selectedComputers | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            return @{
                Success = $true
                Message = "Exported $($computers.Count) computers to $OutputPath"
                Count = $computers.Count
            }
        } else {
            return @{
                Success = $false
                Message = "No computers found matching the criteria"
                Count = 0
            }
        }
    }
    catch {
        Write-DebugLog "Export failed: $_"
        return @{
            Success = $false
            Message = "Export failed: $($_.Exception.Message)"
            Count = 0
        }
    }
}

# Function to export group memberships
function Export-ADGroupMembershipsToCSV {
    param(
        [string]$OutputPath,
        [string]$GroupName = "*",
        [string]$Domain = ""
    )
    
    try {
        $params = @{
            Filter = "Name -like '$GroupName'"
        }
        
        if ($Domain) {
            $params.Server = $Domain
        }
        
        Write-DebugLog "Exporting group memberships for groups matching: $GroupName"
        $groups = Get-ADGroup @params
        
        $memberships = @()
        
        foreach ($group in $groups) {
            try {
                $members = Get-ADGroupMember -Identity $group.DistinguishedName -Server $Domain -ErrorAction SilentlyContinue
                foreach ($member in $members) {
                    $memberships += [PSCustomObject]@{
                        GroupName = $group.Name
                        GroupDN = $group.DistinguishedName
                        MemberName = $member.Name
                        MemberSamAccountName = $member.SamAccountName
                        MemberObjectClass = $member.ObjectClass
                        MemberDN = $member.DistinguishedName
                    }
                }
            }
            catch {
                Write-DebugLog "Failed to get members for group $($group.Name): $_"
            }
        }
        
        if ($memberships) {
            $memberships | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            return @{
                Success = $true
                Message = "Exported $($memberships.Count) group memberships to $OutputPath"
                Count = $memberships.Count
            }
        } else {
            return @{
                Success = $false
                Message = "No group memberships found"
                Count = 0
            }
        }
    }
    catch {
        Write-DebugLog "Export failed: $_"
        return @{
            Success = $false
            Message = "Export failed: $($_.Exception.Message)"
            Count = 0
        }
    }
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Active Directory Export Utility"
$form.Size = New-Object System.Drawing.Size(1100, 800)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false

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

# Create a tab control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(1065, 720)

# Create tab pages
$tabUsers = New-Object System.Windows.Forms.TabPage
$tabUsers.Text = "Users"
$tabGroups = New-Object System.Windows.Forms.TabPage
$tabGroups.Text = "Groups"
$tabComputers = New-Object System.Windows.Forms.TabPage
$tabComputers.Text = "Computers"
$tabMemberships = New-Object System.Windows.Forms.TabPage
$tabMemberships.Text = "Group Memberships"

$tabControl.Controls.Add($tabUsers)
$tabControl.Controls.Add($tabGroups)
$tabControl.Controls.Add($tabComputers)
$tabControl.Controls.Add($tabMemberships)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 740)
$statusLabel.Size = New-Object System.Drawing.Size(1065, 40)
$statusLabel.Text = "Ready to export Active Directory objects..."
$statusLabel.AutoSize = $false
$form.Controls.Add($statusLabel)

# ------------------------------
# Users Tab Content - Reorganized Layout
# ------------------------------

# First row - Domain and Group Filter Type
$labelDomainUsers = New-Object System.Windows.Forms.Label
$labelDomainUsers.Text = "Domain:"
$labelDomainUsers.Location = New-Object System.Drawing.Point(20, 20)
$labelDomainUsers.AutoSize = $true
$tabUsers.Controls.Add($labelDomainUsers)

$comboDomainUsers = New-Object System.Windows.Forms.ComboBox
$comboDomainUsers.Location = New-Object System.Drawing.Point(80, 20)
$comboDomainUsers.Size = New-Object System.Drawing.Size(250, 21)
$comboDomainUsers.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabUsers.Controls.Add($comboDomainUsers)

$labelGroupFilterTypeUsers = New-Object System.Windows.Forms.Label
$labelGroupFilterTypeUsers.Text = "Group Filter:"
$labelGroupFilterTypeUsers.Location = New-Object System.Drawing.Point(350, 20)
$labelGroupFilterTypeUsers.AutoSize = $true
$tabUsers.Controls.Add($labelGroupFilterTypeUsers)

$comboGroupFilterTypeUsers = New-Object System.Windows.Forms.ComboBox
$comboGroupFilterTypeUsers.Items.Add("All Users")
$comboGroupFilterTypeUsers.Items.Add("Members Of")
$comboGroupFilterTypeUsers.Items.Add("Not Members Of")
$comboGroupFilterTypeUsers.SelectedIndex = 0
$comboGroupFilterTypeUsers.Location = New-Object System.Drawing.Point(430, 20)
$comboGroupFilterTypeUsers.Size = New-Object System.Drawing.Size(120, 21)
$comboGroupFilterTypeUsers.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabUsers.Controls.Add($comboGroupFilterTypeUsers)

# Second row - OU and Group Type
$labelSearchBaseUsers = New-Object System.Windows.Forms.Label
$labelSearchBaseUsers.Text = "Search Base (OU):"
$labelSearchBaseUsers.Location = New-Object System.Drawing.Point(20, 60)
$labelSearchBaseUsers.AutoSize = $true
$tabUsers.Controls.Add($labelSearchBaseUsers)

$comboSearchBaseUsers = New-Object System.Windows.Forms.ComboBox
$comboSearchBaseUsers.Location = New-Object System.Drawing.Point(120, 60)
$comboSearchBaseUsers.Size = New-Object System.Drawing.Size(400, 21)
$comboSearchBaseUsers.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabUsers.Controls.Add($comboSearchBaseUsers)

$buttonRefreshOUUsers = New-Object System.Windows.Forms.Button
$buttonRefreshOUUsers.Text = "↻"
$buttonRefreshOUUsers.Location = New-Object System.Drawing.Point(530, 59)
$buttonRefreshOUUsers.Size = New-Object System.Drawing.Size(30, 23)
$tabUsers.Controls.Add($buttonRefreshOUUsers)

$labelGroupTypeUsers = New-Object System.Windows.Forms.Label
$labelGroupTypeUsers.Text = "Group Type:"
$labelGroupTypeUsers.Location = New-Object System.Drawing.Point(580, 60)
$labelGroupTypeUsers.AutoSize = $true
$tabUsers.Controls.Add($labelGroupTypeUsers)

$comboGroupTypeUsers = New-Object System.Windows.Forms.ComboBox
$comboGroupTypeUsers.Items.Add("All Groups")
$comboGroupTypeUsers.Items.Add("Security")
$comboGroupTypeUsers.Items.Add("Distribution")
$comboGroupTypeUsers.Items.Add("DomainLocal")
$comboGroupTypeUsers.Items.Add("Global")
$comboGroupTypeUsers.Items.Add("Universal")
$comboGroupTypeUsers.SelectedIndex = 0
$comboGroupTypeUsers.Location = New-Object System.Drawing.Point(650, 60)
$comboGroupTypeUsers.Size = New-Object System.Drawing.Size(120, 21)
$comboGroupTypeUsers.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabUsers.Controls.Add($comboGroupTypeUsers)

# Third row - Filter and Group Selection
$labelFilterUsers = New-Object System.Windows.Forms.Label
$labelFilterUsers.Text = "Filter:"
$labelFilterUsers.Location = New-Object System.Drawing.Point(20, 100)
$labelFilterUsers.AutoSize = $true
$tabUsers.Controls.Add($labelFilterUsers)

$textFilterUsers = New-Object System.Windows.Forms.TextBox
$textFilterUsers.Location = New-Object System.Drawing.Point(80, 100)
$textFilterUsers.Size = New-Object System.Drawing.Size(200, 21)
$textFilterUsers.Text = "*"
$tabUsers.Controls.Add($textFilterUsers)

$buttonEnabledUsers = New-Object System.Windows.Forms.Button
$buttonEnabledUsers.Text = "Enabled"
$buttonEnabledUsers.Location = New-Object System.Drawing.Point(290, 99)
$buttonEnabledUsers.Size = New-Object System.Drawing.Size(70, 23)
$tabUsers.Controls.Add($buttonEnabledUsers)

$buttonDisabledUsers = New-Object System.Windows.Forms.Button
$buttonDisabledUsers.Text = "Disabled"
$buttonDisabledUsers.Location = New-Object System.Drawing.Point(370, 99)
$buttonDisabledUsers.Size = New-Object System.Drawing.Size(70, 23)
$tabUsers.Controls.Add($buttonDisabledUsers)

$labelGroupFilterUsers = New-Object System.Windows.Forms.Label
$labelGroupFilterUsers.Text = "Specific Group:"
$labelGroupFilterUsers.Location = New-Object System.Drawing.Point(460, 100)
$labelGroupFilterUsers.AutoSize = $true
$tabUsers.Controls.Add($labelGroupFilterUsers)

$comboGroupFilterUsers = New-Object System.Windows.Forms.ComboBox
$comboGroupFilterUsers.Location = New-Object System.Drawing.Point(550, 100)
$comboGroupFilterUsers.Size = New-Object System.Drawing.Size(220, 21)
$comboGroupFilterUsers.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabUsers.Controls.Add($comboGroupFilterUsers)

# Properties selection section
$labelPropertiesUsers = New-Object System.Windows.Forms.Label
$labelPropertiesUsers.Text = "Select Properties to Export:"
$labelPropertiesUsers.Location = New-Object System.Drawing.Point(20, 140)
$labelPropertiesUsers.AutoSize = $true
$tabUsers.Controls.Add($labelPropertiesUsers)

$listBoxPropertiesUsers = New-Object System.Windows.Forms.CheckedListBox
$listBoxPropertiesUsers.Location = New-Object System.Drawing.Point(20, 160)
$listBoxPropertiesUsers.Size = New-Object System.Drawing.Size(400, 350)
$listBoxPropertiesUsers.CheckOnClick = $true
$tabUsers.Controls.Add($listBoxPropertiesUsers)

# Common user properties
$userProperties = @(
    "SamAccountName", "Name", "DisplayName", "GivenName", "Surname", "UserPrincipalName",
    "EmailAddress", "Title", "Department", "Company", "Manager", "Office", "OfficePhone",
    "MobilePhone", "Fax", "HomePhone", "StreetAddress", "City", "State", "PostalCode",
    "Country", "Description", "Created", "Modified", "LastLogonDate", "PasswordLastSet",
    "PasswordExpired", "PasswordNeverExpires", "Enabled", "LockedOut", "AccountExpirationDate",
    "DistinguishedName", "CanonicalName", "HomeDrive", "HomeDirectory", "ProfilePath",
    "ScriptPath", "LogonWorkstations", "PrimaryGroup", "MemberOf"
)

foreach ($prop in $userProperties) {
    $listBoxPropertiesUsers.Items.Add($prop, $true)
}

# Buttons for property selection - moved to the right
$buttonSelectAllUsers = New-Object System.Windows.Forms.Button
$buttonSelectAllUsers.Text = "Select All"
$buttonSelectAllUsers.Location = New-Object System.Drawing.Point(450, 160)
$buttonSelectAllUsers.Size = New-Object System.Drawing.Size(100, 30)
$tabUsers.Controls.Add($buttonSelectAllUsers)

$buttonSelectNoneUsers = New-Object System.Windows.Forms.Button
$buttonSelectNoneUsers.Text = "Select None"
$buttonSelectNoneUsers.Location = New-Object System.Drawing.Point(450, 200)
$buttonSelectNoneUsers.Size = New-Object System.Drawing.Size(100, 30)
$tabUsers.Controls.Add($buttonSelectNoneUsers)

$buttonSelectCommonUsers = New-Object System.Windows.Forms.Button
$buttonSelectCommonUsers.Text = "Common Only"
$buttonSelectCommonUsers.Location = New-Object System.Drawing.Point(450, 240)
$buttonSelectCommonUsers.Size = New-Object System.Drawing.Size(100, 30)
$tabUsers.Controls.Add($buttonSelectCommonUsers)

# Output section - moved to bottom
$labelOutputUsers = New-Object System.Windows.Forms.Label
$labelOutputUsers.Text = "Output File:"
$labelOutputUsers.Location = New-Object System.Drawing.Point(20, 530)
$labelOutputUsers.AutoSize = $true
$tabUsers.Controls.Add($labelOutputUsers)

$textOutputUsers = New-Object System.Windows.Forms.TextBox
$textOutputUsers.Location = New-Object System.Drawing.Point(100, 530)
$textOutputUsers.Size = New-Object System.Drawing.Size(500, 21)
$textOutputUsers.Text = "$env:USERPROFILE\Desktop\AD_Users_Export.csv"
$tabUsers.Controls.Add($textOutputUsers)

$buttonBrowseUsers = New-Object System.Windows.Forms.Button
$buttonBrowseUsers.Text = "Browse..."
$buttonBrowseUsers.Location = New-Object System.Drawing.Point(610, 529)
$buttonBrowseUsers.Size = New-Object System.Drawing.Size(80, 23)
$tabUsers.Controls.Add($buttonBrowseUsers)

# Export button for Users
$buttonExportUsers = New-Object System.Windows.Forms.Button
$buttonExportUsers.Text = "Export Users"
$buttonExportUsers.Location = New-Object System.Drawing.Point(350, 570)
$buttonExportUsers.Size = New-Object System.Drawing.Size(150, 35)
$buttonExportUsers.BackColor = [System.Drawing.Color]::LightGreen
$tabUsers.Controls.Add($buttonExportUsers)

# ------------------------------
# Groups Tab Content (similar structure)
# ------------------------------

# Domain selection for Groups
$labelDomainGroups = New-Object System.Windows.Forms.Label
$labelDomainGroups.Text = "Domain:"
$labelDomainGroups.Location = New-Object System.Drawing.Point(20, 20)
$labelDomainGroups.AutoSize = $true
$tabGroups.Controls.Add($labelDomainGroups)

$comboDomainGroups = New-Object System.Windows.Forms.ComboBox
$comboDomainGroups.Location = New-Object System.Drawing.Point(80, 20)
$comboDomainGroups.Size = New-Object System.Drawing.Size(300, 21)
$comboDomainGroups.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabGroups.Controls.Add($comboDomainGroups)

# Search Base for Groups
$labelSearchBaseGroups = New-Object System.Windows.Forms.Label
$labelSearchBaseGroups.Text = "Search Base (OU):"
$labelSearchBaseGroups.Location = New-Object System.Drawing.Point(20, 60)
$labelSearchBaseGroups.AutoSize = $true
$tabGroups.Controls.Add($labelSearchBaseGroups)

$comboSearchBaseGroups = New-Object System.Windows.Forms.ComboBox
$comboSearchBaseGroups.Location = New-Object System.Drawing.Point(120, 60)
$comboSearchBaseGroups.Size = New-Object System.Drawing.Size(500, 21)
$comboSearchBaseGroups.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabGroups.Controls.Add($comboSearchBaseGroups)

$buttonRefreshOUGroups = New-Object System.Windows.Forms.Button
$buttonRefreshOUGroups.Text = "↻"
$buttonRefreshOUGroups.Location = New-Object System.Drawing.Point(630, 59)
$buttonRefreshOUGroups.Size = New-Object System.Drawing.Size(30, 23)
$tabGroups.Controls.Add($buttonRefreshOUGroups)

# Filter for Groups
$labelFilterGroups = New-Object System.Windows.Forms.Label
$labelFilterGroups.Text = "Filter:"
$labelFilterGroups.Location = New-Object System.Drawing.Point(20, 100)
$labelFilterGroups.AutoSize = $true
$tabGroups.Controls.Add($labelFilterGroups)

$textFilterGroups = New-Object System.Windows.Forms.TextBox
$textFilterGroups.Location = New-Object System.Drawing.Point(80, 100)
$textFilterGroups.Size = New-Object System.Drawing.Size(300, 21)
$textFilterGroups.Text = "*"
$tabGroups.Controls.Add($textFilterGroups)

# Group type filters
$labelGroupType = New-Object System.Windows.Forms.Label
$labelGroupType.Text = "Group Type:"
$labelGroupType.Location = New-Object System.Drawing.Point(400, 102)
$labelGroupType.AutoSize = $true
$tabGroups.Controls.Add($labelGroupType)

$comboGroupType = New-Object System.Windows.Forms.ComboBox
$comboGroupType.Items.Add("All Groups")
$comboGroupType.Items.Add("Security Groups")
$comboGroupType.Items.Add("Distribution Groups")
$comboGroupType.SelectedIndex = 0
$comboGroupType.Location = New-Object System.Drawing.Point(480, 100)
$comboGroupType.Size = New-Object System.Drawing.Size(150, 21)
$comboGroupType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabGroups.Controls.Add($comboGroupType)

# Properties selection for Groups
$labelPropertiesGroups = New-Object System.Windows.Forms.Label
$labelPropertiesGroups.Text = "Select Properties to Export:"
$labelPropertiesGroups.Location = New-Object System.Drawing.Point(20, 140)
$labelPropertiesGroups.AutoSize = $true
$tabGroups.Controls.Add($labelPropertiesGroups)

$listBoxPropertiesGroups = New-Object System.Windows.Forms.CheckedListBox
$listBoxPropertiesGroups.Location = New-Object System.Drawing.Point(20, 160)
$listBoxPropertiesGroups.Size = New-Object System.Drawing.Size(400, 400)
$listBoxPropertiesGroups.CheckOnClick = $true
$tabGroups.Controls.Add($listBoxPropertiesGroups)

# Group properties
$groupProperties = @(
    "SamAccountName", "Name", "DisplayName", "Description", "GroupCategory", "GroupScope",
    "DistinguishedName", "CanonicalName", "Created", "Modified", "ManagedBy", "Members",
    "MemberOf", "Mail", "Notes", "Info", "HomePage", "whenCreated", "whenChanged"
)

foreach ($prop in $groupProperties) {
    $listBoxPropertiesGroups.Items.Add($prop, $true)
}

# Buttons for Groups property selection
$buttonSelectAllGroups = New-Object System.Windows.Forms.Button
$buttonSelectAllGroups.Text = "Select All"
$buttonSelectAllGroups.Location = New-Object System.Drawing.Point(450, 160)
$buttonSelectAllGroups.Size = New-Object System.Drawing.Size(100, 30)
$tabGroups.Controls.Add($buttonSelectAllGroups)

$buttonSelectNoneGroups = New-Object System.Windows.Forms.Button
$buttonSelectNoneGroups.Text = "Select None"
$buttonSelectNoneGroups.Location = New-Object System.Drawing.Point(450, 200)
$buttonSelectNoneGroups.Size = New-Object System.Drawing.Size(100, 30)
$tabGroups.Controls.Add($buttonSelectNoneGroups)

$buttonSelectCommonGroups = New-Object System.Windows.Forms.Button
$buttonSelectCommonGroups.Text = "Common Only"
$buttonSelectCommonGroups.Location = New-Object System.Drawing.Point(450, 240)
$buttonSelectCommonGroups.Size = New-Object System.Drawing.Size(100, 30)
$tabGroups.Controls.Add($buttonSelectCommonGroups)

# Output path for Groups
$labelOutputGroups = New-Object System.Windows.Forms.Label
$labelOutputGroups.Text = "Output File:"
$labelOutputGroups.Location = New-Object System.Drawing.Point(20, 580)
$labelOutputGroups.AutoSize = $true
$tabGroups.Controls.Add($labelOutputGroups)

$textOutputGroups = New-Object System.Windows.Forms.TextBox
$textOutputGroups.Location = New-Object System.Drawing.Point(100, 580)
$textOutputGroups.Size = New-Object System.Drawing.Size(500, 21)
$textOutputGroups.Text = "$env:USERPROFILE\Desktop\AD_Groups_Export.csv"
$tabGroups.Controls.Add($textOutputGroups)

$buttonBrowseGroups = New-Object System.Windows.Forms.Button
$buttonBrowseGroups.Text = "Browse..."
$buttonBrowseGroups.Location = New-Object System.Drawing.Point(610, 579)
$buttonBrowseGroups.Size = New-Object System.Drawing.Size(80, 23)
$tabGroups.Controls.Add($buttonBrowseGroups)

# Export button for Groups
$buttonExportGroups = New-Object System.Windows.Forms.Button
$buttonExportGroups.Text = "Export Groups"
$buttonExportGroups.Location = New-Object System.Drawing.Point(350, 620)
$buttonExportGroups.Size = New-Object System.Drawing.Size(150, 35)
$buttonExportGroups.BackColor = [System.Drawing.Color]::LightBlue
$tabGroups.Controls.Add($buttonExportGroups)

# ------------------------------
# Computers Tab Content (similar structure)
# ------------------------------

# Domain selection for Computers
$labelDomainComputers = New-Object System.Windows.Forms.Label
$labelDomainComputers.Text = "Domain:"
$labelDomainComputers.Location = New-Object System.Drawing.Point(20, 20)
$labelDomainComputers.AutoSize = $true
$tabComputers.Controls.Add($labelDomainComputers)

$comboDomainComputers = New-Object System.Windows.Forms.ComboBox
$comboDomainComputers.Location = New-Object System.Drawing.Point(80, 20)
$comboDomainComputers.Size = New-Object System.Drawing.Size(300, 21)
$comboDomainComputers.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabComputers.Controls.Add($comboDomainComputers)

# Search Base for Computers
$labelSearchBaseComputers = New-Object System.Windows.Forms.Label
$labelSearchBaseComputers.Text = "Search Base (OU):"
$labelSearchBaseComputers.Location = New-Object System.Drawing.Point(20, 60)
$labelSearchBaseComputers.AutoSize = $true
$tabComputers.Controls.Add($labelSearchBaseComputers)

$comboSearchBaseComputers = New-Object System.Windows.Forms.ComboBox
$comboSearchBaseComputers.Location = New-Object System.Drawing.Point(120, 60)
$comboSearchBaseComputers.Size = New-Object System.Drawing.Size(500, 21)
$comboSearchBaseComputers.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabComputers.Controls.Add($comboSearchBaseComputers)

$buttonRefreshOUComputers = New-Object System.Windows.Forms.Button
$buttonRefreshOUComputers.Text = "↻"
$buttonRefreshOUComputers.Location = New-Object System.Drawing.Point(630, 59)
$buttonRefreshOUComputers.Size = New-Object System.Drawing.Size(30, 23)
$tabComputers.Controls.Add($buttonRefreshOUComputers)

# Filter for Computers
$labelFilterComputers = New-Object System.Windows.Forms.Label
$labelFilterComputers.Text = "Filter:"
$labelFilterComputers.Location = New-Object System.Drawing.Point(20, 100)
$labelFilterComputers.AutoSize = $true
$tabComputers.Controls.Add($labelFilterComputers)

$textFilterComputers = New-Object System.Windows.Forms.TextBox
$textFilterComputers.Location = New-Object System.Drawing.Point(80, 100)
$textFilterComputers.Size = New-Object System.Drawing.Size(300, 21)
$textFilterComputers.Text = "*"
$tabComputers.Controls.Add($textFilterComputers)

# Quick filter buttons for Computers
$buttonEnabledComputers = New-Object System.Windows.Forms.Button
$buttonEnabledComputers.Text = "Enabled"
$buttonEnabledComputers.Location = New-Object System.Drawing.Point(400, 99)
$buttonEnabledComputers.Size = New-Object System.Drawing.Size(80, 23)
$tabComputers.Controls.Add($buttonEnabledComputers)

$buttonDisabledComputers = New-Object System.Windows.Forms.Button
$buttonDisabledComputers.Text = "Disabled"
$buttonDisabledComputers.Location = New-Object System.Drawing.Point(490, 99)
$buttonDisabledComputers.Size = New-Object System.Drawing.Size(80, 23)
$tabComputers.Controls.Add($buttonDisabledComputers)

# Properties selection for Computers
$labelPropertiesComputers = New-Object System.Windows.Forms.Label
$labelPropertiesComputers.Text = "Select Properties to Export:"
$labelPropertiesComputers.Location = New-Object System.Drawing.Point(20, 140)
$labelPropertiesComputers.AutoSize = $true
$tabComputers.Controls.Add($labelPropertiesComputers)

$listBoxPropertiesComputers = New-Object System.Windows.Forms.CheckedListBox
$listBoxPropertiesComputers.Location = New-Object System.Drawing.Point(20, 160)
$listBoxPropertiesComputers.Size = New-Object System.Drawing.Size(400, 400)
$listBoxPropertiesComputers.CheckOnClick = $true
$tabComputers.Controls.Add($listBoxPropertiesComputers)

# Computer properties
$computerProperties = @(
    "Name", "SamAccountName", "DNSHostName", "Description", "OperatingSystem",
    "OperatingSystemVersion", "OperatingSystemServicePack", "Enabled", "Created", "Modified",
    "LastLogonDate", "PasswordLastSet", "DistinguishedName", "CanonicalName", "Location",
    "ManagedBy", "ServicePrincipalNames", "TrustedForDelegation", "IPv4Address", "IPv6Address"
)

foreach ($prop in $computerProperties) {
    $listBoxPropertiesComputers.Items.Add($prop, $true)
}

# Buttons for Computers property selection
$buttonSelectAllComputers = New-Object System.Windows.Forms.Button
$buttonSelectAllComputers.Text = "Select All"
$buttonSelectAllComputers.Location = New-Object System.Drawing.Point(450, 160)
$buttonSelectAllComputers.Size = New-Object System.Drawing.Size(100, 30)
$tabComputers.Controls.Add($buttonSelectAllComputers)

$buttonSelectNoneComputers = New-Object System.Windows.Forms.Button
$buttonSelectNoneComputers.Text = "Select None"
$buttonSelectNoneComputers.Location = New-Object System.Drawing.Point(450, 200)
$buttonSelectNoneComputers.Size = New-Object System.Drawing.Size(100, 30)
$tabComputers.Controls.Add($buttonSelectNoneComputers)

$buttonSelectCommonComputers = New-Object System.Windows.Forms.Button
$buttonSelectCommonComputers.Text = "Common Only"
$buttonSelectCommonComputers.Location = New-Object System.Drawing.Point(450, 240)
$buttonSelectCommonComputers.Size = New-Object System.Drawing.Size(100, 30)
$tabComputers.Controls.Add($buttonSelectCommonComputers)

# Output path for Computers
$labelOutputComputers = New-Object System.Windows.Forms.Label
$labelOutputComputers.Text = "Output File:"
$labelOutputComputers.Location = New-Object System.Drawing.Point(20, 580)
$labelOutputComputers.AutoSize = $true
$tabComputers.Controls.Add($labelOutputComputers)

$textOutputComputers = New-Object System.Windows.Forms.TextBox
$textOutputComputers.Location = New-Object System.Drawing.Point(100, 580)
$textOutputComputers.Size = New-Object System.Drawing.Size(500, 21)
$textOutputComputers.Text = "$env:USERPROFILE\Desktop\AD_Computers_Export.csv"
$tabComputers.Controls.Add($textOutputComputers)

$buttonBrowseComputers = New-Object System.Windows.Forms.Button
$buttonBrowseComputers.Text = "Browse..."
$buttonBrowseComputers.Location = New-Object System.Drawing.Point(610, 579)
$buttonBrowseComputers.Size = New-Object System.Drawing.Size(80, 23)
$tabComputers.Controls.Add($buttonBrowseComputers)

# Export button for Computers
$buttonExportComputers = New-Object System.Windows.Forms.Button
$buttonExportComputers.Text = "Export Computers"
$buttonExportComputers.Location = New-Object System.Drawing.Point(350, 620)
$buttonExportComputers.Size = New-Object System.Drawing.Size(150, 35)
$buttonExportComputers.BackColor = [System.Drawing.Color]::LightCoral
$tabComputers.Controls.Add($buttonExportComputers)

# ------------------------------
# Group Memberships Tab Content
# ------------------------------

# Domain selection for Memberships
$labelDomainMemberships = New-Object System.Windows.Forms.Label
$labelDomainMemberships.Text = "Domain:"
$labelDomainMemberships.Location = New-Object System.Drawing.Point(20, 20)
$labelDomainMemberships.AutoSize = $true
$tabMemberships.Controls.Add($labelDomainMemberships)

$comboDomainMemberships = New-Object System.Windows.Forms.ComboBox
$comboDomainMemberships.Location = New-Object System.Drawing.Point(80, 20)
$comboDomainMemberships.Size = New-Object System.Drawing.Size(300, 21)
$comboDomainMemberships.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabMemberships.Controls.Add($comboDomainMemberships)

# Group name filter
$labelGroupFilter = New-Object System.Windows.Forms.Label
$labelGroupFilter.Text = "Group Name Filter:"
$labelGroupFilter.Location = New-Object System.Drawing.Point(20, 60)
$labelGroupFilter.AutoSize = $true
$tabMemberships.Controls.Add($labelGroupFilter)

$textGroupFilter = New-Object System.Windows.Forms.TextBox
$textGroupFilter.Location = New-Object System.Drawing.Point(140, 60)
$textGroupFilter.Size = New-Object System.Drawing.Size(300, 21)
$textGroupFilter.Text = "*"
$tabMemberships.Controls.Add($textGroupFilter)

# Output path for Memberships
$labelOutputMemberships = New-Object System.Windows.Forms.Label
$labelOutputMemberships.Text = "Output File:"
$labelOutputMemberships.Location = New-Object System.Drawing.Point(20, 100)
$labelOutputMemberships.AutoSize = $true
$tabMemberships.Controls.Add($labelOutputMemberships)

$textOutputMemberships = New-Object System.Windows.Forms.TextBox
$textOutputMemberships.Location = New-Object System.Drawing.Point(100, 100)
$textOutputMemberships.Size = New-Object System.Drawing.Size(500, 21)
$textOutputMemberships.Text = "$env:USERPROFILE\Desktop\AD_GroupMemberships_Export.csv"
$tabMemberships.Controls.Add($textOutputMemberships)

$buttonBrowseMemberships = New-Object System.Windows.Forms.Button
$buttonBrowseMemberships.Text = "Browse..."
$buttonBrowseMemberships.Location = New-Object System.Drawing.Point(610, 99)
$buttonBrowseMemberships.Size = New-Object System.Drawing.Size(80, 23)
$tabMemberships.Controls.Add($buttonBrowseMemberships)

# Export button for Memberships
$buttonExportMemberships = New-Object System.Windows.Forms.Button
$buttonExportMemberships.Text = "Export Group Memberships"
$buttonExportMemberships.Location = New-Object System.Drawing.Point(300, 150)
$buttonExportMemberships.Size = New-Object System.Drawing.Size(200, 35)
$buttonExportMemberships.BackColor = [System.Drawing.Color]::LightYellow
$tabMemberships.Controls.Add($buttonExportMemberships)

# Information label for memberships
$labelMembershipInfo = New-Object System.Windows.Forms.Label
$labelMembershipInfo.Text = "This will export all members of groups matching the filter, showing group name, member name, and member type."
$labelMembershipInfo.Location = New-Object System.Drawing.Point(20, 200)
$labelMembershipInfo.Size = New-Object System.Drawing.Size(800, 40)
$labelMembershipInfo.AutoSize = $false
$tabMemberships.Controls.Add($labelMembershipInfo)

# Add the tab control to the form
$form.Controls.Add($tabControl)

# Initialize domain dropdowns
$domains = Get-ADDomainList
foreach ($domain in $domains) {
    $comboDomainUsers.Items.Add($domain)
    $comboDomainGroups.Items.Add($domain)
    $comboDomainComputers.Items.Add($domain)
    $comboDomainMemberships.Items.Add($domain)
}

if ($comboDomainUsers.Items.Count -gt 0) {
    $comboDomainUsers.SelectedIndex = 0
    $comboDomainGroups.SelectedIndex = 0
    $comboDomainComputers.SelectedIndex = 0
    $comboDomainMemberships.SelectedIndex = 0
}

# Function to populate OU dropdown
function Update-OUDropdown {
    param([System.Windows.Forms.ComboBox]$ComboBox, [string]$Domain)
    
    $ComboBox.Items.Clear()
    $ComboBox.Items.Add("(Entire Domain)")
    
    if ($Domain) {
        $ous = Get-ADOUList -Domain $Domain
        foreach ($ou in $ous) {
            $ComboBox.Items.Add($ou.DistinguishedName)
        }
    }
    
    $ComboBox.SelectedIndex = 0
}

# Function to populate group dropdown
function Update-GroupDropdown {
    param(
        [System.Windows.Forms.ComboBox]$ComboBox, 
        [string]$Domain,
        [string]$GroupType = "All"
    )
    
    $ComboBox.Items.Clear()
    $ComboBox.Items.Add("(No Group Filter)")
    
    if ($Domain) {
        $groups = Get-ADGroupsByType -Domain $Domain -GroupType $GroupType
        foreach ($group in $groups) {
            $displayText = "$($group.Name) ($($group.GroupCategory)/$($group.GroupScope))"
            $ComboBox.Items.Add($group.Name)
        }
    }
    
    $ComboBox.SelectedIndex = 0
}

# Event handlers

# Domain selection change handlers
$comboDomainUsers.Add_SelectedIndexChanged({
    Update-OUDropdown -ComboBox $comboSearchBaseUsers -Domain $comboDomainUsers.SelectedItem
    Update-GroupDropdown -ComboBox $comboGroupFilterUsers -Domain $comboDomainUsers.SelectedItem -GroupType $comboGroupTypeUsers.SelectedItem
})

$comboDomainGroups.Add_SelectedIndexChanged({
    Update-OUDropdown -ComboBox $comboSearchBaseGroups -Domain $comboDomainGroups.SelectedItem
})

$comboDomainComputers.Add_SelectedIndexChanged({
    Update-OUDropdown -ComboBox $comboSearchBaseComputers -Domain $comboDomainComputers.SelectedItem
})

# OU refresh button handlers
$buttonRefreshOUUsers.Add_Click({
    Update-OUDropdown -ComboBox $comboSearchBaseUsers -Domain $comboDomainUsers.SelectedItem
})

$buttonRefreshOUGroups.Add_Click({
    Update-OUDropdown -ComboBox $comboSearchBaseGroups -Domain $comboDomainGroups.SelectedItem
})

$buttonRefreshOUComputers.Add_Click({
    Update-OUDropdown -ComboBox $comboSearchBaseComputers -Domain $comboDomainComputers.SelectedItem
})

# Quick filter handlers for Users
$buttonEnabledUsers.Add_Click({
    $textFilterUsers.Text = "Enabled -eq `$true"
})

$buttonDisabledUsers.Add_Click({
    $textFilterUsers.Text = "Enabled -eq `$false"
})

# Quick filter handlers for Computers
$buttonEnabledComputers.Add_Click({
    $textFilterComputers.Text = "Enabled -eq `$true"
})

$buttonDisabledComputers.Add_Click({
    $textFilterComputers.Text = "Enabled -eq `$false"
})

# Group type selection change handler for Users
$comboGroupTypeUsers.Add_SelectedIndexChanged({
    $groupType = if ($comboGroupTypeUsers.SelectedIndex -eq 0) { "All" } else { $comboGroupTypeUsers.SelectedItem }
    Update-GroupDropdown -ComboBox $comboGroupFilterUsers -Domain $comboDomainUsers.SelectedItem -GroupType $groupType
})

# Group filter type change handler
$comboGroupFilterTypeUsers.Add_SelectedIndexChanged({
    $enabled = $comboGroupFilterTypeUsers.SelectedIndex -ne 0
    $comboGroupTypeUsers.Enabled = $enabled
    $comboGroupFilterUsers.Enabled = $enabled
    
    if (-not $enabled) {
        $comboGroupTypeUsers.SelectedIndex = 0
        $comboGroupFilterUsers.SelectedIndex = 0
    } else {
        # Update group dropdown when enabling group filtering
        $groupType = if ($comboGroupTypeUsers.SelectedIndex -eq 0) { "All" } else { $comboGroupTypeUsers.SelectedItem }
        Update-GroupDropdown -ComboBox $comboGroupFilterUsers -Domain $comboDomainUsers.SelectedItem -GroupType $groupType
    }
})

# Property selection handlers for Users
$buttonSelectAllUsers.Add_Click({
    for ($i = 0; $i -lt $listBoxPropertiesUsers.Items.Count; $i++) {
        $listBoxPropertiesUsers.SetItemChecked($i, $true)
    }
})

$buttonSelectNoneUsers.Add_Click({
    for ($i = 0; $i -lt $listBoxPropertiesUsers.Items.Count; $i++) {
        $listBoxPropertiesUsers.SetItemChecked($i, $false)
    }
})

$buttonSelectCommonUsers.Add_Click({
    for ($i = 0; $i -lt $listBoxPropertiesUsers.Items.Count; $i++) {
        $listBoxPropertiesUsers.SetItemChecked($i, $false)
    }
    
    $commonProps = @("SamAccountName", "Name", "DisplayName", "UserPrincipalName", "EmailAddress", "Department", "Title", "Enabled", "LastLogonDate")
    for ($i = 0; $i -lt $listBoxPropertiesUsers.Items.Count; $i++) {
        if ($commonProps -contains $listBoxPropertiesUsers.Items[$i]) {
            $listBoxPropertiesUsers.SetItemChecked($i, $true)
        }
    }
})

# Property selection handlers for Groups
$buttonSelectAllGroups.Add_Click({
    for ($i = 0; $i -lt $listBoxPropertiesGroups.Items.Count; $i++) {
        $listBoxPropertiesGroups.SetItemChecked($i, $true)
    }
})

$buttonSelectNoneGroups.Add_Click({
    for ($i = 0; $i -lt $listBoxPropertiesGroups.Items.Count; $i++) {
        $listBoxPropertiesGroups.SetItemChecked($i, $false)
    }
})

$buttonSelectCommonGroups.Add_Click({
    for ($i = 0; $i -lt $listBoxPropertiesGroups.Items.Count; $i++) {
        $listBoxPropertiesGroups.SetItemChecked($i, $false)
    }
    
    $commonProps = @("SamAccountName", "Name", "Description", "GroupCategory", "GroupScope", "ManagedBy")
    for ($i = 0; $i -lt $listBoxPropertiesGroups.Items.Count; $i++) {
        if ($commonProps -contains $listBoxPropertiesGroups.Items[$i]) {
            $listBoxPropertiesGroups.SetItemChecked($i, $true)
        }
    }
})

# Property selection handlers for Computers
$buttonSelectAllComputers.Add_Click({
    for ($i = 0; $i -lt $listBoxPropertiesComputers.Items.Count; $i++) {
        $listBoxPropertiesComputers.SetItemChecked($i, $true)
    }
})

$buttonSelectNoneComputers.Add_Click({
    for ($i = 0; $i -lt $listBoxPropertiesComputers.Items.Count; $i++) {
        $listBoxPropertiesComputers.SetItemChecked($i, $false)
    }
})

$buttonSelectCommonComputers.Add_Click({
    for ($i = 0; $i -lt $listBoxPropertiesComputers.Items.Count; $i++) {
        $listBoxPropertiesComputers.SetItemChecked($i, $false)
    }
    
    $commonProps = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "LastLogonDate", "Description")
    for ($i = 0; $i -lt $listBoxPropertiesComputers.Items.Count; $i++) {
        if ($commonProps -contains $listBoxPropertiesComputers.Items[$i]) {
            $listBoxPropertiesComputers.SetItemChecked($i, $true)
        }
    }
})

# Browse button handlers
$buttonBrowseUsers.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save Users Export As"
    $saveFileDialog.FileName = "AD_Users_Export.csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textOutputUsers.Text = $saveFileDialog.FileName
    }
})

$buttonBrowseGroups.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save Groups Export As"
    $saveFileDialog.FileName = "AD_Groups_Export.csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textOutputGroups.Text = $saveFileDialog.FileName
    }
})

$buttonBrowseComputers.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save Computers Export As"
    $saveFileDialog.FileName = "AD_Computers_Export.csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textOutputComputers.Text = $saveFileDialog.FileName
    }
})

$buttonBrowseMemberships.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save Group Memberships Export As"
    $saveFileDialog.FileName = "AD_GroupMemberships_Export.csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textOutputMemberships.Text = $saveFileDialog.FileName
    }
})

# Export button handlers
$buttonExportUsers.Add_Click({
    $selectedProperties = @()
    for ($i = 0; $i -lt $listBoxPropertiesUsers.Items.Count; $i++) {
        if ($listBoxPropertiesUsers.GetItemChecked($i)) {
            $selectedProperties += $listBoxPropertiesUsers.Items[$i]
        }
    }
    
    if ($selectedProperties.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one property to export.", "No Properties Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $domain = $comboDomainUsers.SelectedItem
    $searchBase = if ($comboSearchBaseUsers.SelectedIndex -eq 0) { "" } else { $comboSearchBaseUsers.SelectedItem }
    $filter = $textFilterUsers.Text.Trim()
    $outputPath = $textOutputUsers.Text.Trim()
    
    # Group filtering parameters
    $groupFilter = ""
    $groupFilterType = ""
    
    if ($comboGroupFilterTypeUsers.SelectedIndex -eq 1) {
        # Members Of
        $groupFilterType = "MemberOf"
        if ($comboGroupFilterUsers.SelectedIndex -gt 0) {
            $groupFilter = $comboGroupFilterUsers.SelectedItem
        }
    }
    elseif ($comboGroupFilterTypeUsers.SelectedIndex -eq 2) {
        # Not Members Of
        $groupFilterType = "NotMemberOf"
        if ($comboGroupFilterUsers.SelectedIndex -gt 0) {
            $groupFilter = $comboGroupFilterUsers.SelectedItem
        }
    }
    
    if (-not $outputPath) {
        [System.Windows.Forms.MessageBox]::Show("Please specify an output file path.", "Missing Output Path", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $statusLabel.Text = "Exporting users... Please wait."
    $form.Refresh()
    
    Write-DebugLog "Export parameters - Properties: $($selectedProperties -join ', ')"
    Write-DebugLog "Export parameters - Group Filter: $groupFilter, Type: $groupFilterType"
    
    $result = Export-ADUsersToCSV -OutputPath $outputPath -Properties $selectedProperties -Filter $filter -SearchBase $searchBase -Domain $domain -GroupFilter $groupFilter -GroupFilterType $groupFilterType
    
    if ($result.Success) {
        $groupFilterMessage = if ($groupFilter) { " (filtered by group: $groupFilter)" } else { "" }
        [System.Windows.Forms.MessageBox]::Show($result.Message + $groupFilterMessage, "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $statusLabel.Text = "Export completed successfully. $($result.Count) users exported$groupFilterMessage."
    } else {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Export Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Export failed: $($result.Message)"
    }
})

$buttonExportGroups.Add_Click({
    $selectedProperties = @()
    for ($i = 0; $i -lt $listBoxPropertiesGroups.Items.Count; $i++) {
        if ($listBoxPropertiesGroups.GetItemChecked($i)) {
            $selectedProperties += $listBoxPropertiesGroups.Items[$i]
        }
    }
    
    if ($selectedProperties.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one property to export.", "No Properties Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $domain = $comboDomainGroups.SelectedItem
    $searchBase = if ($comboSearchBaseGroups.SelectedIndex -eq 0) { "" } else { $comboSearchBaseGroups.SelectedItem }
    $filter = $textFilterGroups.Text.Trim()
    $outputPath = $textOutputGroups.Text.Trim()
    
    # Adjust filter based on group type selection
    if ($comboGroupType.SelectedIndex -eq 1) {
        $filter = if ($filter -eq "*") { "GroupCategory -eq 'Security'" } else { "($filter) -and (GroupCategory -eq 'Security')" }
    } elseif ($comboGroupType.SelectedIndex -eq 2) {
        $filter = if ($filter -eq "*") { "GroupCategory -eq 'Distribution'" } else { "($filter) -and (GroupCategory -eq 'Distribution')" }
    }
    
    if (-not $outputPath) {
        [System.Windows.Forms.MessageBox]::Show("Please specify an output file path.", "Missing Output Path", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $statusLabel.Text = "Exporting groups... Please wait."
    $form.Refresh()
    
    $result = Export-ADGroupsToCSV -OutputPath $outputPath -Properties $selectedProperties -Filter $filter -SearchBase $searchBase -Domain $domain
    
    if ($result.Success) {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $statusLabel.Text = "Export completed successfully. $($result.Count) groups exported."
    } else {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Export Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Export failed: $($result.Message)"
    }
})

$buttonExportComputers.Add_Click({
    $selectedProperties = @()
    for ($i = 0; $i -lt $listBoxPropertiesComputers.Items.Count; $i++) {
        if ($listBoxPropertiesComputers.GetItemChecked($i)) {
            $selectedProperties += $listBoxPropertiesComputers.Items[$i]
        }
    }
    
    if ($selectedProperties.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one property to export.", "No Properties Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $domain = $comboDomainComputers.SelectedItem
    $searchBase = if ($comboSearchBaseComputers.SelectedIndex -eq 0) { "" } else { $comboSearchBaseComputers.SelectedItem }
    $filter = $textFilterComputers.Text.Trim()
    $outputPath = $textOutputComputers.Text.Trim()
    
    if (-not $outputPath) {
        [System.Windows.Forms.MessageBox]::Show("Please specify an output file path.", "Missing Output Path", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $statusLabel.Text = "Exporting computers... Please wait."
    $form.Refresh()
    
    $result = Export-ADComputersToCSV -OutputPath $outputPath -Properties $selectedProperties -Filter $filter -SearchBase $searchBase -Domain $domain
    
    if ($result.Success) {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $statusLabel.Text = "Export completed successfully. $($result.Count) computers exported."
    } else {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Export Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Export failed: $($result.Message)"
    }
})

$buttonExportMemberships.Add_Click({
    $domain = $comboDomainMemberships.SelectedItem
    $groupFilter = $textGroupFilter.Text.Trim()
    $outputPath = $textOutputMemberships.Text.Trim()
    
    if (-not $outputPath) {
        [System.Windows.Forms.MessageBox]::Show("Please specify an output file path.", "Missing Output Path", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $statusLabel.Text = "Exporting group memberships... Please wait."
    $form.Refresh()
    
    $result = Export-ADGroupMembershipsToCSV -OutputPath $outputPath -GroupName $groupFilter -Domain $domain
    
    if ($result.Success) {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $statusLabel.Text = "Export completed successfully. $($result.Count) memberships exported."
    } else {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Export Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Export failed: $($result.Message)"
    }
})

# Initialize the form
$form.Add_Load({
    $statusLabel.Text = "Active Directory Export Utility loaded successfully."
    
    # Initialize OU dropdowns for the default domain
    if ($comboDomainUsers.Items.Count -gt 0) {
        Update-OUDropdown -ComboBox $comboSearchBaseUsers -Domain $comboDomainUsers.SelectedItem
        Update-OUDropdown -ComboBox $comboSearchBaseGroups -Domain $comboDomainGroups.SelectedItem
        Update-OUDropdown -ComboBox $comboSearchBaseComputers -Domain $comboDomainComputers.SelectedItem
        
        # Initialize group dropdown
        Update-GroupDropdown -ComboBox $comboGroupFilterUsers -Domain $comboDomainUsers.SelectedItem
        
        # Initialize group filter controls state
        $comboGroupTypeUsers.Enabled = $false
        $comboGroupFilterUsers.Enabled = $false
    }
})

# Ensure console is visible if debugging
if (-not $debug) {
    [Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)
}

# Show the form
$form.ShowDialog() | Out-Null
