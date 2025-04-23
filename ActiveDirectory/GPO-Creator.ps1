#Requires -RunAsAdministrator
#Requires -Modules GroupPolicy

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

# Import required modules
if (-not (Get-Module -Name GroupPolicy -ErrorAction SilentlyContinue)) {
    try {
        Import-Module GroupPolicy -ErrorAction Stop
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to import GroupPolicy module. Make sure RSAT tools are installed.", "Module Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}

# Function to create a new GPO
function New-CustomGPO {
    param (
        [string]$GPOName,
        [string]$Comment,
        [System.Collections.Hashtable[]]$RegistrySettings,
        [switch]$LinkToOU,
        [string]$OUPath
    )
    
    try {
        # Check if GPO already exists
        $existingGPO = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue
        
        if ($existingGPO) {
            $result = [System.Windows.Forms.MessageBox]::Show("A GPO with the name '$GPOName' already exists. Do you want to replace it?", "GPO Exists", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Remove-GPO -Name $GPOName -Confirm:$false
            }
            else {
                return @{
                    Success = $false
                    Message = "Operation cancelled by user."
                }
            }
        }
        
        # Create new GPO
        $gpo = New-GPO -Name $GPOName -Comment $Comment
        
        # Process registry settings
        foreach ($setting in $RegistrySettings) {
            $keyPath = $setting.KeyPath
            $valueName = $setting.ValueName
            $valueData = $setting.ValueData
            $valueType = $setting.ValueType
            
            # Set registry value
            Set-GPRegistryValue -Guid $gpo.Id -Key $keyPath -ValueName $valueName -Value $valueData -Type $valueType -ErrorAction Stop
        }
        
        # Link GPO to OU if specified
        if ($LinkToOU -and $OUPath) {
            New-GPLink -Name $GPOName -Target $OUPath -LinkEnabled Yes
        }
        
        return @{
            Success = $true
            Message = "GPO '$GPOName' created successfully."
            GPOId = $gpo.Id
            GPODisplayName = $gpo.DisplayName
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Failed to create GPO: $($_.Exception.Message)"
        }
    }
}

# Function to generate security descriptors for event logs
function Get-EventLogSecurityDescriptor {
    param (
        [string]$PermissionLevel,
        [string]$Principal = "Everyone"
    )
    
    # Define security descriptor components
    $ownerSID = "O:BA"           # Built-in Administrators
    $groupSID = "G:SY"           # Local System
    $dacl = "D:"
    
    # Common permissions
    $systemFullAccess = "(A;;0xf0007;;;SY)"     # System Full Access
    $adminFullAccess = "(A;;0xf0007;;;BA)"      # Administrators Full Access
    $backupOps = "(A;;0x7;;;SO)"                # Backup Operators Read/Write
    
    # Base DACL always includes system and admin permissions
    $dacl += $systemFullAccess + $adminFullAccess + $backupOps
    
    # Convert principal to SID if it's a well-known group
    $principalSID = switch ($Principal) {
        "Everyone" { "S-1-1-0" }
        "Authenticated Users" { "S-1-5-11" }
        "Interactive Users" { "S-1-5-4" }
        "Network" { "S-1-5-2" }
        "Service" { "S-1-5-6" }
        "Batch" { "S-1-5-3" }
        default { $Principal }  # Assume it's already a SID
    }
    
    # Add the appropriate permission based on level
    switch ($PermissionLevel) {
        "FullAccess" { 
            $dacl += "(A;;0xf0007;;;$principalSID)"  # Full control
        }
        "ReadWrite" { 
            $dacl += "(A;;0x7;;;$principalSID)"      # Read/Write
        }
        "Read" { 
            $dacl += "(A;;0x3;;;$principalSID)"      # Read only
        }
    }
    
    # Construct the security descriptor
    $securityDescriptor = $ownerSID + $groupSID + $dacl
    
    return $securityDescriptor
}

# Function to create Event Log Security GPO with option to create the event source
function New-EventLogSecurityGPO {
    param (
        [string]$GPOName,
        [string]$LogName,
        [string]$SecurityDescriptor,
        [bool]$CreateEventSource = $false
    )
    
    $registrySettings = @(
        @{
            KeyPath = "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\$LogName"
            ValueName = "CustomSD"
            ValueData = $SecurityDescriptor
            ValueType = "String"
        }
    )
    
    # If creating a custom event source, add required registry settings
    if ($CreateEventSource) {
        # Add settings to create the event log if it doesn't exist
        $registrySettings += @{
            KeyPath = "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\$LogName"
            ValueName = "File"
            ValueData = "%SystemRoot%\System32\Winevt\Logs\$LogName.evtx"
            ValueType = "ExpandString"
        }
        
        $registrySettings += @{
            KeyPath = "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\$LogName"
            ValueName = "MaxSize"
            ValueData = 1048576  # 1MB default size
            ValueType = "DWord"
        }
        
        $registrySettings += @{
            KeyPath = "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\$LogName"
            ValueName = "Retention"
            ValueData = 0        # 0 = overwrite as needed
            ValueType = "DWord"
        }
    }
    
    $gpoComment = "Configures security permissions and settings for the $LogName event log"
    
    return New-CustomGPO -GPOName $GPOName -Comment $gpoComment -RegistrySettings $registrySettings
}

# Function to create Software Installation GPO
function New-SoftwareInstallationGPO {
    param (
        [string]$GPOName,
        [string]$PackagePath,
        [string]$DeploymentType # "Published" or "Assigned"
    )
    
    try {
        # Check if GPO already exists
        $existingGPO = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue
        
        if ($existingGPO) {
            $result = [System.Windows.Forms.MessageBox]::Show("A GPO with the name '$GPOName' already exists. Do you want to replace it?", "GPO Exists", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Remove-GPO -Name $GPOName -Confirm:$false
            }
            else {
                return @{
                    Success = $false
                    Message = "Operation cancelled by user."
                }
            }
        }
        
        # Create new GPO
        $gpo = New-GPO -Name $GPOName -Comment "Software Installation GPO for $(Split-Path $PackagePath -Leaf)"
        
        # Set software installation
        $deploymentTypeValue = if ($DeploymentType -eq "Published") { 0 } else { 1 } # 0=Published, 1=Assigned
        
        $gpoSession = Open-GPEditSession -GPO $gpo
        Import-GPExtension -Context 'User' -Extension 'Software Installation' -Session $gpoSession
        
        $softwareInstallationSettings = @{
            GPODisplayName = $gpo.DisplayName
            PackageLocation = $PackagePath
            EnableUserConfiguration = $true
            Action = $deploymentType
        }
        
        $null = New-GPSoftwareInstallationSettings @softwareInstallationSettings
        
        Save-GPEditSession $gpoSession
        
        return @{
            Success = $true
            Message = "Software Installation GPO '$GPOName' created successfully."
            GPOId = $gpo.Id
            GPODisplayName = $gpo.DisplayName
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Failed to create Software Installation GPO: $($_.Exception.Message)"
        }
    }
}

# Function to create Firewall Rule GPO
function New-FirewallRuleGPO {
    param (
        [string]$GPOName,
        [string]$DisplayName,
        [string]$Direction, # "Inbound" or "Outbound"
        [string]$Action,    # "Allow" or "Block"
        [int[]]$LocalPorts,
        [string]$Protocol   # "TCP" or "UDP"
    )
    
    $directionValue = if ($Direction -eq "Inbound") { "In" } else { "Out" }
    $actionValue = if ($Action -eq "Allow") { "Allow" } else { "Block" }
    
    try {
        # Create new GPO
        $gpo = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue
        
        if (-not $gpo) {
            $gpo = New-GPO -Name $GPOName -Comment "Firewall rule GPO for $DisplayName"
        }
        
        # Configure the firewall rule using netsh commands
        $ports = $LocalPorts -join ","
        $command = "netsh advfirewall firewall add rule name=`"$DisplayName`" dir=$directionValue action=$actionValue protocol=$Protocol localport=$ports"
        
        # Add the command to the GPO's startup script
        $scriptContent = @"
$command
"@
        
        $scriptPath = "$env:TEMP\FirewallRule_$((New-Guid).Guid).ps1"
        $scriptContent | Out-File -FilePath $scriptPath -Force
        
        # Add the script to the GPO
        $gpoScript = @{
            Name = "ConfigureFirewallRule"
            ScriptType = "Startup"
            GPOName = $GPOName
            ScriptPath = $scriptPath
            Parameters = ""
        }
        
        Set-GPPrefRegistryValue -Name $GPOName -Context Computer -Key "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Startup\0" -ValueName "DisplayName" -Value $DisplayName -Type String
        Set-GPPrefRegistryValue -Name $GPOName -Context Computer -Key "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Startup\0" -ValueName "FileSysPath" -Value "%TEMP%" -Type String
        Set-GPPrefRegistryValue -Name $GPOName -Context Computer -Key "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Startup\0" -ValueName "Parameters" -Value "" -Type String
        Set-GPPrefRegistryValue -Name $GPOName -Context Computer -Key "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Startup\0" -ValueName "Script" -Value (Split-Path $scriptPath -Leaf) -Type String
        
        return @{
            Success = $true
            Message = "Firewall rule GPO '$GPOName' created successfully."
            GPOId = $gpo.Id
            GPODisplayName = $gpo.DisplayName
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Failed to create Firewall Rule GPO: $($_.Exception.Message)"
        }
    }
}

# Function to get existing GPO names
function Get-ExistingGPONames {
    try {
        $gpos = Get-GPO -All -ErrorAction Stop
        return $gpos | Select-Object -ExpandProperty DisplayName | Sort-Object
    }
    catch {
        return @()
    }
}

# Function to populate a ComboBox with existing GPO names
function Update-GPOComboBox {
    param (
        [System.Windows.Forms.ComboBox]$ComboBox
    )
    
    $ComboBox.Items.Clear()
    $ComboBox.Items.Add("<Create New GPO>")
    
    $gpoNames = Get-ExistingGPONames
    foreach ($name in $gpoNames) {
        $ComboBox.Items.Add($name)
    }
    
    $ComboBox.SelectedIndex = 0
}

# Function to import GPO data from file
function Import-GPOData {
    param (
        [string]$FilePath
    )
    
    try {
        if (Test-Path -Path $FilePath) {
            $data = Get-Content -Path $FilePath -Raw | ConvertFrom-Json
            return $data
        }
        else {
            return $null
        }
    }
    catch {
        return $null
    }
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "GPO Creator Utility"
$form.Size = New-Object System.Drawing.Size(800, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# Create a tab control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(765, 600)

# Create tab pages
$tabEventLog = New-Object System.Windows.Forms.TabPage
$tabEventLog.Text = "Event Log Security"
$tabSoftware = New-Object System.Windows.Forms.TabPage
$tabSoftware.Text = "Software Installation"
$tabFirewall = New-Object System.Windows.Forms.TabPage
$tabFirewall.Text = "Firewall Rules"
$tabRegistry = New-Object System.Windows.Forms.TabPage
$tabRegistry.Text = "Registry Settings"

$tabControl.Controls.Add($tabEventLog)
$tabControl.Controls.Add($tabSoftware)
$tabControl.Controls.Add($tabFirewall)
$tabControl.Controls.Add($tabRegistry)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 620)
$statusLabel.Size = New-Object System.Drawing.Size(765, 40)
$statusLabel.Text = "Ready to create GPOs..."
$statusLabel.AutoSize = $false
$form.Controls.Add($statusLabel)

# ------------------------------
# Event Log Security Tab Content
# ------------------------------
$labelGPONameEventLog = New-Object System.Windows.Forms.Label
$labelGPONameEventLog.Text = "GPO Name:"
$labelGPONameEventLog.Location = New-Object System.Drawing.Point(20, 20)
$labelGPONameEventLog.AutoSize = $true
$tabEventLog.Controls.Add($labelGPONameEventLog)

$comboGPONameEventLog = New-Object System.Windows.Forms.ComboBox
$comboGPONameEventLog.Location = New-Object System.Drawing.Point(120, 20)
$comboGPONameEventLog.Size = New-Object System.Drawing.Size(300, 20)
$comboGPONameEventLog.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabEventLog.Controls.Add($comboGPONameEventLog)

$textGPONameEventLog = New-Object System.Windows.Forms.TextBox
$textGPONameEventLog.Location = New-Object System.Drawing.Point(430, 20)
$textGPONameEventLog.Size = New-Object System.Drawing.Size(250, 20)
$textGPONameEventLog.Enabled = $true
$tabEventLog.Controls.Add($textGPONameEventLog)

$buttonRefreshEventLogGPOs = New-Object System.Windows.Forms.Button
$buttonRefreshEventLogGPOs.Text = "↻"
$buttonRefreshEventLogGPOs.Location = New-Object System.Drawing.Point(690, 19)
$buttonRefreshEventLogGPOs.Size = New-Object System.Drawing.Size(30, 23)
$tabEventLog.Controls.Add($buttonRefreshEventLogGPOs)

$labelEventLog = New-Object System.Windows.Forms.Label
$labelEventLog.Text = "Event Log:"
$labelEventLog.Location = New-Object System.Drawing.Point(20, 60)
$labelEventLog.AutoSize = $true
$tabEventLog.Controls.Add($labelEventLog)

$comboEventLog = New-Object System.Windows.Forms.ComboBox
$comboEventLog.Location = New-Object System.Drawing.Point(120, 60)
$comboEventLog.Size = New-Object System.Drawing.Size(300, 20)
$comboEventLog.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabEventLog.Controls.Add($comboEventLog)

# Custom Event Log Section
$buttonAddCustomEventLog = New-Object System.Windows.Forms.Button
$buttonAddCustomEventLog.Text = "Add Custom Event Log"
$buttonAddCustomEventLog.Location = New-Object System.Drawing.Point(430, 60)
$buttonAddCustomEventLog.Size = New-Object System.Drawing.Size(150, 23)
$tabEventLog.Controls.Add($buttonAddCustomEventLog)

# Add available event logs to the dropdown
$eventLogs = Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog" -ErrorAction SilentlyContinue | ForEach-Object { $_.PSChildName }
if ($eventLogs) {
    foreach ($log in $eventLogs) {
        $comboEventLog.Items.Add($log)
    }
    if ($comboEventLog.Items.Count -gt 0) {
        $comboEventLog.SelectedIndex = 0
    }
} else {
    $comboEventLog.Items.Add("<No Event Logs Found>")
    $comboEventLog.SelectedIndex = 0
}

$labelSD = New-Object System.Windows.Forms.Label
$labelSD.Text = "Security Descriptor:"
$labelSD.Location = New-Object System.Drawing.Point(20, 100)
$labelSD.AutoSize = $true
$tabEventLog.Controls.Add($labelSD)

$textSD = New-Object System.Windows.Forms.TextBox
$textSD.Location = New-Object System.Drawing.Point(20, 120)
$textSD.Size = New-Object System.Drawing.Size(700, 80)
$textSD.Multiline = $true
$textSD.ScrollBars = "Vertical"
$tabEventLog.Controls.Add($textSD)

# Default security descriptor for common read/write access
$textSD.Text = "O:BAG:SYD:(A;;0xf0007;;;SY)(A;;0x7;;;BA)(A;;0x7;;;SO)(A;;0x3;;;IU)(A;;0x3;;;SU)(A;;0x3;;;S-1-5-3)(A;;0x3;;;S-1-5-33)(A;;0x1;;;S-1-5-32-573)"

$labelInfo = New-Object System.Windows.Forms.Label
$labelInfo.Text = "Common permission values: 0xf0007 = Full control, 0x7 = Read/Write, 0x3 = Read"
$labelInfo.Location = New-Object System.Drawing.Point(20, 210)
$labelInfo.AutoSize = $true
$tabEventLog.Controls.Add($labelInfo)

$buttonGetCurrentSD = New-Object System.Windows.Forms.Button
$buttonGetCurrentSD.Text = "Get Current Security Descriptor"
$buttonGetCurrentSD.Location = New-Object System.Drawing.Point(20, 250)
$buttonGetCurrentSD.Size = New-Object System.Drawing.Size(200, 30)
$tabEventLog.Controls.Add($buttonGetCurrentSD)

$buttonCreateEventLogGPO = New-Object System.Windows.Forms.Button
$buttonCreateEventLogGPO.Text = "Create Event Log GPO"
$buttonCreateEventLogGPO.Location = New-Object System.Drawing.Point(250, 480)
$buttonCreateEventLogGPO.Size = New-Object System.Drawing.Size(200, 30)
$tabEventLog.Controls.Add($buttonCreateEventLogGPO)

$buttonImportSD = New-Object System.Windows.Forms.Button
$buttonImportSD.Text = "Import From File"
$buttonImportSD.Location = New-Object System.Drawing.Point(480, 250)
$buttonImportSD.Size = New-Object System.Drawing.Size(200, 30)
$tabEventLog.Controls.Add($buttonImportSD)

# Permission buttons section
$groupBoxPermissions = New-Object System.Windows.Forms.GroupBox
$groupBoxPermissions.Text = "Quick Permission Settings"
$groupBoxPermissions.Location = New-Object System.Drawing.Point(20, 300)
$groupBoxPermissions.Size = New-Object System.Drawing.Size(700, 130)
$tabEventLog.Controls.Add($groupBoxPermissions)

# Add 'Apply to:' label
$labelApplyTo = New-Object System.Windows.Forms.Label
$labelApplyTo.Text = "Apply to:"
$labelApplyTo.Location = New-Object System.Drawing.Point(10, 25)
$labelApplyTo.AutoSize = $true
$groupBoxPermissions.Controls.Add($labelApplyTo)

# Add principal combo box
$comboPrincipal = New-Object System.Windows.Forms.ComboBox
$comboPrincipal.Items.Add("Everyone")
$comboPrincipal.Items.Add("Authenticated Users")
$comboPrincipal.Items.Add("Interactive Users")
$comboPrincipal.Items.Add("Network")
$comboPrincipal.Items.Add("Service")
$comboPrincipal.Items.Add("Batch")
$comboPrincipal.Location = New-Object System.Drawing.Point(80, 22)
$comboPrincipal.Size = New-Object System.Drawing.Size(200, 21)
$comboPrincipal.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboPrincipal.SelectedIndex = 0
$groupBoxPermissions.Controls.Add($comboPrincipal)

# Custom SID field
$labelCustomSID = New-Object System.Windows.Forms.Label
$labelCustomSID.Text = "Custom SID:"
$labelCustomSID.Location = New-Object System.Drawing.Point(300, 25)
$labelCustomSID.AutoSize = $true
$groupBoxPermissions.Controls.Add($labelCustomSID)

$textCustomSID = New-Object System.Windows.Forms.TextBox
$textCustomSID.Location = New-Object System.Drawing.Point(380, 22)
$textCustomSID.Size = New-Object System.Drawing.Size(300, 21)
$textCustomSID.PlaceholderText = "Enter SID or group name (S-1-5-...)"
$groupBoxPermissions.Controls.Add($textCustomSID)

# Add permission buttons
$buttonFullAccess = New-Object System.Windows.Forms.Button
$buttonFullAccess.Text = "Apply Full Access"
$buttonFullAccess.Location = New-Object System.Drawing.Point(10, 60)
$buttonFullAccess.Size = New-Object System.Drawing.Size(160, 30)
$groupBoxPermissions.Controls.Add($buttonFullAccess)

$buttonReadWrite = New-Object System.Windows.Forms.Button
$buttonReadWrite.Text = "Apply Read/Write"
$buttonReadWrite.Location = New-Object System.Drawing.Point(180, 60)
$buttonReadWrite.Size = New-Object System.Drawing.Size(160, 30)
$groupBoxPermissions.Controls.Add($buttonReadWrite)

$buttonReadOnly = New-Object System.Windows.Forms.Button
$buttonReadOnly.Text = "Apply Read Only"
$buttonReadOnly.Location = New-Object System.Drawing.Point(350, 60)
$buttonReadOnly.Size = New-Object System.Drawing.Size(160, 30)
$groupBoxPermissions.Controls.Add($buttonReadOnly)

$buttonImportedDefault = New-Object System.Windows.Forms.Button
$buttonImportedDefault.Text = "Apply Default Template"
$buttonImportedDefault.Location = New-Object System.Drawing.Point(520, 60)
$buttonImportedDefault.Size = New-Object System.Drawing.Size(160, 30)
$groupBoxPermissions.Controls.Add($buttonImportedDefault)

# Checkbox for creating the required registry keys for custom event logs
$checkboxCreateEventSource = New-Object System.Windows.Forms.CheckBox
$checkboxCreateEventSource.Text = "Create required registry keys for custom event log"
$checkboxCreateEventSource.Location = New-Object System.Drawing.Point(20, 440)
$checkboxCreateEventSource.Size = New-Object System.Drawing.Size(350, 30)
$checkboxCreateEventSource.Checked = $false
$tabEventLog.Controls.Add($checkboxCreateEventSource)

# ------------------------------
# Software Installation Tab Content
# ------------------------------
$labelGPONameSoftware = New-Object System.Windows.Forms.Label
$labelGPONameSoftware.Text = "GPO Name:"
$labelGPONameSoftware.Location = New-Object System.Drawing.Point(20, 20)
$labelGPONameSoftware.AutoSize = $true
$tabSoftware.Controls.Add($labelGPONameSoftware)

$comboGPONameSoftware = New-Object System.Windows.Forms.ComboBox
$comboGPONameSoftware.Location = New-Object System.Drawing.Point(120, 20)
$comboGPONameSoftware.Size = New-Object System.Drawing.Size(300, 20)
$comboGPONameSoftware.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabSoftware.Controls.Add($comboGPONameSoftware)

$textGPONameSoftware = New-Object System.Windows.Forms.TextBox
$textGPONameSoftware.Location = New-Object System.Drawing.Point(430, 20)
$textGPONameSoftware.Size = New-Object System.Drawing.Size(250, 20)
$textGPONameSoftware.Enabled = $true
$tabSoftware.Controls.Add($textGPONameSoftware)

$buttonRefreshSoftwareGPOs = New-Object System.Windows.Forms.Button
$buttonRefreshSoftwareGPOs.Text = "↻"
$buttonRefreshSoftwareGPOs.Location = New-Object System.Drawing.Point(690, 19)
$buttonRefreshSoftwareGPOs.Size = New-Object System.Drawing.Size(30, 23)
$tabSoftware.Controls.Add($buttonRefreshSoftwareGPOs)

$labelPackagePath = New-Object System.Windows.Forms.Label
$labelPackagePath.Text = "Package Path:"
$labelPackagePath.Location = New-Object System.Drawing.Point(20, 60)
$labelPackagePath.AutoSize = $true
$tabSoftware.Controls.Add($labelPackagePath)

$textPackagePath = New-Object System.Windows.Forms.TextBox
$textPackagePath.Location = New-Object System.Drawing.Point(120, 60)
$textPackagePath.Size = New-Object System.Drawing.Size(500, 20)
$tabSoftware.Controls.Add($textPackagePath)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "Browse..."
$buttonBrowse.Location = New-Object System.Drawing.Point(630, 60)
$buttonBrowse.Size = New-Object System.Drawing.Size(80, 23)
$tabSoftware.Controls.Add($buttonBrowse)

$labelDeploymentType = New-Object System.Windows.Forms.Label
$labelDeploymentType.Text = "Deployment Type:"
$labelDeploymentType.Location = New-Object System.Drawing.Point(20, 100)
$labelDeploymentType.AutoSize = $true
$tabSoftware.Controls.Add($labelDeploymentType)

$comboDeploymentType = New-Object System.Windows.Forms.ComboBox
$comboDeploymentType.Items.Add("Published")
$comboDeploymentType.Items.Add("Assigned")
$comboDeploymentType.SelectedIndex = 1
$comboDeploymentType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboDeploymentType.Location = New-Object System.Drawing.Point(120, 100)
$comboDeploymentType.Size = New-Object System.Drawing.Size(200, 20)
$tabSoftware.Controls.Add($comboDeploymentType)

$buttonCreateSoftwareGPO = New-Object System.Windows.Forms.Button
$buttonCreateSoftwareGPO.Text = "Create Software Installation GPO"
$buttonCreateSoftwareGPO.Location = New-Object System.Drawing.Point(120, 150)
$buttonCreateSoftwareGPO.Size = New-Object System.Drawing.Size(250, 30)
$tabSoftware.Controls.Add($buttonCreateSoftwareGPO)

# ------------------------------
# Firewall Rules Tab Content
# ------------------------------
$labelGPONameFirewall = New-Object System.Windows.Forms.Label
$labelGPONameFirewall.Text = "GPO Name:"
$labelGPONameFirewall.Location = New-Object System.Drawing.Point(20, 20)
$labelGPONameFirewall.AutoSize = $true
$tabFirewall.Controls.Add($labelGPONameFirewall)

$comboGPONameFirewall = New-Object System.Windows.Forms.ComboBox
$comboGPONameFirewall.Location = New-Object System.Drawing.Point(120, 20)
$comboGPONameFirewall.Size = New-Object System.Drawing.Size(300, 20)
$comboGPONameFirewall.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabFirewall.Controls.Add($comboGPONameFirewall)

$textGPONameFirewall = New-Object System.Windows.Forms.TextBox
$textGPONameFirewall.Location = New-Object System.Drawing.Point(430, 20)
$textGPONameFirewall.Size = New-Object System.Drawing.Size(250, 20)
$textGPONameFirewall.Enabled = $true
$tabFirewall.Controls.Add($textGPONameFirewall)

$buttonRefreshFirewallGPOs = New-Object System.Windows.Forms.Button
$buttonRefreshFirewallGPOs.Text = "↻"
$buttonRefreshFirewallGPOs.Location = New-Object System.Drawing.Point(690, 19)
$buttonRefreshFirewallGPOs.Size = New-Object System.Drawing.Size(30, 23)
$tabFirewall.Controls.Add($buttonRefreshFirewallGPOs)

$labelRuleName = New-Object System.Windows.Forms.Label
$labelRuleName.Text = "Rule Display Name:"
$labelRuleName.Location = New-Object System.Drawing.Point(20, 60)
$labelRuleName.AutoSize = $true
$tabFirewall.Controls.Add($labelRuleName)

$textRuleName = New-Object System.Windows.Forms.TextBox
$textRuleName.Location = New-Object System.Drawing.Point(120, 60)
$textRuleName.Size = New-Object System.Drawing.Size(300, 20)
$tabFirewall.Controls.Add($textRuleName)

$labelDirection = New-Object System.Windows.Forms.Label
$labelDirection.Text = "Direction:"
$labelDirection.Location = New-Object System.Drawing.Point(20, 100)
$labelDirection.AutoSize = $true
$tabFirewall.Controls.Add($labelDirection)

$comboDirection = New-Object System.Windows.Forms.ComboBox
$comboDirection.Items.Add("Inbound")
$comboDirection.Items.Add("Outbound")
$comboDirection.SelectedIndex = 0
$comboDirection.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboDirection.Location = New-Object System.Drawing.Point(120, 100)
$comboDirection.Size = New-Object System.Drawing.Size(200, 20)
$tabFirewall.Controls.Add($comboDirection)

$labelAction = New-Object System.Windows.Forms.Label
$labelAction.Text = "Action:"
$labelAction.Location = New-Object System.Drawing.Point(20, 140)
$labelAction.AutoSize = $true
$tabFirewall.Controls.Add($labelAction)

$comboAction = New-Object System.Windows.Forms.ComboBox
$comboAction.Items.Add("Allow")
$comboAction.Items.Add("Block")
$comboAction.SelectedIndex = 0
$comboAction.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboAction.Location = New-Object System.Drawing.Point(120, 140)
$comboAction.Size = New-Object System.Drawing.Size(200, 20)
$tabFirewall.Controls.Add($comboAction)

$labelProtocol = New-Object System.Windows.Forms.Label
$labelProtocol.Text = "Protocol:"
$labelProtocol.Location = New-Object System.Drawing.Point(20, 180)
$labelProtocol.AutoSize = $true
$tabFirewall.Controls.Add($labelProtocol)

$comboProtocol = New-Object System.Windows.Forms.ComboBox
$comboProtocol.Items.Add("TCP")
$comboProtocol.Items.Add("UDP")
$comboProtocol.SelectedIndex = 0
$comboProtocol.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboProtocol.Location = New-Object System.Drawing.Point(120, 180)
$comboProtocol.Size = New-Object System.Drawing.Size(200, 20)
$tabFirewall.Controls.Add($comboProtocol)

$labelPorts = New-Object System.Windows.Forms.Label
$labelPorts.Text = "Local Ports (comma-separated):"
$labelPorts.Location = New-Object System.Drawing.Point(20, 220)
$labelPorts.AutoSize = $true
$tabFirewall.Controls.Add($labelPorts)

$textPorts = New-Object System.Windows.Forms.TextBox
$textPorts.Location = New-Object System.Drawing.Point(180, 220)
$textPorts.Size = New-Object System.Drawing.Size(200, 20)
$tabFirewall.Controls.Add($textPorts)

$buttonCreateFirewallGPO = New-Object System.Windows.Forms.Button
$buttonCreateFirewallGPO.Text = "Create Firewall Rule GPO"
$buttonCreateFirewallGPO.Location = New-Object System.Drawing.Point(120, 260)
$buttonCreateFirewallGPO.Size = New-Object System.Drawing.Size(250, 30)
$tabFirewall.Controls.Add($buttonCreateFirewallGPO)

# ------------------------------
# Registry Settings Tab Content
# ------------------------------
$labelGPONameRegistry = New-Object System.Windows.Forms.Label
$labelGPONameRegistry.Text = "GPO Name:"
$labelGPONameRegistry.Location = New-Object System.Drawing.Point(20, 20)
$labelGPONameRegistry.AutoSize = $true
$tabRegistry.Controls.Add($labelGPONameRegistry)

$comboGPONameRegistry = New-Object System.Windows.Forms.ComboBox
$comboGPONameRegistry.Location = New-Object System.Drawing.Point(120, 20)
$comboGPONameRegistry.Size = New-Object System.Drawing.Size(300, 20)
$comboGPONameRegistry.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabRegistry.Controls.Add($comboGPONameRegistry)

$textGPONameRegistry = New-Object System.Windows.Forms.TextBox
$textGPONameRegistry.Location = New-Object System.Drawing.Point(430, 20)
$textGPONameRegistry.Size = New-Object System.Drawing.Size(250, 20)
$textGPONameRegistry.Enabled = $true
$tabRegistry.Controls.Add($textGPONameRegistry)

$buttonRefreshRegistryGPOs = New-Object System.Windows.Forms.Button
$buttonRefreshRegistryGPOs.Text = "↻"
$buttonRefreshRegistryGPOs.Location = New-Object System.Drawing.Point(690, 19)
$buttonRefreshRegistryGPOs.Size = New-Object System.Drawing.Size(30, 23)
$tabRegistry.Controls.Add($buttonRefreshRegistryGPOs)

$labelRegKey = New-Object System.Windows.Forms.Label
$labelRegKey.Text = "Registry Key:"
$labelRegKey.Location = New-Object System.Drawing.Point(20, 60)
$labelRegKey.AutoSize = $true
$tabRegistry.Controls.Add($labelRegKey)

$textRegKey = New-Object System.Windows.Forms.TextBox
$textRegKey.Location = New-Object System.Drawing.Point(120, 60)
$textRegKey.Size = New-Object System.Drawing.Size(600, 20)
$textRegKey.Text = "HKLM\SOFTWARE\Policies\MyCompany"
$tabRegistry.Controls.Add($textRegKey)

$labelValueName = New-Object System.Windows.Forms.Label
$labelValueName.Text = "Value Name:"
$labelValueName.Location = New-Object System.Drawing.Point(20, 100)
$labelValueName.AutoSize = $true
$tabRegistry.Controls.Add($labelValueName)

$textValueName = New-Object System.Windows.Forms.TextBox
$textValueName.Location = New-Object System.Drawing.Point(120, 100)
$textValueName.Size = New-Object System.Drawing.Size(300, 20)
$tabRegistry.Controls.Add($textValueName)

$labelValueType = New-Object System.Windows.Forms.Label
$labelValueType.Text = "Value Type:"
$labelValueType.Location = New-Object System.Drawing.Point(20, 140)
$labelValueType.AutoSize = $true
$tabRegistry.Controls.Add($labelValueType)

$comboValueType = New-Object System.Windows.Forms.ComboBox
$comboValueType.Items.Add("String")
$comboValueType.Items.Add("ExpandString")
$comboValueType.Items.Add("Binary")
$comboValueType.Items.Add("DWord")
$comboValueType.Items.Add("QWord")
$comboValueType.Items.Add("MultiString")
$comboValueType.SelectedIndex = 0
$comboValueType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboValueType.Location = New-Object System.Drawing.Point(120, 140)
$comboValueType.Size = New-Object System.Drawing.Size(200, 20)
$tabRegistry.Controls.Add($comboValueType)

$labelValueData = New-Object System.Windows.Forms.Label
$labelValueData.Text = "Value Data:"
$labelValueData.Location = New-Object System.Drawing.Point(20, 180)
$labelValueData.AutoSize = $true
$tabRegistry.Controls.Add($labelValueData)

$textValueData = New-Object System.Windows.Forms.TextBox
$textValueData.Location = New-Object System.Drawing.Point(120, 180)
$textValueData.Size = New-Object System.Drawing.Size(600, 60)
$textValueData.Multiline = $true
$tabRegistry.Controls.Add($textValueData)

$buttonAddRegSetting = New-Object System.Windows.Forms.Button
$buttonAddRegSetting.Text = "Add to List"
$buttonAddRegSetting.Location = New-Object System.Drawing.Point(120, 250)
$buttonAddRegSetting.Size = New-Object System.Drawing.Size(100, 30)
$tabRegistry.Controls.Add($buttonAddRegSetting)

$listViewRegSettings = New-Object System.Windows.Forms.ListView
$listViewRegSettings.Location = New-Object System.Drawing.Point(20, 290)
$listViewRegSettings.Size = New-Object System.Drawing.Size(700, 180)
$listViewRegSettings.View = [System.Windows.Forms.View]::Details
$listViewRegSettings.FullRowSelect = $true
$listViewRegSettings.Columns.Add("Key", 300)
$listViewRegSettings.Columns.Add("Name", 150)
$listViewRegSettings.Columns.Add("Type", 100)
$listViewRegSettings.Columns.Add("Data", 150)
$tabRegistry.Controls.Add($listViewRegSettings)

$buttonRemoveRegSetting = New-Object System.Windows.Forms.Button
$buttonRemoveRegSetting.Text = "Remove Selected"
$buttonRemoveRegSetting.Location = New-Object System.Drawing.Point(20, 480)
$buttonRemoveRegSetting.Size = New-Object System.Drawing.Size(120, 30)
$tabRegistry.Controls.Add($buttonRemoveRegSetting)

$buttonCreateRegGPO = New-Object System.Windows.Forms.Button
$buttonCreateRegGPO.Text = "Create Registry GPO"
$buttonCreateRegGPO.Location = New-Object System.Drawing.Point(550, 480)
$buttonCreateRegGPO.Size = New-Object System.Drawing.Size(170, 30)
$tabRegistry.Controls.Add($buttonCreateRegGPO)

# Add the tab control to the form
$form.Controls.Add($tabControl)

# Add event handlers

# Initialize GPO ComboBoxes
Update-GPOComboBox -ComboBox $comboGPONameEventLog
Update-GPOComboBox -ComboBox $comboGPONameSoftware
Update-GPOComboBox -ComboBox $comboGPONameFirewall
Update-GPOComboBox -ComboBox $comboGPONameRegistry

# Event handler for GPO ComboBox selection change
$comboGPONameEventLog.add_SelectedIndexChanged({
    if ($comboGPONameEventLog.SelectedIndex -eq 0) {
        $textGPONameEventLog.Enabled = $true
        $textGPONameEventLog.Clear()
    } else {
        $textGPONameEventLog.Enabled = $false
        $textGPONameEventLog.Text = $comboGPONameEventLog.SelectedItem.ToString()
    }
})

$comboGPONameSoftware.add_SelectedIndexChanged({
    if ($comboGPONameSoftware.SelectedIndex -eq 0) {
        $textGPONameSoftware.Enabled = $true
        $textGPONameSoftware.Clear()
    } else {
        $textGPONameSoftware.Enabled = $false
        $textGPONameSoftware.Text = $comboGPONameSoftware.SelectedItem.ToString()
    }
})

$comboGPONameFirewall.add_SelectedIndexChanged({
    if ($comboGPONameFirewall.SelectedIndex -eq 0) {
        $textGPONameFirewall.Enabled = $true
        $textGPONameFirewall.Clear()
    } else {
        $textGPONameFirewall.Enabled = $false
        $textGPONameFirewall.Text = $comboGPONameFirewall.SelectedItem.ToString()
    }
})

$comboGPONameRegistry.add_SelectedIndexChanged({
    if ($comboGPONameRegistry.SelectedIndex -eq 0) {
        $textGPONameRegistry.Enabled = $true
        $textGPONameRegistry.Clear()
    } else {
        $textGPONameRegistry.Enabled = $false
        $textGPONameRegistry.Text = $comboGPONameRegistry.SelectedItem.ToString()
    }
})

# Refresh button handlers
$buttonRefreshEventLogGPOs.Add_Click({ Update-GPOComboBox -ComboBox $comboGPONameEventLog })
$buttonRefreshSoftwareGPOs.Add_Click({ Update-GPOComboBox -ComboBox $comboGPONameSoftware })
$buttonRefreshFirewallGPOs.Add_Click({ Update-GPOComboBox -ComboBox $comboGPONameFirewall })
$buttonRefreshRegistryGPOs.Add_Click({ Update-GPOComboBox -ComboBox $comboGPONameRegistry })

# Event Log tab handlers
$buttonGetCurrentSD.Add_Click({
    $selectedLog = $comboEventLog.SelectedItem
    
    if ($selectedLog) {
        try {
            $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\$selectedLog"
            $sd = (Get-ItemProperty -Path $registryPath -Name "CustomSD" -ErrorAction SilentlyContinue).CustomSD
            
            if ($sd) {
                $textSD.Text = $sd
                $statusLabel.Text = "Retrieved current security descriptor for $selectedLog"
            }
            else {
                $statusLabel.Text = "No CustomSD found for $selectedLog. Using default template."
            }
        }
        catch {
            $statusLabel.Text = "Error: $($_.Exception.Message)"
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please select an Event Log.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# Add handler for the custom event log button
$buttonAddCustomEventLog.Add_Click({
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
            foreach ($item in $comboEventLog.Items) {
                if ($item -eq $customLogName) {
                    $exists = $true
                    break
                }
            }
            
            if (-not $exists) {
                # Remove placeholder if it exists
                if ($comboEventLog.Items.Count -eq 1 -and $comboEventLog.Items[0] -eq "<No Event Logs Found>") {
                    $comboEventLog.Items.Clear()
                }
                
                # Add the custom log name and select it
                $comboEventLog.Items.Add($customLogName)
            }
            
            # Select the item
            $comboEventLog.SelectedItem = $customLogName
            $statusLabel.Text = "Custom event log '$customLogName' added to the selection list."
            
            # Recommend creating registry keys for custom event logs
            $checkboxCreateEventSource.Checked = $true
            [System.Windows.Forms.MessageBox]::Show("For custom event logs, it's recommended to create the required registry keys. The checkbox has been checked for you.", "Custom Event Log", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
})

# Modify the import button handler to handle custom event logs
$buttonImportSD.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "GPO Creator Files (*.gpodata)|*.gpodata|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Event Log Security Data File"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $importedData = Import-GPOData -FilePath $openFileDialog.FileName
        
        if ($importedData -and $importedData.Type -eq "EventLogSecurity") {
            # Populate the form with the imported data
            $comboGPONameEventLog.SelectedIndex = 0
            $textGPONameEventLog.Text = "EventLogSecurity_$($importedData.LogName)"
            
            # Find the log in the combobox
            $logName = $importedData.LogName
            $logIndex = -1
            
            # Loop through all items to find a match regardless of case
            for ($i = 0; $i -lt $comboEventLog.Items.Count; $i++) {
                if ($comboEventLog.Items[$i].ToString() -eq $logName) {
                    $logIndex = $i
                    break
                }
            }
            
            if ($logIndex -ge 0) {
                $comboEventLog.SelectedIndex = $logIndex
            } else {
                # Add the custom log to the dropdown since it doesn't exist
                if ($comboEventLog.Items.Count -eq 1 -and $comboEventLog.Items[0] -eq "<No Event Logs Found>") {
                    $comboEventLog.Items.Clear()
                }
                
                $comboEventLog.Items.Add($logName)
                $comboEventLog.SelectedItem = $logName
                
                $statusLabel.Text = "Imported event log security data for $($importedData.LogName) (added as custom event log)"
            }
            
            $textSD.Text = $importedData.SecurityDescriptor
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("The selected file does not contain valid event log security data.", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

$buttonCreateEventLogGPO.Add_Click({
    $gpoName = $textGPONameEventLog.Text.Trim()
    $logName = $comboEventLog.SelectedItem
    $securityDescriptor = $textSD.Text.Trim()
    $createEventSource = $checkboxCreateEventSource.Checked
    
    if (-not $gpoName) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a GPO Name.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if (-not $logName) {
        [System.Windows.Forms.MessageBox]::Show("Please select an Event Log.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if (-not $securityDescriptor) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a Security Descriptor.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $result = New-EventLogSecurityGPO -GPOName $gpoName -LogName $logName -SecurityDescriptor $securityDescriptor -CreateEventSource $createEventSource
    
    if ($result.Success) {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $statusLabel.Text = "Created GPO: $($result.GPODisplayName)"
        
        # Refresh the GPO lists
        Update-GPOComboBox -ComboBox $comboGPONameEventLog
        Update-GPOComboBox -ComboBox $comboGPONameSoftware
        Update-GPOComboBox -ComboBox $comboGPONameFirewall
        Update-GPOComboBox -ComboBox $comboGPONameRegistry
    }
    else {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Failed to create GPO: $($result.Message)"
    }
})

# Permission button event handlers
$buttonFullAccess.Add_Click({
    $principal = if ($textCustomSID.Text.Trim()) {
        $textCustomSID.Text.Trim()
    } else {
        $comboPrincipal.SelectedItem.ToString()
    }
    
    $sd = Get-EventLogSecurityDescriptor -PermissionLevel "FullAccess" -Principal $principal
    $textSD.Text = $sd
    $statusLabel.Text = "Applied Full Access permissions for $principal"
})

$buttonReadWrite.Add_Click({
    $principal = if ($textCustomSID.Text.Trim()) {
        $textCustomSID.Text.Trim()
    } else {
        $comboPrincipal.SelectedItem.ToString()
    }
    
    $sd = Get-EventLogSecurityDescriptor -PermissionLevel "ReadWrite" -Principal $principal
    $textSD.Text = $sd
    $statusLabel.Text = "Applied Read/Write permissions for $principal"
})

$buttonReadOnly.Add_Click({
    $principal = if ($textCustomSID.Text.Trim()) {
        $textCustomSID.Text.Trim()
    } else {
        $comboPrincipal.SelectedItem.ToString()
    }
    
    $sd = Get-EventLogSecurityDescriptor -PermissionLevel "Read" -Principal $principal
    $textSD.Text = $sd
    $statusLabel.Text = "Applied Read Only permissions for $principal"
})

$buttonImportedDefault.Add_Click({
    # Default security descriptor for common read/write access
    $textSD.Text = "O:BAG:SYD:(A;;0xf0007;;;SY)(A;;0x7;;;BA)(A;;0x7;;;SO)(A;;0x3;;;IU)(A;;0x3;;;SU)(A;;0x3;;;S-1-5-3)(A;;0x3;;;S-1-5-33)(A;;0x1;;;S-1-5-32-573)"
    $statusLabel.Text = "Applied default security descriptor template"
})

# Software Installation tab handlers
$buttonBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "MSI Files (*.msi)|*.msi|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select an MSI Package"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textPackagePath.Text = $openFileDialog.FileName
    }
})

$buttonCreateSoftwareGPO.Add_Click({
    $gpoName = $textGPONameSoftware.Text.Trim()
    $packagePath = $textPackagePath.Text.Trim()
    $deploymentType = $comboDeploymentType.SelectedItem.ToString()
    
    if (-not $gpoName) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a GPO Name.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if (-not $packagePath -or -not (Test-Path $packagePath)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid package path.", "Invalid Path", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $result = New-SoftwareInstallationGPO -GPOName $gpoName -PackagePath $packagePath -DeploymentType $deploymentType
    
    if ($result.Success) {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $statusLabel.Text = "Created GPO: $($result.GPODisplayName)"
        
        # Refresh the GPO lists
        Update-GPOComboBox -ComboBox $comboGPONameEventLog
        Update-GPOComboBox -ComboBox $comboGPONameSoftware
        Update-GPOComboBox -ComboBox $comboGPONameFirewall
        Update-GPOComboBox -ComboBox $comboGPONameRegistry
    }
    else {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Failed to create GPO: $($result.Message)"
    }
})

# Firewall Rule tab handlers
$buttonCreateFirewallGPO.Add_Click({
    $gpoName = $textGPONameFirewall.Text.Trim()
    $ruleName = $textRuleName.Text.Trim()
    $direction = $comboDirection.SelectedItem.ToString()
    $action = $comboAction.SelectedItem.ToString()
    $protocol = $comboProtocol.SelectedItem.ToString()
    $portsText = $textPorts.Text.Trim()
    
    if (-not $gpoName) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a GPO Name.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if (-not $ruleName) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a Rule Display Name.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if (-not $portsText) {
        [System.Windows.Forms.MessageBox]::Show("Please enter port number(s).", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $portArray = $portsText -split ',' | ForEach-Object { [int]$_.Trim() }
    
    $result = New-FirewallRuleGPO -GPOName $gpoName -DisplayName $ruleName -Direction $direction -Action $action -LocalPorts $portArray -Protocol $protocol
    
    if ($result.Success) {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $statusLabel.Text = "Created GPO: $($result.GPODisplayName)"
        
        # Refresh the GPO lists
        Update-GPOComboBox -ComboBox $comboGPONameEventLog
        Update-GPOComboBox -ComboBox $comboGPONameSoftware
        Update-GPOComboBox -ComboBox $comboGPONameFirewall
        Update-GPOComboBox -ComboBox $comboGPONameRegistry
    }
    else {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Failed to create GPO: $($result.Message)"
    }
})

# Registry Settings tab handlers
$buttonAddRegSetting.Add_Click({
    $keyPath = $textRegKey.Text.Trim()
    $valueName = $textValueName.Text.Trim()
    $valueType = $comboValueType.SelectedItem.ToString()
    $valueData = $textValueData.Text.Trim()
    
    if (-not $keyPath) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a Registry Key.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if (-not $valueName) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a Value Name.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $item = New-Object System.Windows.Forms.ListViewItem($keyPath)
    $item.SubItems.Add($valueName)
    $item.SubItems.Add($valueType)
    $item.SubItems.Add($valueData)
    $listViewRegSettings.Items.Add($item)
    
    $textValueName.Clear()
    $textValueData.Clear()
})

$buttonRemoveRegSetting.Add_Click({
    foreach ($item in $listViewRegSettings.SelectedItems) {
        $listViewRegSettings.Items.Remove($item)
    }
})

$buttonCreateRegGPO.Add_Click({
    $gpoName = $textGPONameRegistry.Text.Trim()
    
    if (-not $gpoName) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a GPO Name.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if ($listViewRegSettings.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please add at least one registry setting.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $registrySettings = @()
    foreach ($item in $listViewRegSettings.Items) {
        $keyPath = $item.Text
        $valueName = $item.SubItems[1].Text
        $valueType = $item.SubItems[2].Text
        $valueData = $item.SubItems[3].Text
        
        $registrySettings += @{
            KeyPath = $keyPath
            ValueName = $valueName
            ValueData = $valueData
            ValueType = $valueType
        }
    }
    
    $result = New-CustomGPO -GPOName $gpoName -Comment "Registry settings GPO" -RegistrySettings $registrySettings
    
    if ($result.Success) {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $statusLabel.Text = "Created GPO: $($result.GPODisplayName)"
        
        # Refresh the GPO lists
        Update-GPOComboBox -ComboBox $comboGPONameEventLog
        Update-GPOComboBox -ComboBox $comboGPONameSoftware
        Update-GPOComboBox -ComboBox $comboGPONameFirewall
        Update-GPOComboBox -ComboBox $comboGPONameRegistry
    }
    else {
        [System.Windows.Forms.MessageBox]::Show($result.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = "Failed to create GPO: $($result.Message)"
    }
})

# Show the form
$form.ShowDialog()
