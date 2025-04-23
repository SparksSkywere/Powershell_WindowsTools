#Requires -Version 5.0
#Requires -Modules ActiveDirectory

param(
    [Parameter(Mandatory = $false)]
    [string]$Username
)

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# Global variables
$script:userInfo = $null
$script:simulatedPolicies = @()
$script:redirectedFolders = @{}

# Function to verify user exists in AD
function Test-ADUser {
    param([string]$Username)
    
    try {
        $adUser = Get-ADUser -Identity $Username -Properties *
        return $adUser
    }
    catch {
        [System.Windows.MessageBox]::Show("User not found in Active Directory. Please check the username.", "User Not Found", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $null
    }
}

# Function to check password status and determine if change required
function Test-PasswordChangeRequired {
    param($UserObject)
    
    # In a real environment, we'd check PasswordExpired, PasswordLastSet, etc.
    # For simulation, we'll randomly decide
    $random = Get-Random -Minimum 1 -Maximum 10
    if ($random -le 3) { 
        return $true 
    }
    
    return $false
}

# Function to show login screen
function Show-LoginScreen {
    $loginForm = New-Object System.Windows.Forms.Form
    $loginForm.Text = "Windows Login"
    $loginForm.Size = New-Object System.Drawing.Size(400, 250)
    $loginForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $loginForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $loginForm.MaximizeBox = $false
    $loginForm.MinimizeBox = $false
    
    # Windows logo
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Size = New-Object System.Drawing.Size(70, 70)
    $pictureBox.Location = New-Object System.Drawing.Point(165, 20)
    $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    
    # Use a solid blue panel instead of an actual image
    $pictureBox.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    
    $loginForm.Controls.Add($pictureBox)
    
    # User label
    $userLabel = New-Object System.Windows.Forms.Label
    $userLabel.Location = New-Object System.Drawing.Point(100, 100)
    $userLabel.Size = New-Object System.Drawing.Size(200, 20)
    $userLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $userLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)
    if ($Username) {
        $userLabel.Text = $Username
        $script:userInfo = Test-ADUser -Username $Username
    } else {
        $userLabel.Text = "Enter Username"
    }
    $loginForm.Controls.Add($userLabel)
    
    # Username textbox (only shown if no username was provided)
    if (-not $Username) {
        $usernameTextBox = New-Object System.Windows.Forms.TextBox
        $usernameTextBox.Location = New-Object System.Drawing.Point(100, 125)
        $usernameTextBox.Size = New-Object System.Drawing.Size(200, 20)
        $usernameTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
        $loginForm.Controls.Add($usernameTextBox)
    }
    
    # Password field
    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Location = New-Object System.Drawing.Point(100, 150)
    $passwordLabel.Size = New-Object System.Drawing.Size(200, 20)
    $passwordLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $passwordLabel.Text = "Password"
    $loginForm.Controls.Add($passwordLabel)
    
    $passwordBox = New-Object System.Windows.Forms.MaskedTextBox
    $passwordBox.PasswordChar = "*"
    $passwordBox.Location = New-Object System.Drawing.Point(100, 170)
    $passwordBox.Size = New-Object System.Drawing.Size(200, 20)
    $passwordBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
    $loginForm.Controls.Add($passwordBox)
    
    # Sign in button
    $signInButton = New-Object System.Windows.Forms.Button
    $signInButton.Location = New-Object System.Drawing.Point(150, 200)
    $signInButton.Size = New-Object System.Drawing.Size(100, 30)
    $signInButton.Text = "Sign in"
    $signInButton.Add_Click({
        if (-not $Username -and $usernameTextBox.Text -eq "") {
            [System.Windows.MessageBox]::Show("Please enter a username", "Login Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ($passwordBox.Text -eq "") {
            [System.Windows.MessageBox]::Show("Please enter a password", "Login Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # If username wasn't provided as parameter, get it from textbox
        if (-not $Username) {
            $script:userInfo = Test-ADUser -Username $usernameTextBox.Text
            if (-not $script:userInfo) { return }
        }
        
        # Simulate login process
        Show-LoginProgress
        $loginForm.Close()
    })
    $loginForm.Controls.Add($signInButton)
    
    $loginForm.AcceptButton = $signInButton
    $loginForm.ShowDialog() | Out-Null
}

# Function to show login progress
function Show-LoginProgress {
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Logging in..."
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $progressForm.ControlBox = $false
    
    # Progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(50, 50)
    $progressBar.Size = New-Object System.Drawing.Size(300, 20)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progressBar.MarqueeAnimationSpeed = 30
    $progressForm.Controls.Add($progressBar)
    
    # Status label
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(50, 80)
    $statusLabel.Size = New-Object System.Drawing.Size(300, 20)
    $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $statusLabel.Text = "Applying user settings..."
    $progressForm.Controls.Add($statusLabel)
    
    # Show the form without blocking
    $progressForm.Show()
    $progressForm.Refresh()
    
    # Simulate login tasks
    Start-Sleep -Seconds 1
    $statusLabel.Text = "Loading user profile..."
    $progressForm.Refresh()
    Start-Sleep -Seconds 1
    
    $statusLabel.Text = "Applying group policies..."
    $progressForm.Refresh()
    Initialize-GroupPolicies
    Start-Sleep -Seconds 1
    
    $statusLabel.Text = "Setting up folder redirection..."
    $progressForm.Refresh()
    Initialize-FolderRedirection
    Start-Sleep -Seconds 1
    
    $progressForm.Close()
    
    # Check if password change is required
    if (Test-PasswordChangeRequired -UserObject $script:userInfo) {
        Show-PasswordChangePrompt
    }
    
    # Show the Desktop
    Show-SimulatedDesktop
}

# Function to show password change prompt
function Show-PasswordChangePrompt {
    $pwdForm = New-Object System.Windows.Forms.Form
    $pwdForm.Text = "Change Password"
    $pwdForm.Size = New-Object System.Drawing.Size(400, 250)
    $pwdForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $pwdForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $pwdForm.MaximizeBox = $false
    $pwdForm.MinimizeBox = $false
    
    # Current password
    $currentPwdLabel = New-Object System.Windows.Forms.Label
    $currentPwdLabel.Location = New-Object System.Drawing.Point(30, 20)
    $currentPwdLabel.Size = New-Object System.Drawing.Size(150, 20)
    $currentPwdLabel.Text = "Current password:"
    $pwdForm.Controls.Add($currentPwdLabel)
    
    $currentPwdBox = New-Object System.Windows.Forms.MaskedTextBox
    $currentPwdBox.PasswordChar = "*"
    $currentPwdBox.Location = New-Object System.Drawing.Point(180, 20)
    $currentPwdBox.Size = New-Object System.Drawing.Size(180, 20)
    $pwdForm.Controls.Add($currentPwdBox)
    
    # New password
    $newPwdLabel = New-Object System.Windows.Forms.Label
    $newPwdLabel.Location = New-Object System.Drawing.Point(30, 50)
    $newPwdLabel.Size = New-Object System.Drawing.Size(150, 20)
    $newPwdLabel.Text = "New password:"
    $pwdForm.Controls.Add($newPwdLabel)
    
    $newPwdBox = New-Object System.Windows.Forms.MaskedTextBox
    $newPwdBox.PasswordChar = "*"
    $newPwdBox.Location = New-Object System.Drawing.Point(180, 50)
    $newPwdBox.Size = New-Object System.Drawing.Size(180, 20)
    $pwdForm.Controls.Add($newPwdBox)
    
    # Confirm new password
    $confirmPwdLabel = New-Object System.Windows.Forms.Label
    $confirmPwdLabel.Location = New-Object System.Drawing.Point(30, 80)
    $confirmPwdLabel.Size = New-Object System.Drawing.Size(150, 20)
    $confirmPwdLabel.Text = "Confirm new password:"
    $pwdForm.Controls.Add($confirmPwdLabel)
    
    $confirmPwdBox = New-Object System.Windows.Forms.MaskedTextBox
    $confirmPwdBox.PasswordChar = "*"
    $confirmPwdBox.Location = New-Object System.Drawing.Point(180, 80)
    $confirmPwdBox.Size = New-Object System.Drawing.Size(180, 20)
    $pwdForm.Controls.Add($confirmPwdBox)
    
    # Password requirements
    $requirementsLabel = New-Object System.Windows.Forms.Label
    $requirementsLabel.Location = New-Object System.Drawing.Point(30, 110)
    $requirementsLabel.Size = New-Object System.Drawing.Size(330, 60)
    $requirementsLabel.Text = "Password must meet the following requirements:`r`n- At least 8 characters long`r`n- Include uppercase, lowercase, numbers, and symbols`r`n- Cannot contain your username"
    $pwdForm.Controls.Add($requirementsLabel)
    
    # Buttons
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(100, 180)
    $okButton.Size = New-Object System.Drawing.Size(90, 30)
    $okButton.Text = "OK"
    $okButton.Add_Click({
        if ($currentPwdBox.Text -eq "") {
            [System.Windows.MessageBox]::Show("Please enter your current password", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ($newPwdBox.Text -eq "") {
            [System.Windows.MessageBox]::Show("Please enter a new password", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ($newPwdBox.Text -ne $confirmPwdBox.Text) {
            [System.Windows.MessageBox]::Show("New passwords do not match", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Check password complexity
        if (-not (Test-PasswordComplexity -Password $newPwdBox.Text)) {
            [System.Windows.MessageBox]::Show("Password does not meet complexity requirements", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        [System.Windows.MessageBox]::Show("Password changed successfully", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        $pwdForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $pwdForm.Close()
    })
    $pwdForm.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(210, 180)
    $cancelButton.Size = New-Object System.Drawing.Size(90, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({
        $pwdForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $pwdForm.Close()
    })
    $pwdForm.Controls.Add($cancelButton)
    
    $pwdForm.AcceptButton = $okButton
    $pwdForm.CancelButton = $cancelButton
    
    $result = $pwdForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        # In real life, users might be forced to change password
        # For simulation, we'll just show a reminder
        [System.Windows.MessageBox]::Show("Your password must be changed soon. You will be prompted again at next login.", "Password Change Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
    }
}

# Function to test password complexity
function Test-PasswordComplexity {
    param([string]$Password)
    
    if ($Password.Length -lt 8) { return $false }
    if ($Password -notmatch "[A-Z]") { return $false }
    if ($Password -notmatch "[a-z]") { return $false }
    if ($Password -notmatch "[0-9]") { return $false }
    if ($Password -notmatch "[^a-zA-Z0-9]") { return $false }
    
    return $true
}

# Function to initialize simulated GPO settings
function Initialize-GroupPolicies {
    $script:simulatedPolicies = @(
        @{
            Name = "Password Policy"
            Description = "Enforces password complexity, history, and maximum age"
            Applied = $true
        },
        @{
            Name = "Folder Redirection"
            Description = "Redirects user folders to network locations"
            Applied = $true
        },
        @{
            Name = "Desktop Settings"
            Description = "Configures standard desktop icons and settings"
            Applied = $true
        },
        @{
            Name = "Security Settings"
            Description = "Enforces security measures like screen lock and device encryption"
            Applied = $true
        },
        @{
            Name = "Software Deployment"
            Description = "Deploys standard software packages to user's machine"
            Applied = $true
        }
    )
}

# Function to initialize folder redirection
function Initialize-FolderRedirection {
    $script:redirectedFolders = @{
        "Desktop" = "\\fileserver\users\$($script:userInfo.SamAccountName)\Desktop"
        "Documents" = "\\fileserver\users\$($script:userInfo.SamAccountName)\Documents"
        "Pictures" = "\\fileserver\users\$($script:userInfo.SamAccountName)\Pictures"
        "Downloads" = "\\fileserver\users\$($script:userInfo.SamAccountName)\Downloads"
    }
}

# Function to show simulated desktop
function Show-SimulatedDesktop {
    $desktopForm = New-Object System.Windows.Forms.Form
    $desktopForm.Text = "Windows Desktop - Simulated for $($script:userInfo.Name)"
    $desktopForm.Size = New-Object System.Drawing.Size(1024, 768)
    $desktopForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $desktopForm.BackColor = [System.Drawing.Color]::FromArgb(0, 103, 192)
    
    # Desktop icons - blue rectangles with labels as placeholders
    $iconSize = 80
    $icons = @(
        @{ Name = "This PC"; X = 20; Y = 20 },
        @{ Name = "Recycle Bin"; X = 20; Y = 120 },
        @{ Name = "File Explorer"; X = 20; Y = 220 },
        @{ Name = "Microsoft Edge"; X = 20; Y = 320 },
        @{ Name = "Office"; X = 20; Y = 420 }
    )
    
    foreach ($icon in $icons) {
        $iconPanel = New-Object System.Windows.Forms.Panel
        $iconPanel.Size = New-Object System.Drawing.Size(60, 60)
        $iconPanel.Location = New-Object System.Drawing.Point($icon.X, $icon.Y)
        $iconPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 144, 255)
        $desktopForm.Controls.Add($iconPanel)
        
        $iconLabel = New-Object System.Windows.Forms.Label
        $iconLabel.Text = $icon.Name
        $iconLabel.AutoSize = $false
        $iconLabel.Size = New-Object System.Drawing.Size(80, 20)
        $iconLabel.Location = New-Object System.Drawing.Point($icon.X - 10, $icon.Y + 65)
        $iconLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $iconLabel.ForeColor = [System.Drawing.Color]::White
        $desktopForm.Controls.Add($iconLabel)
    }
    
    # Taskbar
    $taskbar = New-Object System.Windows.Forms.Panel
    $taskbar.Size = New-Object System.Drawing.Size(1024, 40)
    $taskbar.Location = New-Object System.Drawing.Point(0, 728)
    $taskbar.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $desktopForm.Controls.Add($taskbar)
    
    # Start button
    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Size = New-Object System.Drawing.Size(40, 30)
    $startButton.Location = New-Object System.Drawing.Point(10, 734)
    $startButton.Text = "Start"
    $startButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $startButton.ForeColor = [System.Drawing.Color]::White
    $startButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $startButton.Add_Click({
        Show-StartMenu
    })
    $desktopForm.Controls.Add($startButton)
    
    # Time
    $timeLabel = New-Object System.Windows.Forms.Label
    $timeLabel.Size = New-Object System.Drawing.Size(70, 30)
    $timeLabel.Location = New-Object System.Drawing.Point(944, 734)
    $timeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $timeLabel.ForeColor = [System.Drawing.Color]::White
    $timeLabel.Text = Get-Date -Format "HH:mm"
    $desktopForm.Controls.Add($timeLabel)
    
    # Add a context menu for the desktop
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $viewMenuItem = $contextMenu.Items.Add("View")
    $sortByMenuItem = $contextMenu.Items.Add("Sort By")
    $refreshMenuItem = $contextMenu.Items.Add("Refresh")
    $contextMenu.Items.Add("-") # Separator
    $newMenuItem = $contextMenu.Items.Add("New")
    $contextMenu.Items.Add("-") # Separator
    $displaySettingsMenuItem = $contextMenu.Items.Add("Display Settings")
    $personalizationMenuItem = $contextMenu.Items.Add("Personalize")
    
    $desktopForm.ContextMenuStrip = $contextMenu
    
    # Menu items to open important dialogs
    $refreshMenuItem.Add_Click({
        [System.Windows.MessageBox]::Show("Desktop refreshed", "System", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    })
    
    $displaySettingsMenuItem.Add_Click({
        Show-GPOResults
    })
    
    $personalizationMenuItem.Add_Click({
        Show-FolderRedirection
    })
    
    # Add additional buttons to taskbar
    $taskbarIcons = @(
        @{ Text = "S"; ToolTip = "Search" },
        @{ Text = "F"; ToolTip = "File Explorer" },
        @{ Text = "E"; ToolTip = "Edge" },
        @{ Text = "M"; ToolTip = "Mail" },
        @{ Text = "C"; ToolTip = "Settings" }
    )
    
    $posX = 60
    foreach ($icon in $taskbarIcons) {
        $iconBtn = New-Object System.Windows.Forms.Button
        $iconBtn.Size = New-Object System.Drawing.Size(35, 30)
        $iconBtn.Location = New-Object System.Drawing.Point($posX, 734)
        $iconBtn.Text = $icon.Text
        $iconBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $iconBtn.BackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
        $iconBtn.ForeColor = [System.Drawing.Color]::White
        $iconBtn.FlatAppearance.BorderSize = 0
        $iconBtn.Tag = $icon.ToolTip
        $iconBtn.Add_Click({
            $clickedIcon = $this.Tag
            if ($clickedIcon -eq "Settings") {
                Show-GPOResults
            } elseif ($clickedIcon -eq "File Explorer") {
                Show-FolderRedirection
            } else {
                [System.Windows.MessageBox]::Show("Opening $clickedIcon", "Simulation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        })
        $desktopForm.Controls.Add($iconBtn)
        $posX += 40
    }
    
    # Show the desktop
    $desktopForm.ShowDialog() | Out-Null
}

# Function to show Start Menu
function Show-StartMenu {
    $startMenuForm = New-Object System.Windows.Forms.Form
    $startMenuForm.Text = "Start Menu"
    $startMenuForm.Size = New-Object System.Drawing.Size(350, 500)
    $startMenuForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $startMenuForm.Location = New-Object System.Drawing.Point(10, 228) # Position above start button
    $startMenuForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $startMenuForm.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
    $startMenuForm.TopMost = $true
    
    # User info at top
    $userPanel = New-Object System.Windows.Forms.Panel
    $userPanel.Size = New-Object System.Drawing.Size(350, 60)
    $userPanel.Location = New-Object System.Drawing.Point(0, 0)
    $userPanel.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $startMenuForm.Controls.Add($userPanel)
    
    $userIcon = New-Object System.Windows.Forms.Panel
    $userIcon.Size = New-Object System.Drawing.Size(40, 40)
    $userIcon.Location = New-Object System.Drawing.Point(20, 10)
    $userIcon.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $userPanel.Controls.Add($userIcon)
    
    $userLabel = New-Object System.Windows.Forms.Label
    $userLabel.Text = $script:userInfo.Name
    $userLabel.Size = New-Object System.Drawing.Size(200, 20)
    $userLabel.Location = New-Object System.Drawing.Point(70, 10)
    $userLabel.ForeColor = [System.Drawing.Color]::White
    $userLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $userPanel.Controls.Add($userLabel)
    
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Logged in"
    $statusLabel.Size = New-Object System.Drawing.Size(200, 20)
    $statusLabel.Location = New-Object System.Drawing.Point(70, 30)
    $statusLabel.ForeColor = [System.Drawing.Color]::LightGray
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $userPanel.Controls.Add($statusLabel)
    
    # App list
    $appListView = New-Object System.Windows.Forms.ListView
    $appListView.Size = New-Object System.Drawing.Size(350, 400)
    $appListView.Location = New-Object System.Drawing.Point(0, 70)
    $appListView.View = [System.Windows.Forms.View]::List
    $appListView.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
    $appListView.ForeColor = [System.Drawing.Color]::White
    $appListView.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    
    $apps = @(
        "Settings",
        "File Explorer",
        "Mail",
        "Microsoft Edge",
        "Microsoft Store",
        "Photos",
        "Calculator",
        "Notepad",
        "Microsoft Office",
        "Control Panel",
        "Command Prompt",
        "PowerShell"
    )
    
    foreach ($app in $apps) {
        $appListView.Items.Add($app) | Out-Null
    }
    
    $appListView.Add_ItemActivate({
        $selectedApp = $appListView.SelectedItems[0].Text
        if ($selectedApp -eq "Settings") {
            Show-GPOResults
        } elseif ($selectedApp -eq "File Explorer") {
            Show-FolderRedirection
        } elseif ($selectedApp -eq "Control Panel") {
            Show-ControlPanel
        } else {
            [System.Windows.MessageBox]::Show("Opening $selectedApp", "Simulation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
        $startMenuForm.Close()
    })
    
    $startMenuForm.Controls.Add($appListView)
    
    # Power options
    $powerPanel = New-Object System.Windows.Forms.Panel
    $powerPanel.Size = New-Object System.Drawing.Size(350, 30)
    $powerPanel.Location = New-Object System.Drawing.Point(0, 470)
    $powerPanel.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $startMenuForm.Controls.Add($powerPanel)
    
    $powerButton = New-Object System.Windows.Forms.Button
    $powerButton.Size = New-Object System.Drawing.Size(70, 25)
    $powerButton.Location = New-Object System.Drawing.Point(20, 2)
    $powerButton.Text = "Power"
    $powerButton.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
    $powerButton.ForeColor = [System.Drawing.Color]::White
    $powerButton.Add_Click({
        $startMenuForm.Close()
    })
    $powerPanel.Controls.Add($powerButton)
    
    $logoutButton = New-Object System.Windows.Forms.Button
    $logoutButton.Size = New-Object System.Drawing.Size(70, 25)
    $logoutButton.Location = New-Object System.Drawing.Point(100, 2)
    $logoutButton.Text = "Log out"
    $logoutButton.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
    $logoutButton.ForeColor = [System.Drawing.Color]::White
    $logoutButton.Add_Click({
        $result = [System.Windows.MessageBox]::Show("Are you sure you want to log out?", "Log Out", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $startMenuForm.Close()
            $GLOBAL:ExitCode = 0
            [System.Windows.Forms.Application]::Exit()
        }
    })
    $powerPanel.Controls.Add($logoutButton)
    
    # Close when clicking outside the form
    $startMenuForm.Add_Deactivate({
        $startMenuForm.Close()
    })
    
    $startMenuForm.ShowDialog() | Out-Null
}

# Function to show GPO results
function Show-GPOResults {
    $gpoForm = New-Object System.Windows.Forms.Form
    $gpoForm.Text = "Applied Group Policies"
    $gpoForm.Size = New-Object System.Drawing.Size(500, 400)
    $gpoForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $gpoForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $gpoForm.MaximizeBox = $false
    $gpoForm.MinimizeBox = $false
    
    $gpoListView = New-Object System.Windows.Forms.ListView
    $gpoListView.Size = New-Object System.Drawing.Size(460, 300)
    $gpoListView.Location = New-Object System.Drawing.Point(20, 20)
    $gpoListView.View = [System.Windows.Forms.View]::Details
    $gpoListView.FullRowSelect = $true
    $gpoListView.GridLines = $true
    
    $gpoListView.Columns.Add("Policy Name", 150) | Out-Null
    $gpoListView.Columns.Add("Description", 200) | Out-Null
    $gpoListView.Columns.Add("Status", 80) | Out-Null
    
    foreach ($policy in $script:simulatedPolicies) {
        $item = New-Object System.Windows.Forms.ListViewItem($policy.Name)
        $item.SubItems.Add($policy.Description) | Out-Null
        $item.SubItems.Add($(if ($policy.Applied) { "Applied" } else { "Not Applied" })) | Out-Null
        $gpoListView.Items.Add($item) | Out-Null
    }
    
    $gpoForm.Controls.Add($gpoListView)
    
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(200, 330)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Text = "Close"
    $closeButton.Add_Click({
        $gpoForm.Close()
    })
    $gpoForm.Controls.Add($closeButton)
    
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Location = New-Object System.Drawing.Point(380, 330)
    $refreshButton.Size = New-Object System.Drawing.Size(100, 30)
    $refreshButton.Text = "Refresh"
    $refreshButton.Add_Click({
        [System.Windows.MessageBox]::Show("Group policy refresh initiated. This would trigger a gpupdate /force in a real environment.", "GPO Refresh", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    })
    $gpoForm.Controls.Add($refreshButton)
    
    $gpoForm.ShowDialog() | Out-Null
}

# Function to show folder redirection
function Show-FolderRedirection {
    $fileExplorerForm = New-Object System.Windows.Forms.Form
    $fileExplorerForm.Text = "File Explorer"
    $fileExplorerForm.Size = New-Object System.Drawing.Size(800, 500)
    $fileExplorerForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    
    # Navigation bar
    $navPanel = New-Object System.Windows.Forms.Panel
    $navPanel.Size = New-Object System.Drawing.Size(800, 30)
    $navPanel.Location = New-Object System.Drawing.Point(0, 0)
    $navPanel.BackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $fileExplorerForm.Controls.Add($navPanel)
    
    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Size = New-Object System.Drawing.Size(30, 25)
    $backButton.Location = New-Object System.Drawing.Point(5, 2)
    $backButton.Text = "←"
    $backButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $navPanel.Controls.Add($backButton)
    
    $forwardButton = New-Object System.Windows.Forms.Button
    $forwardButton.Size = New-Object System.Drawing.Size(30, 25)
    $forwardButton.Location = New-Object System.Drawing.Point(40, 2)
    $forwardButton.Text = "->"
    $forwardButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $navPanel.Controls.Add($forwardButton)
    
    $addressBar = New-Object System.Windows.Forms.TextBox
    $addressBar.Size = New-Object System.Drawing.Size(690, 25)
    $addressBar.Location = New-Object System.Drawing.Point(75, 2)
    $addressBar.Text = "This PC"
    $navPanel.Controls.Add($addressBar)
    
    # Sidebar
    $sidebarPanel = New-Object System.Windows.Forms.Panel
    $sidebarPanel.Size = New-Object System.Drawing.Size(200, 470)
    $sidebarPanel.Location = New-Object System.Drawing.Point(0, 30)
    $sidebarPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $fileExplorerForm.Controls.Add($sidebarPanel)
    
    $sidebarItems = @(
        @{ Name = "Quick access"; Y = 10 },
        @{ Name = "Desktop"; Y = 35 },
        @{ Name = "Downloads"; Y = 60 },
        @{ Name = "Documents"; Y = 85 },
        @{ Name = "Pictures"; Y = 110 },
        @{ Name = "This PC"; Y = 135 },
        @{ Name = "Network"; Y = 160 }
    )
    
    foreach ($item in $sidebarItems) {
        $sidebarLabel = New-Object System.Windows.Forms.Label
        $sidebarLabel.Text = $item.Name
        $sidebarLabel.Size = New-Object System.Drawing.Size(180, 20)
        $sidebarLabel.Location = New-Object System.Drawing.Point(10, $item.Y)
        $sidebarLabel.ForeColor = [System.Drawing.Color]::Black
        $sidebarLabel.Tag = $item.Name
        $sidebarLabel.Add_Click({
            $selectedFolder = $this.Tag
            Update-FileExplorerView -FolderName $selectedFolder
        })
        $sidebarPanel.Controls.Add($sidebarLabel)
    }
    
    # File list view
    $fileListView = New-Object System.Windows.Forms.ListView
    $fileListView.Size = New-Object System.Drawing.Size(600, 470)
    $fileListView.Location = New-Object System.Drawing.Point(200, 30)
    $fileListView.View = [System.Windows.Forms.View]::Details
    $fileListView.FullRowSelect = $true
    
    $fileListView.Columns.Add("Name", 200) | Out-Null
    $fileListView.Columns.Add("Type", 100) | Out-Null
    $fileListView.Columns.Add("Location", 250) | Out-Null
    $fileListView.Columns.Add("Size", 50) | Out-Null
    
    # Add folder data to show redirection
    foreach ($folder in $script:redirectedFolders.Keys) {
        $item = New-Object System.Windows.Forms.ListViewItem($folder)
        $item.SubItems.Add("Folder") | Out-Null
        $item.SubItems.Add($script:redirectedFolders[$folder]) | Out-Null
        $item.SubItems.Add("--") | Out-Null
        $fileListView.Items.Add($item) | Out-Null
    }
    
    $fileExplorerForm.Controls.Add($fileListView)
    
    # Function to update file explorer view based on selected folder
    function Update-FileExplorerView {
        param([string]$FolderName)
        
        $addressBar.Text = $FolderName
        $fileListView.Items.Clear()
        
        switch ($FolderName) {
            "Desktop" {
                $item1 = New-Object System.Windows.Forms.ListViewItem("README.txt")
                $item1.SubItems.Add("Text Document") | Out-Null
                $item1.SubItems.Add($script:redirectedFolders["Desktop"]) | Out-Null
                $item1.SubItems.Add("2 KB") | Out-Null
                $fileListView.Items.Add($item1) | Out-Null
                
                $item2 = New-Object System.Windows.Forms.ListViewItem("Company Portal")
                $item2.SubItems.Add("Shortcut") | Out-Null
                $item2.SubItems.Add($script:redirectedFolders["Desktop"]) | Out-Null
                $item2.SubItems.Add("1 KB") | Out-Null
                $fileListView.Items.Add($item2) | Out-Null
            }
            "Documents" {
                $item1 = New-Object System.Windows.Forms.ListViewItem("Work Files")
                $item1.SubItems.Add("Folder") | Out-Null
                $item1.SubItems.Add($script:redirectedFolders["Documents"]) | Out-Null
                $item1.SubItems.Add("--") | Out-Null
                $fileListView.Items.Add($item1) | Out-Null
                
                $item2 = New-Object System.Windows.Forms.ListViewItem("Report.docx")
                $item2.SubItems.Add("Word Document") | Out-Null
                $item2.SubItems.Add($script:redirectedFolders["Documents"]) | Out-Null
                $item2.SubItems.Add("25 KB") | Out-Null
                $fileListView.Items.Add($item2) | Out-Null
            }
            "Pictures" {
                $item1 = New-Object System.Windows.Forms.ListViewItem("Company Logo.png")
                $item1.SubItems.Add("PNG Image") | Out-Null
                $item1.SubItems.Add($script:redirectedFolders["Pictures"]) | Out-Null
                $item1.SubItems.Add("150 KB") | Out-Null
                $fileListView.Items.Add($item1) | Out-Null
            }
            "Downloads" {
                $item1 = New-Object System.Windows.Forms.ListViewItem("Setup.exe")
                $item1.SubItems.Add("Application") | Out-Null
                $item1.SubItems.Add($script:redirectedFolders["Downloads"]) | Out-Null
                $item1.SubItems.Add("1.2 MB") | Out-Null
                $fileListView.Items.Add($item1) | Out-Null
            }
            "This PC" {
                foreach ($folder in $script:redirectedFolders.Keys) {
                    $item = New-Object System.Windows.Forms.ListViewItem($folder)
                    $item.SubItems.Add("Folder") | Out-Null
                    $item.SubItems.Add($script:redirectedFolders[$folder]) | Out-Null
                    $item.SubItems.Add("--") | Out-Null
                    $fileListView.Items.Add($item) | Out-Null
                }
                
                $localItems = @(
                    @{ Name = "Local Disk (C:)"; Type = "Hard Disk Drive"; Path = "C:"; Size = "120 GB" },
                    @{ Name = "DVD Drive (D:)"; Type = "DVD Drive"; Path = "D:"; Size = "--" }
                )
                
                foreach ($localItem in $localItems) {
                    $item = New-Object System.Windows.Forms.ListViewItem($localItem.Name)
                    $item.SubItems.Add($localItem.Type) | Out-Null
                    $item.SubItems.Add($localItem.Path) | Out-Null
                    $item.SubItems.Add($localItem.Size) | Out-Null
                    $fileListView.Items.Add($item) | Out-Null
                }
            }
            "Network" {
                $item1 = New-Object System.Windows.Forms.ListViewItem("fileserver")
                $item1.SubItems.Add("Network Share") | Out-Null
                $item1.SubItems.Add("\\fileserver") | Out-Null
                $item1.SubItems.Add("--") | Out-Null
                $fileListView.Items.Add($item1) | Out-Null
            }
            default {
                # Default items for any other view
                foreach ($folder in $script:redirectedFolders.Keys) {
                    $item = New-Object System.Windows.Forms.ListViewItem($folder)
                    $item.SubItems.Add("Folder") | Out-Null
                    $item.SubItems.Add($script:redirectedFolders[$folder]) | Out-Null
                    $item.SubItems.Add("--") | Out-Null
                    $fileListView.Items.Add($item) | Out-Null
                }
            }
        }
    }
    
    # Handle double-click on items
    $fileListView.Add_DoubleClick({
        if ($fileListView.SelectedItems.Count -gt 0) {
            $selectedItem = $fileListView.SelectedItems[0].Text
            if ($script:redirectedFolders.ContainsKey($selectedItem)) {
                Update-FileExplorerView -FolderName $selectedItem
            }
        }
    })
    
    $backButton.Add_Click({
        Update-FileExplorerView -FolderName "This PC"
    })
    
    # Initial view
    Update-FileExplorerView -FolderName "This PC"
    
    $fileExplorerForm.ShowDialog() | Out-Null
}

# Function to show Control Panel
function Show-ControlPanel {
    $cpForm = New-Object System.Windows.Forms.Form
    $cpForm.Text = "Control Panel"
    $cpForm.Size = New-Object System.Drawing.Size(600, 400)
    $cpForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    
    $cpListView = New-Object System.Windows.Forms.ListView
    $cpListView.Size = New-Object System.Drawing.Size(560, 320)
    $cpListView.Location = New-Object System.Drawing.Point(20, 20)
    $cpListView.View = [System.Windows.Forms.View]::Tile
    $cpListView.TileSize = New-Object System.Drawing.Size(180, 60)
    
    $cpItems = @(
        "System and Security",
        "Network and Internet",
        "Hardware and Sound",
        "Programs",
        "User Accounts",
        "Appearance and Personalization",
        "Clock and Region",
        "Ease of Access"
    )
    
    foreach ($item in $cpItems) {
        $cpListView.Items.Add($item) | Out-Null
    }
    
    $cpListView.Add_DoubleClick({
        if ($cpListView.SelectedItems.Count -gt 0) {
            $selectedItem = $cpListView.SelectedItems[0].Text
            if ($selectedItem -eq "System and Security") {
                Show-GPOResults
            } elseif ($selectedItem -eq "User Accounts") {
                Show-PasswordChangePrompt
            } else {
                [System.Windows.MessageBox]::Show("Opening $selectedItem settings", "Control Panel", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        }
    })
    
    $cpForm.Controls.Add($cpListView)
    
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(250, 350)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Text = "Close"
    $closeButton.Add_Click({
        $cpForm.Close()
    })
    $cpForm.Controls.Add($closeButton)
    
    $cpForm.ShowDialog() | Out-Null
}

# Main execution starts here
if (-not $Username) {
    Show-LoginScreen
} else {
    $script:userInfo = Test-ADUser -Username $Username
    if ($script:userInfo) {
        Show-LoginScreen
    }
}
