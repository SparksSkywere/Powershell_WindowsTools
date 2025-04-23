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

# Function to verify user exists in AD and retrieve user properties
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
    
    try {
        # Check if UserObject is null
        if ($null -eq $UserObject) {
            Write-Error "User object is null"
            return $false
        }
        
        # Check if password is expired or must change at next logon
        if ($UserObject.PasswordExpired -or ($UserObject.PasswordNeverExpires -ne $null -and $UserObject.PasswordNeverExpires -eq $false)) {
            $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy -ErrorAction SilentlyContinue).MaxPasswordAge
            if ($null -ne $maxPasswordAge -and $maxPasswordAge -ne 0 -and $null -ne $UserObject.PasswordLastSet) {
                $passwordLastSet = $UserObject.PasswordLastSet
                $expiryDate = $passwordLastSet.Add($maxPasswordAge)
                
                # If password expires in less than 7 days, prompt for change
                if ((New-TimeSpan -Start (Get-Date) -End $expiryDate).Days -lt 7) {
                    return $true
                }
            }
        }
        
        # Check if user must change password at next logon
        if ($null -ne $UserObject.pwdLastSet -and $UserObject.pwdLastSet -eq 0) {
            return $true
        }
        
        return $false
    }
    catch {
        Write-Error "Failed to check password status: $_"
        # Default to false if there's an error
        return $false
    }
}

# Function to show login screen
function Show-LoginScreen {
    $loginForm = New-Object System.Windows.Forms.Form
    $loginForm.Text = "Windows Login"
    $loginForm.Size = New-Object System.Drawing.Size(400, 300)
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
    $signInButton.Location = New-Object System.Drawing.Point(150, 210)
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
    
    # Show the Desktop - renamed from Show-SimulatedDesktop to Show-UserDesktop to reflect real functionality
    Show-UserDesktop
}

# Function to initialize actual group policies applied to the user
function Initialize-GroupPolicies {
    try {
        $userDN = $script:userInfo.DistinguishedName
        $appliedGPOs = @()
        
        # Get GPOs applied to the user
        $gpResultOutput = & gpresult /scope user /user $script:userInfo.SamAccountName /r
        
        # Process gpresult output to get applied GPOs
        $captureGPO = $false
        foreach ($line in $gpResultOutput) {
            if ($line -match "Applied Group Policy Objects") {
                $captureGPO = $true
                continue
            }
            
            if ($captureGPO -and $line.Trim() -ne "") {
                if ($line -match "The following GPOs were not applied") {
                    $captureGPO = $false
                    continue
                }
                
                if ($line.Trim() -match "^(.+)$") {
                    $gpoName = $matches[1].Trim()
                    if ($gpoName -ne "") {
                        $appliedGPOs += @{
                            Name = $gpoName
                            Description = "Applied Group Policy"
                            Applied = $true
                        }
                    }
                }
            }
        }
        
        # If no GPOs were found using gpresult, try fallback method
        if ($appliedGPOs.Count -eq 0) {
            $userGroups = $script:userInfo.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
            foreach ($group in $userGroups) {
                try {
                    $linkedGPOs = Get-GPO -All | Where-Object { 
                        $_ | Get-GPPermission -TargetName $group -TargetType Group -ErrorAction SilentlyContinue 
                    }
                    
                    foreach ($gpo in $linkedGPOs) {
                        $appliedGPOs += @{
                            Name = $gpo.DisplayName
                            Description = "Linked via group $group"
                            Applied = $true
                        }
                    }
                }
                catch {
                    # Continue if there's an error with one group
                    continue
                }
            }
        }
        
        # If still no GPOs found, add some defaults based on domain
        if ($appliedGPOs.Count -eq 0) {
            $domain = $script:userInfo.DistinguishedName -replace "^.*?DC=(.+?)(?:,DC=.+)*$", '$1'
            $appliedGPOs += @{
                Name = "Default Domain Policy"
                Description = "Default policies for domain $domain"
                Applied = $true
            }
        }
        
        $script:simulatedPolicies = $appliedGPOs
    }
    catch {
        Write-Error "Failed to retrieve Group Policies: $_"
        # Fallback to some defaults
        $script:simulatedPolicies = @(
            @{
                Name = "Default Domain Policy"
                Description = "Default domain security settings"
                Applied = $true
            },
            @{
                Name = "Default User Settings"
                Description = "Standard user configuration"
                Applied = $true
            }
        )
    }
}

# Function to initialize folder redirection based on actual AD settings
function Initialize-FolderRedirection {
    try {
        $userHomeDirectory = $script:userInfo.HomeDirectory
        $userHomeDrive = $script:userInfo.HomeDrive
        
        $script:redirectedFolders = @{}
        
        # If home directory is set, use that
        if ($userHomeDirectory) {
            $script:redirectedFolders["Home"] = $userHomeDirectory
        }
        
        # Try to get folder redirection from Group Policy
        # This is a complex operation, so we'll do our best to detect it
        
        # Common redirection patterns based on username
        $username = $script:userInfo.SamAccountName
        $domain = $script:userInfo.DistinguishedName -replace "^.*?DC=(.+?)(?:,DC=.+)*$", '$1'
        
        # Check if any network drives are mapped for this user
        $networkDrives = @()
        try {
            $networkDrives = Get-WmiObject -Class Win32_MappedLogicalDisk | 
                             Select-Object Name, ProviderName
        }
        catch {
            Write-Warning "Could not retrieve mapped network drives: $_"
        }
        
        # Common patterns for redirected folders
        $patterns = @(
            "\\$domain\users\$username",
            "\\$domain\profiles\$username",
            "\\$domain\home\$username",
            "\\fileserver\users\$username"
        )
        
        # Add detected network drives
        foreach ($drive in $networkDrives) {
            $driveName = $drive.Name -replace ":", ""
            $script:redirectedFolders["$driveName Drive"] = $drive.ProviderName
        }
        
        # Check for common desktop redirection patterns
        if (-not $script:redirectedFolders.ContainsKey("Desktop")) {
            foreach ($pattern in $patterns) {
                $testPath = "$pattern\Desktop"
                if (Test-Path $testPath -ErrorAction SilentlyContinue) {
                    $script:redirectedFolders["Desktop"] = $testPath
                    break
                }
            }
        }
        
        # Check for common documents redirection patterns
        if (-not $script:redirectedFolders.ContainsKey("Documents")) {
            foreach ($pattern in $patterns) {
                $testPath = "$pattern\Documents"
                if (Test-Path $testPath -ErrorAction SilentlyContinue) {
                    $script:redirectedFolders["Documents"] = $testPath
                    break
                }
            }
        }
        
        # If we didn't find any redirected folders, add default ones based on home directory
        if ($script:redirectedFolders.Count -eq 0 -and $userHomeDirectory) {
            $script:redirectedFolders = @{
                "Desktop" = "$userHomeDirectory\Desktop"
                "Documents" = "$userHomeDirectory\Documents"
                "Pictures" = "$userHomeDirectory\Pictures"
                "Downloads" = "$userHomeDirectory\Downloads"
            }
        }
        elseif ($script:redirectedFolders.Count -eq 0) {
            # If we still don't have any redirection info, use default path format
            $script:redirectedFolders = @{
                "Desktop" = "\\fileserver\users\$username\Desktop"
                "Documents" = "\\fileserver\users\$username\Documents"
                "Pictures" = "\\fileserver\users\$username\Pictures"
                "Downloads" = "\\fileserver\users\$username\Downloads"
            }
        }
    }
    catch {
        Write-Error "Failed to initialize folder redirection: $_"
        # Default values if error occurs
        $username = $script:userInfo.SamAccountName
        $script:redirectedFolders = @{
            "Desktop" = "\\fileserver\users\$username\Desktop"
            "Documents" = "\\fileserver\users\$username\Documents"
            "Pictures" = "\\fileserver\users\$username\Pictures"
            "Downloads" = "\\fileserver\users\$username\Downloads"
        }
    }
}

# Function to show password change prompt with actual AD password change
function Show-PasswordChangePrompt {
    $pwdForm = New-Object System.Windows.Forms.Form
    $pwdForm.Text = "Change Password"
    $pwdForm.Size = New-Object System.Drawing.Size(450, 300)
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
    $currentPwdBox.Location = New-Object System.Drawing.Point(190, 20)
    $currentPwdBox.Size = New-Object System.Drawing.Size(220, 20)
    $pwdForm.Controls.Add($currentPwdBox)
    
    # New password
    $newPwdLabel = New-Object System.Windows.Forms.Label
    $newPwdLabel.Location = New-Object System.Drawing.Point(30, 50)
    $newPwdLabel.Size = New-Object System.Drawing.Size(150, 20)
    $newPwdLabel.Text = "New password:"
    $pwdForm.Controls.Add($newPwdLabel)
    
    $newPwdBox = New-Object System.Windows.Forms.MaskedTextBox
    $newPwdBox.PasswordChar = "*"
    $newPwdBox.Location = New-Object System.Drawing.Point(190, 50)
    $newPwdBox.Size = New-Object System.Drawing.Size(220, 20)
    $pwdForm.Controls.Add($newPwdBox)
    
    # Confirm new password
    $confirmPwdLabel = New-Object System.Windows.Forms.Label
    $confirmPwdLabel.Location = New-Object System.Drawing.Point(30, 80)
    $confirmPwdLabel.Size = New-Object System.Drawing.Size(150, 20)
    $confirmPwdLabel.Text = "Confirm new password:"
    $pwdForm.Controls.Add($confirmPwdLabel)
    
    $confirmPwdBox = New-Object System.Windows.Forms.MaskedTextBox
    $confirmPwdBox.PasswordChar = "*"
    $confirmPwdBox.Location = New-Object System.Drawing.Point(190, 80)
    $confirmPwdBox.Size = New-Object System.Drawing.Size(220, 20)
    $pwdForm.Controls.Add($confirmPwdBox)
    
    # Password requirements
    $requirementsLabel = New-Object System.Windows.Forms.Label
    $requirementsLabel.Location = New-Object System.Drawing.Point(30, 110)
    $requirementsLabel.Size = New-Object System.Drawing.Size(380, 80)
    $requirementsLabel.Text = "Password must meet the following requirements:`r`n- At least 8 characters long`r`n- Include uppercase, lowercase, numbers, and symbols`r`n- Cannot contain your username"
    $pwdForm.Controls.Add($requirementsLabel)
    
    # Buttons
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(120, 210)
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
        
        try {
            # Attempt to change the password in AD
            $username = $script:userInfo.SamAccountName
            $currentPwd = ConvertTo-SecureString -String $currentPwdBox.Text -AsPlainText -Force
            $newPwd = ConvertTo-SecureString -String $newPwdBox.Text -AsPlainText -Force
            
            Set-ADAccountPassword -Identity $username -OldPassword $currentPwd -NewPassword $newPwd -ErrorAction Stop
            
            [System.Windows.MessageBox]::Show("Password changed successfully", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            $pwdForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $pwdForm.Close()
        }
        catch [System.UnauthorizedAccessException] {
            [System.Windows.MessageBox]::Show("Current password is incorrect", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to change password: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })
    $pwdForm.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(240, 210)
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

# Function to show user desktop environment with Shell32 icons and GPO wallpaper
function Show-UserDesktop {
    $desktopForm = New-Object System.Windows.Forms.Form
    $desktopForm.Text = "Windows Desktop - $($script:userInfo.Name)"
    $desktopForm.Size = New-Object System.Drawing.Size(1024, 768)
    $desktopForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    
    # Get a wallpaper from GPO or use default Windows wallpaper
    $wallpaperPath = Get-WallpaperFromGPO
    if ($wallpaperPath -and (Test-Path $wallpaperPath)) {
        try {
            $wallpaper = [System.Drawing.Image]::FromFile($wallpaperPath)
            $desktopForm.BackgroundImage = $wallpaper
            $desktopForm.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
        }
        catch {
            Write-Warning "Failed to load wallpaper: $_"
            $desktopForm.BackColor = [System.Drawing.Color]::FromArgb(0, 103, 192)
        }
    }
    else {
        $desktopForm.BackColor = [System.Drawing.Color]::FromArgb(0, 103, 192)
    }
    
    # Desktop icons with shell32 icons instead of colored rectangles
    $iconDefinitions = @(
        @{ Name = "This PC"; IconIndex = 15; X = 20; Y = 20 },
        @{ Name = "Recycle Bin"; IconIndex = 31; X = 20; Y = 120 },
        @{ Name = "File Explorer"; IconIndex = 3; X = 20; Y = 220 },
        @{ Name = "Microsoft Edge"; IconIndex = 14; X = 20; Y = 320 },
        @{ Name = "Office"; IconIndex = 1; X = 20; Y = 420 }
    )
    
    # Extract shell32 icons
    $shell32Path = [System.Environment]::ExpandEnvironmentVariables("%SystemRoot%\System32\shell32.dll")
    
    foreach ($icon in $iconDefinitions) {
        try {
            # Try to create a PictureBox with the shell32 icon
            $iconBox = New-Object System.Windows.Forms.PictureBox
            $iconBox.Size = New-Object System.Drawing.Size(48, 48)
            $iconBox.Location = New-Object System.Drawing.Point([int]$icon.X, [int]$icon.Y)
            $iconBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
            
            # Extract the icon from shell32.dll
            $extractedIcon = [System.Drawing.Icon]::ExtractAssociatedIcon($shell32Path)
            if ($extractedIcon) {
                $iconBox.Image = $extractedIcon.ToBitmap()
            }
            else {
                # Fallback to colored rectangle
                $iconBox.BackColor = [System.Drawing.Color]::FromArgb(30, 144, 255)
            }
            
            $desktopForm.Controls.Add($iconBox)
            
            # Icon text label
            $iconLabel = New-Object System.Windows.Forms.Label
            $iconLabel.Text = $icon.Name
            $iconLabel.AutoSize = $false
            $iconLabel.Size = New-Object System.Drawing.Size(80, 20)
            $iconLabelX = [int]$icon.X - 16
            $iconLabelY = [int]$icon.Y + 50
            $iconLabel.Location = New-Object System.Drawing.Point($iconLabelX, $iconLabelY)
            $iconLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $iconLabel.ForeColor = [System.Drawing.Color]::White
            $iconLabel.BackColor = [System.Drawing.Color]::Transparent
            $desktopForm.Controls.Add($iconLabel)
        }
        catch {
            Write-Warning "Failed to create desktop icon: $_"
        }
    }
    
    # Taskbar
    $taskbar = New-Object System.Windows.Forms.Panel
    $taskbar.Size = New-Object System.Drawing.Size($desktopForm.ClientSize.Width, 40)
    $taskbarY = $desktopForm.ClientSize.Height - 40
    $taskbar.Location = New-Object System.Drawing.Point(0, $taskbarY)
    $taskbar.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $desktopForm.Controls.Add($taskbar)
    
    # Start button
    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Size = New-Object System.Drawing.Size(40, 30)
    $startButton.Location = New-Object System.Drawing.Point(10, 5)
    $startButton.Text = "Start"
    $startButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $startButton.ForeColor = [System.Drawing.Color]::White
    $startButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $startButton.Add_Click({
        Show-StartMenu
    })
    $taskbar.Controls.Add($startButton)
    
    # Time
    $timeLabel = New-Object System.Windows.Forms.Label
    $timeLabel.Size = New-Object System.Drawing.Size(70, 30)
    $timeLabelX = $taskbar.Width - 80
    $timeLabel.Location = New-Object System.Drawing.Point($timeLabelX, 5)
    $timeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $timeLabel.ForeColor = [System.Drawing.Color]::White
    $timeLabel.Text = Get-Date -Format "HH:mm"
    $taskbar.Controls.Add($timeLabel)
    
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
        $iconBtn.Location = New-Object System.Drawing.Point($posX, 5)
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
                [System.Windows.MessageBox]::Show("Opening $clickedIcon", "Action", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        })
        $taskbar.Controls.Add($iconBtn)
        $posX += 40
    }
    
    $desktopForm.ShowDialog() | Out-Null
}

# Function to get wallpaper from GPO or default sources
function Get-WallpaperFromGPO {
    try {
        # Try to get wallpaper path from registry (set by GPO)
        $wallpaperPath = Get-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name Wallpaper -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Wallpaper
        
        if ($wallpaperPath -and (Test-Path $wallpaperPath)) {
            return $wallpaperPath
        }
        
        # If no registry setting, try to get from group policies
        $gpoWallpaperPath = $null
        $gpResultOutput = & gpresult /r
        foreach ($line in $gpResultOutput) {
            if ($line -match "Desktop Wallpaper:\s+(.+)") {
                $gpoWallpaperPath = $matches[1].Trim()
                break
            }
        }
        
        if ($gpoWallpaperPath -and (Test-Path $gpoWallpaperPath)) {
            return $gpoWallpaperPath
        }
        
        # Default Windows wallpapers as fallback
        $defaultWallpapers = @(
            "$env:windir\Web\Wallpaper\Windows\img0.jpg",
            "$env:windir\Web\Wallpaper\Theme1\img1.jpg",
            "$env:windir\Web\Wallpaper\Theme2\img1.jpg"
        )
        
        foreach ($wallpaper in $defaultWallpapers) {
            if (Test-Path $wallpaper) {
                return $wallpaper
            }
        }
    }
    catch {
        Write-Warning "Failed to determine wallpaper: $_"
    }
    
    return $null
}

# Function to show Start Menu
function Show-StartMenu {
    $startMenuForm = New-Object System.Windows.Forms.Form
    $startMenuForm.Text = "Start Menu"
    $startMenuForm.Size = New-Object System.Drawing.Size(350, 500)
    $startMenuForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    
    # Calculate position to appear above start button
    $desktopScreenBounds = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $startMenuFormY = $desktopScreenBounds.Height - 510
    $startMenuForm.Location = New-Object System.Drawing.Point(10, $startMenuFormY)
    
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
            [System.Windows.MessageBox]::Show("Opening $selectedApp", "Action", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
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
    $closeButton.Location = New-Object System.Drawing.Point(140, 330)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Text = "Close"
    $closeButton.Add_Click({
        $gpoForm.Close()
    })
    $gpoForm.Controls.Add($closeButton)
    
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Location = New-Object System.Drawing.Point(260, 330)
    $refreshButton.Size = New-Object System.Drawing.Size(100, 30)
    $refreshButton.Text = "Refresh"
    $refreshButton.Add_Click({
        try {
            # Actually run gpupdate
            $result = Start-Process -FilePath "gpupdate.exe" -ArgumentList "/force" -NoNewWindow -Wait -PassThru
            if ($result.ExitCode -eq 0) {
                [System.Windows.MessageBox]::Show("Group policy refresh completed successfully.", "GPO Refresh", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                # Refresh the displayed GPOs
                Initialize-GroupPolicies
                $gpoForm.Close()
                Show-GPOResults
            } else {
                [System.Windows.MessageBox]::Show("Group policy refresh failed with exit code $($result.ExitCode)", "GPO Refresh", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            }
        } catch {
            [System.Windows.MessageBox]::Show("Failed to refresh group policies: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
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
    $addressBar.Size = New-Object System.Drawing.Size($navPanel.Width - 110, 25)
    $addressBar.Location = New-Object System.Drawing.Point(75, 2)
    $addressBar.Text = "This PC"
    $navPanel.Controls.Add($addressBar)
    
    # Sidebar
    $sidebarPanel = New-Object System.Windows.Forms.Panel
    $sidebarPanel.Size = New-Object System.Drawing.Size(200, $fileExplorerForm.ClientSize.Height - 30)
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
    $fileListView.Size = New-Object System.Drawing.Size($fileExplorerForm.ClientSize.Width - 200, $fileExplorerForm.ClientSize.Height - 30)
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
    
    # Function to update file explorer view based on real folders
    function Update-FileExplorerView {
        param([string]$FolderName)
        
        $addressBar.Text = $FolderName
        $fileListView.Items.Clear()
        
        # Get the actual path for the selected folder
        $folderPath = if ($script:redirectedFolders.ContainsKey($FolderName)) {
            $script:redirectedFolders[$FolderName]
        }
        elseif ($FolderName -eq "This PC") {
            $null  # Special case
        }
        elseif ($FolderName -eq "Network") {
            $null  # Special case
        }
        else {
            $FolderName  # Use as is
        }
        
        # Handle special folders
        switch ($FolderName) {
            "This PC" {
                # Show redirected folders
                foreach ($folder in $script:redirectedFolders.Keys) {
                    $item = New-Object System.Windows.Forms.ListViewItem($folder)
                    $item.SubItems.Add("Folder") | Out-Null
                    $item.SubItems.Add($script:redirectedFolders[$folder]) | Out-Null
                    $item.SubItems.Add("--") | Out-Null
                    $fileListView.Items.Add($item) | Out-Null
                }
                
                # Show local drives
                $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -in @(2, 3, 5) }  # Local, Network, Removable
                foreach ($drive in $drives) {
                    $driveType = switch ($drive.DriveType) {
                        2 { "Removable Disk" }
                        3 { "Local Disk" }
                        5 { "DVD Drive" }
                        default { "Drive" }
                    }
                    
                    $size = if ($drive.Size) { 
                        "{0:N1} GB" -f ($drive.Size / 1GB) 
                    } else { 
                        "--" 
                    }
                    
                    $item = New-Object System.Windows.Forms.ListViewItem("$($drive.DeviceID)")
                    $item.SubItems.Add($driveType) | Out-Null
                    $item.SubItems.Add($drive.DeviceID) | Out-Null
                    $item.SubItems.Add($size) | Out-Null
                    $fileListView.Items.Add($item) | Out-Null
                }
                return
            }
            
            "Network" {
                # Show network shares
                try {
                    $networkShares = Get-WmiObject -Class Win32_Share | Where-Object { $_.Path -ne "" }
                    foreach ($share in $networkShares) {
                        $item = New-Object System.Windows.Forms.ListViewItem($share.Name)
                        $item.SubItems.Add("Network Share") | Out-Null
                        $item.SubItems.Add("\\$(hostname)\$($share.Name)") | Out-Null
                        $item.SubItems.Add("--") | Out-Null
                        $fileListView.Items.Add($item) | Out-Null
                    }
                }
                catch {
                    $item = New-Object System.Windows.Forms.ListViewItem("Error listing network shares: $($_.Exception.Message)")
                    $fileListView.Items.Add($item) | Out-Null
                }
                return
            }
        }
        
        # For regular folders, try to show actual content
        if ($folderPath -and (Test-Path $folderPath -ErrorAction SilentlyContinue)) {
            try {
                $items = Get-ChildItem -Path $folderPath -ErrorAction Stop
                foreach ($fileItem in $items) {
                    $type = if ($fileItem.PSIsContainer) { "Folder" } else { $fileItem.Extension -replace '^\.' }
                    $size = if ($fileItem.PSIsContainer) { "--" } else { "{0:N0} KB" -f ($fileItem.Length / 1KB) }
                    
                    $item = New-Object System.Windows.Forms.ListViewItem($fileItem.Name)
                    $item.SubItems.Add($type) | Out-Null
                    $item.SubItems.Add($fileItem.FullName) | Out-Null
                    $item.SubItems.Add($size) | Out-Null
                    $fileListView.Items.Add($item) | Out-Null
                }
            }
            catch {
                $item = New-Object System.Windows.Forms.ListViewItem("Error: $($_.Exception.Message)")
                $item.SubItems.Add("Error") | Out-Null
                $item.SubItems.Add($folderPath) | Out-Null
                $item.SubItems.Add("--") | Out-Null
                $fileListView.Items.Add($item) | Out-Null
            }
        }
        else {
            # If path doesn't exist or couldn't be accessed
            $item = New-Object System.Windows.Forms.ListViewItem("This folder is not accessible")
            $item.SubItems.Add("Error") | Out-Null
            $item.SubItems.Add($folderPath) | Out-Null
            $item.SubItems.Add("--") | Out-Null
            $fileListView.Items.Add($item) | Out-Null
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
    $cpForm.Size = New-Object System.Drawing.Size(600, 450)
    $cpForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    
    $cpListView = New-Object System.Windows.Forms.ListView
    $cpListView.Size = New-Object System.Drawing.Size(560, 350)
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
                try {
                    # Try to open actual control panel applets
                    switch ($selectedItem) {
                        "Network and Internet" { Start-Process "control.exe" "ncpa.cpl" }
                        "Hardware and Sound" { Start-Process "control.exe" "mmsys.cpl" }
                        "Programs" { Start-Process "appwiz.cpl" }
                        "Appearance and Personalization" { Start-Process "control.exe" "desk.cpl" }
                        "Clock and Region" { Start-Process "control.exe" "timedate.cpl" }
                        "Ease of Access" { Start-Process "control.exe" "access.cpl" }
                        default { [System.Windows.MessageBox]::Show("Opening $selectedItem settings", "Control Panel", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) }
                    }
                }
                catch {
                    [System.Windows.MessageBox]::Show(("Failed to open {0}: {1}" -f $selectedItem, $_.Exception.Message), "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            }
        }
    })
    
    $cpForm.Controls.Add($cpListView)
    
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(250, 380)
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
