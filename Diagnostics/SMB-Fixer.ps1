# Requires elevation for fixing SMB issues
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script requires administrator privileges. Please run PowerShell as administrator."
    exit
}

# Initialize logging
$logPath = Join-Path $PSScriptRoot "SMB-Fixer.log"
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $logPath -Value $logMessage
    Write-Host $logMessage
}

function Fix-SMBv1 {
    Write-Log "Checking SMBv1 status..." "INFO"
    try {
        $smbv1 = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol
        if ($smbv1.State -eq "Enabled") {
            Write-Log "Disabling SMBv1 protocol for security..." "WARNING"
            Disable-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -NoRestart
            Write-Log "SMBv1 has been disabled" "SUCCESS"
            return $true
        } else {
            Write-Log "SMBv1 is already disabled" "SUCCESS"
            return $false
        }
    } catch {
        Write-Log "Error fixing SMBv1: $_" "ERROR"
        return $false
    }
}

function Fix-SMBSecurity {
    Write-Log "Configuring SMB security settings..." "INFO"
    try {
        # Enable SMB signing and encryption
        Set-SmbServerConfiguration -EnableSecuritySignature $true -RequireSecuritySignature $true -EnableSMB2Protocol $true -Confirm:$false
        Set-SmbClientConfiguration -RequireSecuritySignature $true -EnableSecuritySignature $true -Confirm:$false
        
        # Enable SMB encryption where supported
        Set-SmbServerConfiguration -EncryptData $true -Confirm:$false
        Write-Log "SMB security settings have been configured" "SUCCESS"
        return $true
    } catch {
        Write-Log "Error configuring SMB security: $_" "ERROR"
        return $false
    }
}

function Fix-SMBServices {
    Write-Log "Checking SMB-related services..." "INFO"
    $services = @(
        "LanmanServer",      # Server service
        "LanmanWorkstation", # Workstation service
        "Browser",           # Computer Browser service
        "MRxSmb"            # SMB Redirector
    )
    
    $fixedServices = 0
    foreach ($service in $services) {
        try {
            $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
            if ($svc) {
                if ($svc.Status -ne "Running") {
                    Set-Service -Name $service -StartupType Automatic
                    Start-Service -Name $service
                    Write-Log "Service $service has been started and set to Automatic" "SUCCESS"
                    $fixedServices++
                } else {
                    Write-Log "Service $service is already running" "INFO"
                }
            }
        } catch {
            Write-Log "Error fixing service $service $($_.Exception.Message)" "ERROR"
        }
    }
    return $fixedServices -gt 0
}

function Fix-SMBFirewall {
    Write-Log "Configuring SMB firewall rules..." "INFO"
    try {
        # Enable File and Printer Sharing rules
        $rules = Get-NetFirewallRule -DisplayGroup "File and Printer Sharing" -ErrorAction Stop
        foreach ($rule in $rules) {
            if (-not $rule.Enabled) {
                Enable-NetFirewallRule -Name $rule.Name
                Write-Log "Enabled firewall rule: $($rule.DisplayName)" "SUCCESS"
            }
        }
        
        # Enable SMB-specific rules
        $smbRules = Get-NetFirewallRule -DisplayName "*SMB*" -ErrorAction SilentlyContinue
        foreach ($rule in $smbRules) {
            if (-not $rule.Enabled) {
                Enable-NetFirewallRule -Name $rule.Name
                Write-Log "Enabled SMB firewall rule: $($rule.DisplayName)" "SUCCESS"
            }
        }
        return $true
    } catch {
        Write-Log "Error configuring firewall rules: $_" "ERROR"
        return $false
    }
}

function Reset-SMBServer {
    Write-Log "Resetting SMB Server configuration..." "INFO"
    try {
        # Stop SMB services
        Stop-Service -Name "LanmanServer" -Force
        Stop-Service -Name "LanmanWorkstation" -Force
        
        # Reset SMB configurations to default secure settings
        Set-SmbServerConfiguration -EnableSMB2Protocol $true `
                                 -EnableSMB1Protocol $false `
                                 -RequireSecuritySignature $true `
                                 -EnableSecuritySignature $true `
                                 -EncryptData $true `
                                 -Confirm:$false
        
        # Start services again
        Start-Service -Name "LanmanWorkstation"
        Start-Service -Name "LanmanServer"
        
        Write-Log "SMB Server has been reset to secure defaults" "SUCCESS"
        return $true
    } catch {
        Write-Log "Error resetting SMB Server: $_" "ERROR"
        return $false
    }
}

function Fix-SMBSharePermissions {
    param (
        [string]$ShareName
    )
    Write-Log "Checking permissions for share: $ShareName" "INFO"
    try {
        $share = Get-SmbShare -Name $ShareName -ErrorAction Stop
        $perms = Get-SmbShareAccess -Name $ShareName
        
        # Check for overly permissive settings
        $unsecurePerms = $perms | Where-Object { $_.AccessRight -eq 'Full' -and $_.AccountName -eq 'Everyone' }
        if ($unsecurePerms) {
            Write-Log "Fixing overly permissive settings on $ShareName" "WARNING"
            Revoke-SmbShareAccess -Name $ShareName -AccountName "Everyone" -Force
            Grant-SmbShareAccess -Name $ShareName -AccountName "Authenticated Users" -AccessRight Read -Force
            Write-Log "Share permissions have been secured for $ShareName" "SUCCESS"
            return $true
        } else {
            Write-Log "Share permissions for $ShareName are already secure" "INFO"
            return $false
        }
    } catch {
        Write-Log "Error fixing share permissions for $ShareName ${_}" "ERROR"
        return $false
    }
}

function Fix-SMBNetworkSettings {
    Write-Log "Checking and fixing network settings for SMB..." "INFO"
    try {
        # Optimize TCP settings
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name "EnableConnectionRateLimit" -Value 0
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters" -Name "DisableBandwidthThrottling" -Value 1
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters" -Name "FileInfoCacheLifetime" -Value 0
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters" -Name "DirectoryCacheLifetime" -Value 0
        
        # Optimize network adapters
        Get-NetAdapter | Where-Object {$_.Status -eq "Up"} | ForEach-Object {
            Set-NetAdapterAdvancedProperty -Name $_.Name -RegistryKeyword "*JumboPacket" -RegistryValue 9014 -ErrorAction SilentlyContinue
            Disable-NetAdapterPowerManagement -Name $_.Name -ErrorAction SilentlyContinue
        }
        
        Write-Log "Network settings optimized for SMB" "SUCCESS"
        return $true
    } catch {
        Write-Log "Error fixing network settings: $_" "ERROR"
        return $false
    }
}

function Clear-SMBCache {
    Write-Log "Clearing SMB cache..." "INFO"
    try {
        # Stop services
        Stop-Service -Name "LanmanWorkstation" -Force
        Stop-Service -Name "LanmanServer" -Force
        
        # Clear SMB client cache
        Remove-Item -Path "$env:SystemRoot\System32\config\systemprofile\AppData\Local\Microsoft\Windows\SMB*" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "$env:SystemRoot\System32\config\systemprofile\AppData\Local\Microsoft\Windows\WebClient\Cache" -Recurse -Force -ErrorAction SilentlyContinue
        
        # Restart services
        Start-Service -Name "LanmanWorkstation"
        Start-Service -Name "LanmanServer"
        
        Write-Log "SMB cache cleared successfully" "SUCCESS"
        return $true
    } catch {
        Write-Log "Error clearing SMB cache: $_" "ERROR"
        return $false
    }
}

function Fix-SMBRegistry {
    Write-Log "Applying SMB registry optimizations..." "INFO"
    try {
        # SMB client optimizations
        $clientPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters"
        Set-ItemProperty -Path $clientPath -Name "MaxCmds" -Value 2048
        Set-ItemProperty -Path $clientPath -Name "MaxThreads" -Value 64
        Set-ItemProperty -Path $clientPath -Name "MaxCollectionCount" -Value 64
        
        # SMB server optimizations
        $serverPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
        Set-ItemProperty -Path $serverPath -Name "IRPStackSize" -Value 32
        Set-ItemProperty -Path $serverPath -Name "SizReqBuf" -Value 17424
        Set-ItemProperty -Path $serverPath -Name "Size" -Value 3
        
        Write-Log "Registry optimizations applied" "SUCCESS"
        return $true
    } catch {
        Write-Log "Error applying registry optimizations: $_" "ERROR"
        return $false
    }
}

function Test-SMBPorts {
    Write-Log "Testing SMB ports..." "INFO"
    try {
        $ports = @(445, 139)
        $results = @()
        
        foreach ($port in $ports) {
            $test = Test-NetConnection -ComputerName localhost -Port $port -WarningAction SilentlyContinue
            if (-not $test.TcpTestSucceeded) {
                Write-Log "Port $port is not accessible" "WARNING"
                # Enable port in firewall if needed
                Enable-NetFirewallRule -DisplayName "*SMB*" -Direction Inbound -LocalPort $port -Protocol TCP
            } else {
                Write-Log "Port $port is accessible" "SUCCESS"
            }
        }
        return $true
    } catch {
        Write-Log "Error testing SMB ports: $_" "ERROR"
        return $false
    }
}

function Show-FixerMenu {
    $continue = $true
    while ($continue) {
        Clear-Host
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host "              SMB FIXER TOOL                   " -ForegroundColor Cyan
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host " 1. Fix SMBv1 (Disable)" -ForegroundColor Green
        Write-Host " 2. Configure SMB Security Settings" -ForegroundColor Green
        Write-Host " 3. Fix SMB Services" -ForegroundColor Green
        Write-Host " 4. Configure Firewall Rules" -ForegroundColor Green
        Write-Host " 5. Reset SMB Server to Secure Defaults" -ForegroundColor Yellow
        Write-Host " 6. Fix Share Permissions" -ForegroundColor Yellow
        Write-Host " 7. Fix Network Settings" -ForegroundColor Yellow
        Write-Host " 8. Clear SMB Cache" -ForegroundColor Yellow
        Write-Host " 9. Apply Registry Optimizations" -ForegroundColor Yellow
        Write-Host "10. Test SMB Ports" -ForegroundColor Yellow
        Write-Host "11. Fix All Issues" -ForegroundColor Red
        Write-Host "12. View Log File" -ForegroundColor Magenta
        Write-Host ""
        Write-Host " Q. Quit" -ForegroundColor Red
        Write-Host ""
        
        $choice = Read-Host "Enter your choice"
        
        switch ($choice) {
            "1" {
                Fix-SMBv1
                Read-Host "Press Enter to continue"
            }
            "2" {
                Fix-SMBSecurity
                Read-Host "Press Enter to continue"
            }
            "3" {
                Fix-SMBServices
                Read-Host "Press Enter to continue"
            }
            "4" {
                Fix-SMBFirewall
                Read-Host "Press Enter to continue"
            }
            "5" {
                Reset-SMBServer
                Read-Host "Press Enter to continue"
            }
            "6" {
                $shareName = Read-Host "Enter share name to fix"
                Fix-SMBSharePermissions -ShareName $shareName
                Read-Host "Press Enter to continue"
            }
            "7" {
                Fix-SMBNetworkSettings
                Read-Host "Press Enter to continue"
            }
            "8" {
                Clear-SMBCache
                Read-Host "Press Enter to continue"
            }
            "9" {
                Fix-SMBRegistry
                Read-Host "Press Enter to continue"
            }
            "10" {
                Test-SMBPorts
                Read-Host "Press Enter to continue"
            }
            "11" {
                Write-Log "Starting complete SMB fix..." "INFO"
                Fix-SMBv1
                Fix-SMBSecurity
                Fix-SMBServices
                Fix-SMBFirewall
                Reset-SMBServer
                Fix-SMBNetworkSettings
                Clear-SMBCache
                Fix-SMBRegistry
                Test-SMBPorts
                
                # Fix all shares
                Get-SmbShare | Where-Object { -not $_.Special } | ForEach-Object {
                    Fix-SMBSharePermissions -ShareName $_.Name
                }
                
                Write-Log "Complete SMB fix finished" "SUCCESS"
                Read-Host "Press Enter to continue"
            }
            "12" {
                if (Test-Path $logPath) {
                    Get-Content $logPath | More
                } else {
                    Write-Host "No log file found." -ForegroundColor Red
                }
                Read-Host "Press Enter to continue"
            }
            "Q" { $continue = $false }
            "q" { $continue = $false }
            default {
                Write-Host "Invalid choice. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    }
}

# Start the fixer tool
Write-Log "SMB Fixer started" "INFO"
Show-FixerMenu
Write-Log "SMB Fixer completed" "INFO"
