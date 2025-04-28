# Requires elevation for certain diagnostics functions
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Some functions require administrator privileges. Run PowerShell as administrator for full functionality."
}

function Write-ColorOutput {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$ForegroundColor = "White"
    )
    
    $originalColor = $host.UI.RawUI.ForegroundColor
    $host.UI.RawUI.ForegroundColor = $ForegroundColor
    Write-Output $Message
    $host.UI.RawUI.ForegroundColor = $originalColor
}

function Write-StatusMessage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Status = "Info",
        
        [Parameter(Mandatory = $false)]
        [switch]$NoNewline
    )
    
    $statusColors = @{
        "Success" = "Green"
        "Error"   = "Red"
        "Warning" = "Yellow"
        "Info"    = "Cyan"
        "Note"    = "White"
    }
    
    $statusPrefix = switch ($Status) {
        "Success" { "[+] " }
        "Error"   { "[!] " }
        "Warning" { "[*] " }
        "Info"    { "[i] " }
        "Note"    { "[-] " }
        default   { "" }
    }
    
    $color = $statusColors[$Status]
    if (-not $color) { $color = "White" }
    
    if ($NoNewline) {
        Write-Host "$statusPrefix$Message" -ForegroundColor $color -NoNewline
    } else {
        Write-Host "$statusPrefix$Message" -ForegroundColor $color
    }
}

function Test-SMBConfiguration {
    Write-StatusMessage "Checking SMB Configuration..." -Status "Info"
    
    try {
        # Check SMB versions enabled
        $smbConfig = Get-SmbServerConfiguration -ErrorAction Stop
        Write-StatusMessage "SMB Server is present on this system." -Status "Success"
        
        Write-StatusMessage "SMB Protocol Versions:" -Status "Info"
        Write-StatusMessage "SMBv1 Enabled: $($smbConfig.EnableSMB1Protocol)" -Status $(if ($smbConfig.EnableSMB1Protocol) { "Warning" } else { "Success" })
        if ($smbConfig.EnableSMB1Protocol) {
            Write-StatusMessage "WARNING: SMBv1 is enabled. This version has known security vulnerabilities." -Status "Warning"
        }
        
        Write-StatusMessage "SMBv2/v3 Enabled: $($smbConfig.EnableSMB2Protocol)" -Status $(if ($smbConfig.EnableSMB2Protocol) { "Success" } else { "Warning" })
        
        # Check SMB signing
        Write-StatusMessage "SMB Signing Required: $($smbConfig.RequireSecuritySignature)" -Status "Info"
        Write-StatusMessage "SMB Encryption Required: $($smbConfig.EncryptData)" -Status "Info"
        
        # Check SMB multichannel
        Write-StatusMessage "SMB Multichannel Enabled: $($smbConfig.EnableMultiChannel)" -Status "Info"
        
        # Check SMB Direct (RDMA)
        if (Get-Command Get-SmbClientNetworkInterface -ErrorAction SilentlyContinue) {
            $rdmaInterfaces = Get-SmbClientNetworkInterface | Where-Object { $_.RdmaCapable -eq $true }
            if ($rdmaInterfaces) {
                Write-StatusMessage "RDMA-capable interfaces found for SMB Direct:" -Status "Success"
                $rdmaInterfaces | ForEach-Object {
                    Write-StatusMessage "  - $($_.InterfaceIndex): $($_.InterfaceDescription)" -Status "Note"
                }
            } else {
                Write-StatusMessage "No RDMA-capable interfaces found for SMB Direct." -Status "Info"
            }
        }
        
        # Check SMB client configuration
        $clientConfig = Get-SmbClientConfiguration -ErrorAction Stop
        Write-StatusMessage "SMB Client Configuration:" -Status "Info"
        Write-StatusMessage "  - Signing Enabled: $($clientConfig.RequireSecuritySignature)" -Status "Info"
        Write-StatusMessage "  - Encryption Enabled: $($clientConfig.EnableSecuritySignature)" -Status "Info"
    }
    catch {
        Write-StatusMessage "Failed to retrieve SMB configuration: $_" -Status "Error"
    }
    
    # Check for SMB-related services
    $smbServices = @(
        "LanmanServer",       # Server service
        "LanmanWorkstation",  # Workstation service
        "Browser",            # Computer Browser service
        "MRxSmb"              # SMB Redirector
    )
    
    Write-StatusMessage "Checking SMB-related services:" -Status "Info"
    foreach ($service in $smbServices) {
        try {
            $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
            if ($svc) {
                $statusColor = if ($svc.Status -eq "Running") { "Success" } else { "Warning" }
                Write-StatusMessage "  - $($svc.DisplayName): $($svc.Status)" -Status $statusColor
            } else {
                Write-StatusMessage "  - $service Not Found" -Status "Info"
            }
        }
        catch {
            Write-StatusMessage "  - Error checking service $service" -Status "Error"
        }
    }
}

function Get-SMBShareInfo {
    Write-StatusMessage "Enumerating SMB Shares..." -Status "Info"
    
    try {
        $shares = Get-SmbShare -ErrorAction Stop
        
        if ($shares.Count -eq 0) {
            Write-StatusMessage "No SMB shares found on this system." -Status "Warning"
            return
        }
        
        Write-StatusMessage "Found $($shares.Count) SMB shares:" -Status "Success"
        
        foreach ($share in $shares) {
            $shareType = if ($share.Special) { "Special" } else { "User-created" }
            $sharePath = $share.Path
            
            Write-Host
            Write-StatusMessage "Share: $($share.Name)" -Status "Info"
            Write-StatusMessage "  - Path: $sharePath" -Status "Note"
            Write-StatusMessage "  - Description: $($share.Description)" -Status "Note"
            Write-StatusMessage "  - Type: $shareType" -Status "Note"
            
            # Check share permissions
            try {
                $perms = Get-SmbShareAccess -Name $share.Name -ErrorAction Stop
                Write-StatusMessage "  - Permissions:" -Status "Info"
                foreach ($perm in $perms) {
                    Write-StatusMessage "    * $($perm.AccountName): $($perm.AccessRight)" -Status "Note"
                }
            }
            catch {
                Write-StatusMessage "  - Could not retrieve permissions: $_" -Status "Error"
            }
            
            # Check if path exists and is accessible
            if (-not [string]::IsNullOrEmpty($sharePath)) {
                if (Test-Path -Path $sharePath -ErrorAction SilentlyContinue) {
                    Write-StatusMessage "  - Path exists and is accessible." -Status "Success"
                    
                    try {
                        $acl = Get-Acl -Path $sharePath -ErrorAction Stop
                        Write-StatusMessage "  - NTFS Permissions (first 3):" -Status "Info"
                        $acl.Access | Select-Object -First 3 | ForEach-Object {
                            Write-StatusMessage "    * $($_.IdentityReference): $($_.FileSystemRights)" -Status "Note"
                        }
                        
                        if ($acl.Access.Count -gt 3) {
                            Write-StatusMessage "    * ... and $($acl.Access.Count - 3) more entries" -Status "Note"
                        }
                    }
                    catch {
                        Write-StatusMessage "  - Could not retrieve NTFS permissions: $_" -Status "Error"
                    }
                }
                else {
                    Write-StatusMessage "  - Share path does not exist or is not accessible." -Status "Error"
                }
            }
        }
    }
    catch {
        Write-StatusMessage "Error enumerating SMB shares: $_" -Status "Error"
    }
}

function Test-RemoteSMBShare {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [string]$ShareName
    )
    
    if ([string]::IsNullOrEmpty($ShareName)) {
        Write-StatusMessage "Testing SMB connectivity to $ComputerName..." -Status "Info"
    } else {
        Write-StatusMessage "Testing SMB connectivity to \\$ComputerName\$ShareName..." -Status "Info"
    }
    
    # Test basic connectivity first
    $pingResult = Test-Connection -ComputerName $ComputerName -Count 2 -Quiet
    if (-not $pingResult) {
        Write-StatusMessage "Cannot ping host $ComputerName. Check network connectivity." -Status "Error"
        return $false
    } else {
        Write-StatusMessage "Host $ComputerName responds to ping." -Status "Success"
    }
    
    # Check if SMB ports are open
    $tcpPorts = @(445, 139)
    foreach ($port in $tcpPorts) {
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $connection = $tcp.BeginConnect($ComputerName, $port, $null, $null)
            $wait = $connection.AsyncWaitHandle.WaitOne(1000, $false)
            
            if ($wait) {
                try {
                    $tcp.EndConnect($connection)
                    Write-StatusMessage "SMB port $port is open on $ComputerName." -Status "Success"
                    $tcp.Close()
                    break
                }
                catch {
                    Write-StatusMessage "Could not complete connection to port $port $_" -Status "Error"
                }
            } else {
                Write-StatusMessage "SMB port $port is closed or filtered on $ComputerName." -Status "Error"
            }
            $tcp.Close()
        }
        catch {
            Write-StatusMessage "Error testing port $port $_" -Status "Error"
        }
    }
    
    # Try to enumerate shares or test specific share
    try {
        if ([string]::IsNullOrEmpty($ShareName)) {
            Write-StatusMessage "Attempting to enumerate shares on $ComputerName..." -Status "Info"
            $shares = net view "\\$ComputerName" 2>&1
            
            if ($LASTEXITCODE -eq 0) {
                Write-StatusMessage "Successfully enumerated shares on $ComputerName." -Status "Success"
                return $true
            } else {
                Write-StatusMessage "Failed to enumerate shares on $ComputerName. Error code: $LASTEXITCODE" -Status "Error"
                Write-StatusMessage "Response: $shares" -Status "Error"
                return $false
            }
        } else {
            Write-StatusMessage "Testing access to \\$ComputerName\$ShareName..." -Status "Info"
            
            $testPath = "\\$ComputerName\$ShareName"
            if (Test-Path -Path $testPath -ErrorAction Stop) {
                Write-StatusMessage "Successfully accessed \\$ComputerName\$ShareName" -Status "Success"
                
                # Try to list contents
                try {
                    $items = Get-ChildItem -Path $testPath -ErrorAction Stop -Force | Select-Object -First 1
                    Write-StatusMessage "Can list contents of the share." -Status "Success"
                    return $true
                }
                catch {
                    Write-StatusMessage "Could access share but couldn't list contents: $_" -Status "Warning"
                    return $false
                }
            } else {
                Write-StatusMessage "Could not access \\$ComputerName\$ShareName" -Status "Error"
                return $false
            }
        }
    }
    catch {
        Write-StatusMessage "Error testing SMB share: $_" -Status "Error"
        return $false
    }
}

function Test-SMBNetworking {
    Write-StatusMessage "Checking network configuration for SMB..." -Status "Info"
    
    # Check if SMB ports are open in the firewall
    Write-StatusMessage "Checking Windows Firewall settings for SMB..." -Status "Info"
    
    try {
        $firewallRules = Get-NetFirewallRule -DisplayGroup "File and Printer Sharing" -ErrorAction Stop
        
        $enabledRules = $firewallRules | Where-Object { $_.Enabled -eq $true }
        $totalRules = $firewallRules.Count
        $enabledCount = $enabledRules.Count
        
        Write-StatusMessage "File and Printer Sharing firewall rules: $enabledCount of $totalRules enabled" -Status $(if ($enabledCount -gt 0) { "Success" } else { "Warning" })
        
        foreach ($rule in $firewallRules) {
            $status = if ($rule.Enabled) { "Enabled" } else { "Disabled" }
            $statusColor = if ($rule.Enabled) { "Success" } else { "Warning" }
            Write-StatusMessage "  - $($rule.DisplayName): $status" -Status $statusColor
        }
        
        # Check SMB-specific rules
        $smbRules = Get-NetFirewallRule -DisplayName "*SMB*" -ErrorAction SilentlyContinue
        if ($smbRules) {
            Write-StatusMessage "SMB-specific firewall rules:" -Status "Info"
            foreach ($rule in $smbRules) {
                $status = if ($rule.Enabled) { "Enabled" } else { "Disabled" }
                $statusColor = if ($rule.Enabled) { "Success" } else { "Warning" }
                Write-StatusMessage "  - $($rule.DisplayName): $status" -Status $statusColor
            }
        }
    }
    catch {
        Write-StatusMessage "Failed to retrieve firewall rules: $_" -Status "Error"
    }
    
    # Check network interfaces
    Write-StatusMessage "Checking network interfaces..." -Status "Info"
    try {
        $interfaces = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
        
        if ($interfaces.Count -eq 0) {
            Write-StatusMessage "No active network interfaces found!" -Status "Error"
        } else {
            Write-StatusMessage "Active network interfaces:" -Status "Success"
            foreach ($iface in $interfaces) {
                Write-StatusMessage "  - $($iface.Name): $($iface.InterfaceDescription)" -Status "Note"
                Write-StatusMessage "    Status: $($iface.Status), Speed: $($iface.LinkSpeed)" -Status "Note"
                
                # Get IP configuration for this interface
                $ipConfig = Get-NetIPAddress -InterfaceIndex $iface.ifIndex -ErrorAction SilentlyContinue
                foreach ($ip in $ipConfig) {
                    if ($ip.AddressFamily -eq "IPv4") {
                        Write-StatusMessage "    IPv4: $($ip.IPAddress)/$($ip.PrefixLength)" -Status "Note"
                    }
                }
            }
        }
    }
    catch {
        Write-StatusMessage "Failed to retrieve network interface information: $_" -Status "Error"
    }
    
    # Check for DNS resolution issues
    Write-StatusMessage "Checking DNS resolution..." -Status "Info"
    try {
        $computerName = $env:COMPUTERNAME
        $dnsResult = Resolve-DnsName -Name $computerName -ErrorAction SilentlyContinue
        
        if ($dnsResult) {
            Write-StatusMessage "DNS resolution for $computerName successful" -Status "Success"
            Write-StatusMessage "  - Resolved to: $($dnsResult.IPAddress)" -Status "Note"
        } else {
            Write-StatusMessage "Failed to resolve own hostname via DNS" -Status "Warning"
        }
    }
    catch {
        Write-StatusMessage "DNS resolution error: $_" -Status "Error"
    }
}

function Measure-SMBPerformance {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,
        
        [Parameter(Mandatory = $false)]
        [int]$FileSizeMB = 100,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipWrite
    )
    
    if (-not (Test-Path -Path $Path -ErrorAction SilentlyContinue)) {
        Write-StatusMessage "The specified path does not exist: $Path" -Status "Error"
        return
    }
    
    $testFile = Join-Path -Path $Path -ChildPath "SMB_Performance_Test_$(Get-Random).dat"
    
    Write-StatusMessage "Testing SMB performance with $FileSizeMB MB file..." -Status "Info"
    
    # Write test
    if (-not $SkipWrite) {
        try {
            Write-StatusMessage "Testing write performance..." -Status "Info" -NoNewline
            $writeTimer = [System.Diagnostics.Stopwatch]::StartNew()
            
            $buffer = New-Object byte[] (1MB)
            (New-Object Random).NextBytes($buffer)
            
            $stream = [System.IO.File]::OpenWrite($testFile)
            for ($i = 0; $i -lt $FileSizeMB; $i++) {
                $stream.Write($buffer, 0, $buffer.Length)
                
                # Show progress every 10 MB
                if ($i % 10 -eq 0) {
                    Write-Host "." -NoNewline
                }
            }
            $stream.Close()
            $writeTimer.Stop()
            Write-Host ""
            
            $writeSpeed = [math]::Round($FileSizeMB / $writeTimer.Elapsed.TotalSeconds, 2)
            Write-StatusMessage "Write complete: $writeSpeed MB/s" -Status "Success"
        }
        catch {
            Write-StatusMessage "Write test failed: $_" -Status "Error"
            return
        }
    }
    
    # Read test
    if (Test-Path -Path $testFile -ErrorAction SilentlyContinue) {
        try {
            Write-StatusMessage "Testing read performance..." -Status "Info" -NoNewline
            $readTimer = [System.Diagnostics.Stopwatch]::StartNew()
            
            $buffer = New-Object byte[] (1MB)
            $stream = [System.IO.File]::OpenRead($testFile)
            
            $totalRead = 0
            $bytesRead = 0
            
            do {
                $bytesRead = $stream.Read($buffer, 0, $buffer.Length)
                $totalRead += $bytesRead
                
                # Show progress every 10 MB
                if ($totalRead % (10MB) -lt $buffer.Length) {
                    Write-Host "." -NoNewline
                }
                
            } while ($bytesRead -gt 0)
            
            $stream.Close()
            $readTimer.Stop()
            Write-Host ""
            
            $readSpeed = [math]::Round(($totalRead / 1MB) / $readTimer.Elapsed.TotalSeconds, 2)
            Write-StatusMessage "Read complete: $readSpeed MB/s" -Status "Success"
        }
        catch {
            Write-StatusMessage "Read test failed: $_" -Status "Error"
        }
        
        # Clean up test file
        try {
            Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
            Write-StatusMessage "Test file removed." -Status "Info"
        }
        catch {
            Write-StatusMessage "Could not remove test file: $_" -Status "Warning"
        }
    }
    else {
        Write-StatusMessage "Test file not found. Write test may have failed." -Status "Error"
    }
    
    # Quick concurrent connection test
    try {
        Write-StatusMessage "Testing SMB connection handling..." -Status "Info"
        
        $concurrentTestPath = Split-Path -Path $Path -Parent
        $maxConnections = 10
        $connections = @()
        
        for ($i = 1; $i -le $maxConnections; $i++) {
            Write-Host "." -NoNewline
            
            try {
                $newConn = New-PSDrive -Name "TestSMB$i" -PSProvider FileSystem -Root $concurrentTestPath -ErrorAction Stop
                $connections += "TestSMB$i"
            }
            catch {
                Write-StatusMessage "Failed after $($i-1) connections: $_" -Status "Warning"
                break
            }
        }
        
        Write-Host ""
        Write-StatusMessage "Successfully established $($connections.Count) concurrent connections." -Status "Success"
        
        # Clean up the drives
        foreach ($conn in $connections) {
            Remove-PSDrive -Name $conn -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-StatusMessage "Connection test failed: $_" -Status "Error"
    }
}

function Test-WindowsSearchIndexer {
    Write-StatusMessage "Checking Windows Search and Indexer..." -Status "Info"
    
    # Check if running on Windows Server
    $isServer = (Get-CimInstance -ClassName Win32_OperatingSystem).ProductType -ne 1
    
    if ($isServer) {
        Write-StatusMessage "Running on Windows Server. Checking Search service components..." -Status "Info"
        
        # Check if Windows Search Service is installed
        try {
            $searchFeature = Get-WindowsFeature -Name "Search-Service" -ErrorAction SilentlyContinue
            if ($searchFeature -and $searchFeature.Installed) {
                Write-StatusMessage "Windows Search Service feature is installed." -Status "Success"
            } else {
                Write-StatusMessage "Windows Search Service feature is not installed on this server." -Status "Warning"
                Write-StatusMessage "Use 'Install-WindowsFeature -Name Search-Service' to install it if needed." -Status "Info"
            }
        }
        catch {
            Write-StatusMessage "Could not check Windows Search feature status: $_" -Status "Error"
        }
        
        # Check File Server role
        try {
            $fileServerRole = Get-WindowsFeature -Name "FS-FileServer" -ErrorAction SilentlyContinue
            if ($fileServerRole -and $fileServerRole.Installed) {
                Write-StatusMessage "File Server role is installed." -Status "Success"
                
                # Check if File Server Search is installed
                $fsSearch = Get-WindowsFeature -Name "FS-Search-Service" -ErrorAction SilentlyContinue
                if ($fsSearch -and $fsSearch.Installed) {
                    Write-StatusMessage "File Server Search Service is installed." -Status "Success"
                } else {
                    Write-StatusMessage "File Server Search Service is not installed." -Status "Info"
                    Write-StatusMessage "This may be needed for optimal SMB share search functionality." -Status "Info"
                }
            }
        }
        catch {
            Write-StatusMessage "Could not check File Server role status: $_" -Status "Error"
        }
    }
    
    # Check Search service status
    try {
        $searchService = Get-Service -Name "WSearch" -ErrorAction Stop
        $serviceStatus = $searchService.Status
        $startupType = $searchService.StartType
        
        if ($serviceStatus -eq "Running") {
            Write-StatusMessage "Windows Search service is running." -Status "Success"
        } else {
            Write-StatusMessage "Windows Search service is $serviceStatus." -Status "Warning"
            
            # Try to start the service if it's not running
            if ($serviceStatus -ne "Running") {
                Write-StatusMessage "Attempting to start Windows Search service..." -Status "Info"
                try {
                    Start-Service -Name "WSearch" -ErrorAction Stop
                    Write-StatusMessage "Windows Search service started successfully." -Status "Success"
                }
                catch {
                    Write-StatusMessage "Failed to start Windows Search service: $_" -Status "Error"
                    
                    # Check if this is because the service is disabled or not installed
                    if ($startupType -eq "Disabled") {
                        Write-StatusMessage "Search service is disabled. Consider changing startup type to Automatic." -Status "Warning"
                        if ($isServer) {
                            Write-StatusMessage "On Windows Server, Search service is often disabled by default." -Status "Info"
                        }
                    }
                }
            }
        }
        
        Write-StatusMessage "Service startup type: $startupType" -Status "Info"
    }
    catch {
        Write-StatusMessage "Error checking Windows Search service: $_" -Status "Error"
        
        # If service is not found, check if we're on Server Core
        if ($isServer) {
            $serverCore = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\").InstallationType -eq "Server Core"
            if ($serverCore) {
                Write-StatusMessage "Running on Server Core installation where Search service may not be available." -Status "Info"
            }
        }
    }
    
    # Define the registry check as a function within the scope - MUST BE DEFINED BEFORE IT'S USED
    function CheckIndexedLocationsViaRegistry {
        # Fallback to registry-based indexed locations check
        try {
            Write-StatusMessage "Checking indexed locations via registry..." -Status "Info"
            
            # Search for indexed locations in registry
            $scopesKey = "HKLM:\SOFTWARE\Microsoft\Windows Search\CrawlScopeManager\Windows\SystemIndex\WorkingSetRules"
            if (Test-Path -Path $scopesKey) {
                $scopes = Get-ChildItem -Path $scopesKey -ErrorAction SilentlyContinue
                if ($scopes -and $scopes.Count -gt 0) {
                    Write-StatusMessage "Found $(($scopes | Measure-Object).Count) indexed locations in registry." -Status "Success"
                    
                    # Try to sample some of the scope URLs if available
                    $sampleScopes = $scopes | Select-Object -First 5
                    foreach ($scope in $sampleScopes) {
                        $scopeProps = Get-ItemProperty -Path $scope.PSPath -ErrorAction SilentlyContinue
                        if ($scopeProps -and $scopeProps.URL) {
                            Write-StatusMessage "  - $($scopeProps.URL)" -Status "Note"
                        }
                    }
                    
                    $smbShares = $scopes | Where-Object { 
                        $url = (Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue).URL
                        $url -like "file://*" -and $url -like "*//*" 
                    }
                    
                    if ($smbShares -and $smbShares.Count -gt 0) {
                        Write-StatusMessage "SMB shares found in index configuration." -Status "Success"
                        
                        # Display sample SMB shares
                        $sampleSmbShares = $smbShares | Select-Object -First 3
                        foreach ($share in $sampleSmbShares) {
                            $url = (Get-ItemProperty -Path $share.PSPath -ErrorAction SilentlyContinue).URL
                            if ($url) {
                                Write-StatusMessage "  - $url" -Status "Note"
                            }
                        }
                        
                        if ($smbShares.Count -gt 3) {
                            Write-StatusMessage "  - ... and $($smbShares.Count - 3) more SMB shares" -Status "Note"
                        }
                    } else {
                        Write-StatusMessage "No SMB shares found in index configuration." -Status "Info"
                    }
                } else {
                    Write-StatusMessage "No indexed locations found in registry." -Status "Warning"
                }
            } else {
                # Try alternate registry locations
                $altScopesKeys = @(
                    "HKLM:\SOFTWARE\Microsoft\Windows Search\CrawlScopeManager\Windows\SystemIndex\DefaultRules",
                    "HKCU:\SOFTWARE\Microsoft\Windows Search\CrawlScopeManager\Windows\SystemIndex\WorkingSetRules"
                )
                
                $foundAltKey = $false
                foreach ($altKey in $altScopesKeys) {
                    if (Test-Path -Path $altKey) {
                        $foundAltKey = $true
                        $scopes = Get-ChildItem -Path $altKey -ErrorAction SilentlyContinue
                        if ($scopes -and $scopes.Count -gt 0) {
                            Write-StatusMessage "Found $(($scopes | Measure-Object).Count) indexed locations in alternate registry location." -Status "Success"
                            break
                        }
                    }
                }
                
                if (-not $foundAltKey) {
                    Write-StatusMessage "Could not find indexed locations in registry." -Status "Warning"
                }
            }
        }
        catch {
            Write-StatusMessage "Error checking indexed locations in registry: $_" -Status "Error"
        }
    }
    
    # Check if Windows Search API is available
    $searchApiAvailable = $false
    try {
        # Test COM class registration first using Registry
        $clsid = [Guid]::Parse("7D096C5F-AC08-4F1F-BEB7-5C22C517CE39") # Microsoft.Search.Administration.SearchManager CLSID
        $regKey = "HKCR:\CLSID\{$clsid}"
        if (Test-Path -Path $regKey -ErrorAction SilentlyContinue) {
            $searchApiAvailable = $true
        } else {
            # Try alternate registry path for servers
            $altRegKey = "HKLM:\SOFTWARE\Classes\CLSID\{$clsid}"
            if (Test-Path -Path $altRegKey -ErrorAction SilentlyContinue) {
                $searchApiAvailable = $true
            } else {
                Write-StatusMessage "Windows Search API is not registered on this system." -Status "Warning"
                if ($isServer) {
                    Write-StatusMessage "On Windows Server, Search components may need to be installed separately." -Status "Info" 
                } else {
                    Write-StatusMessage "Some Windows editions do not include the full Search API components." -Status "Info"
                }
            }
        }
        
        # Try to actually create the COM object to verify it's fully functional
        if ($searchApiAvailable) {
            try {
                $null = New-Object -ComObject Microsoft.Search.Administration.SearchManager -ErrorAction Stop
                # If we get here, the COM object creation was successful
            }
            catch {
                $searchApiAvailable = $false
                Write-StatusMessage "Search API appears to be registered but not functional: $_" -Status "Warning"
            }
        }
    }
    catch {
        $searchApiAvailable = $false
        Write-StatusMessage "Error checking Search API availability: $_" -Status "Error"
    }
    
    # For Windows Server, check if iFilters are installed
    if ($isServer) {
        Write-StatusMessage "Checking for installed iFilters (document filters for indexing)..." -Status "Info"
        try {
            # Check both common iFilter registry locations
            $iFiltersKeys = @(
                "HKLM:\SOFTWARE\Microsoft\Windows Search\IFilters",
                "HKLM:\SOFTWARE\Microsoft\Windows Search\Filters",
                "HKLM:\SOFTWARE\Microsoft\Windows Search\ContentIndexCommon\Filters",
                "HKLM:\SOFTWARE\Classes\CLSID\{975797FC-4E2A-11D0-B702-00C04FD8DBF7}\InprocServer32"
            )
            
            $filtersFound = $false
            
            foreach ($key in $iFiltersKeys) {
                if (Test-Path $key) {
                    $iFilters = Get-ChildItem -Path $key -ErrorAction SilentlyContinue
                    if ($iFilters -and $iFilters.Count -gt 0) {
                        $filtersFound = $true
                        Write-StatusMessage "Found $($iFilters.Count) registered iFilters in $key" -Status "Success"
                        $iFilters | ForEach-Object {
                            $filterName = Split-Path -Path $_.Name -Leaf
                            Write-StatusMessage "  - $filterName" -Status "Note"
                        }
                    }
                }
            }
            
            if (-not $filtersFound) {
                Write-StatusMessage "No iFilters registry keys found or no filters registered." -Status "Warning"
                Write-StatusMessage "This is common if Windows Search is not fully configured." -Status "Info"
            }
        }
        catch {
            Write-StatusMessage "Error checking iFilters: $_" -Status "Error"
        }
    }
    
    # Check Indexer status using WMI/COM if available
    if ($searchApiAvailable) {
        try {
            $searchManager = New-Object -ComObject Microsoft.Search.Administration.SearchManager
            $catalog = $searchManager.GetCatalog("SystemIndex")
            $manager = $catalog.GetIndexerManager()
            
            $indexerStatus = $manager.Status
            Write-StatusMessage "Indexer status: $indexerStatus" -Status $(if ($indexerStatus -eq 1) { "Success" } else { "Warning" })
            
            # Get indexing status
            $catalogStatus = $catalog.Status
            Write-StatusMessage "Catalog status: $catalogStatus" -Status "Info"
            
            # Get indexed item count
            $itemCount = $catalog.NumberOfItems
            Write-StatusMessage "Total indexed items: $itemCount" -Status "Info"
            
            # Check if indexing is ongoing
            if ($manager.GetStatus().IsInProgress) {
                $currentActivity = $manager.GetStatus().CurrentActivity
                Write-StatusMessage "Indexing is in progress: $currentActivity" -Status "Info"
                
                $remainingItems = $manager.GetStatus().NumberOfItemsToBeIndexed
                Write-StatusMessage "Items remaining to be indexed: $remainingItems" -Status "Info"
            }
            
            # Check indexed locations within the same try block to reuse the COM object
            Write-StatusMessage "Checking indexed locations..." -Status "Info"
            
            $crawlManager = $catalog.GetCrawlScopeManager()
            $rootScopes = $crawlManager.EnumerateRootScopes(1) # 1 = Default (User) scopes
            
            if ($rootScopes.Count -gt 0) {
                Write-StatusMessage "Indexed locations:" -Status "Success"
                foreach ($scope in $rootScopes) {
                    $status = if ($scope.IsIncluded) { "Included" } else { "Excluded" }
                    $statusColor = if ($scope.IsIncluded) { "Success" } else { "Warning" }
                    Write-StatusMessage "  - $($scope.URL): $status" -Status $statusColor
                }
            } else {
                Write-StatusMessage "No indexed locations found!" -Status "Warning"
            }
            
            # Check SMB shares in index
            $smbScopes = $rootScopes | Where-Object { $_.URL -like "file://*" -and $_.URL -like "*//*" }
            if ($smbScopes.Count -gt 0) {
                Write-StatusMessage "SMB shares in index:" -Status "Success"
                foreach ($scope in $smbScopes) {
                    $status = if ($scope.IsIncluded) { "Included" } else { "Excluded" }
                    $statusColor = if ($scope.IsIncluded) { "Success" } else { "Warning" }
                    Write-StatusMessage "  - $($scope.URL): $status" -Status $statusColor
                }
            } else {
                Write-StatusMessage "No SMB shares found in index." -Status "Info"
            }
        }
        catch {
            Write-StatusMessage "Error accessing Windows Search index: $_" -Status "Error"
            Write-StatusMessage "This may indicate corruption or configuration issues with the search index." -Status "Warning"
            $searchApiAvailable = $false
            
            # Fall back to registry-based check
            CheckIndexedLocationsViaRegistry
        }
    }
    else {
        # Fallback to basic WMI-based check for SearchIndexer process
        try {
            $searchProcess = Get-Process -Name "SearchIndexer" -ErrorAction SilentlyContinue
            if ($searchProcess) {
                Write-StatusMessage "SearchIndexer process is running (PID: $($searchProcess.Id))" -Status "Success"
                $cpuUsage = $searchProcess.CPU
                $memoryUsageMB = [math]::Round($searchProcess.WorkingSet / 1MB, 2)
                Write-StatusMessage "Process CPU time: $cpuUsage seconds, Memory: $memoryUsageMB MB" -Status "Info"
            } else {
                Write-StatusMessage "SearchIndexer process is not running." -Status "Warning"
            }
        }
        catch {
            Write-StatusMessage "Error checking SearchIndexer process: $_" -Status "Error"
        }
        
        # Use registry check for indexed locations when COM API isn't available
        CheckIndexedLocationsViaRegistry
    }
    
    # Check index health
    try {
        $indexHealth = Get-WmiObject -Namespace "root\Microsoft\Windows\Search" -Class "MSFT_SearchIndexer" -ErrorAction SilentlyContinue
        if ($indexHealth) {
            Write-StatusMessage "Search index health status: $($indexHealth.HealthStatus)" -Status "Info"
        } else {
            Write-StatusMessage "Could not determine search index health. Search may not be properly configured." -Status "Warning"
        }
    }
    catch {
        Write-StatusMessage "Error checking search index health: $_" -Status "Error"
    }
}

function Test-SMBNetworking {
    Write-StatusMessage "Checking network configuration for SMB..." -Status "Info"
    
    # Check if SMB ports are open in the firewall
    Write-StatusMessage "Checking Windows Firewall settings for SMB..." -Status "Info"
    
    try {
        $firewallRules = Get-NetFirewallRule -DisplayGroup "File and Printer Sharing" -ErrorAction Stop
        
        $enabledRules = $firewallRules | Where-Object { $_.Enabled -eq $true }
        $totalRules = $firewallRules.Count
        $enabledCount = $enabledRules.Count
        
        Write-StatusMessage "File and Printer Sharing firewall rules: $enabledCount of $totalRules enabled" -Status $(if ($enabledCount -gt 0) { "Success" } else { "Warning" })
        
        foreach ($rule in $firewallRules) {
            $status = if ($rule.Enabled) { "Enabled" } else { "Disabled" }
            $statusColor = if ($rule.Enabled) { "Success" } else { "Warning" }
            Write-StatusMessage "  - $($rule.DisplayName): $status" -Status $statusColor
        }
        
        # Check SMB-specific rules
        $smbRules = Get-NetFirewallRule -DisplayName "*SMB*" -ErrorAction SilentlyContinue
        if ($smbRules) {
            Write-StatusMessage "SMB-specific firewall rules:" -Status "Info"
            foreach ($rule in $smbRules) {
                $status = if ($rule.Enabled) { "Enabled" } else { "Disabled" }
                $statusColor = if ($rule.Enabled) { "Success" } else { "Warning" }
                Write-StatusMessage "  - $($rule.DisplayName): $status" -Status $statusColor
            }
        }
    }
    catch {
        Write-StatusMessage "Failed to retrieve firewall rules: $_" -Status "Error"
    }
    
    # Check network interfaces
    Write-StatusMessage "Checking network interfaces..." -Status "Info"
    try {
        $interfaces = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
        
        if ($interfaces.Count -eq 0) {
            Write-StatusMessage "No active network interfaces found!" -Status "Error"
        } else {
            Write-StatusMessage "Active network interfaces:" -Status "Success"
            foreach ($iface in $interfaces) {
                Write-StatusMessage "  - $($iface.Name): $($iface.InterfaceDescription)" -Status "Note"
                Write-StatusMessage "    Status: $($iface.Status), Speed: $($iface.LinkSpeed)" -Status "Note"
                
                # Get IP configuration for this interface
                $ipConfig = Get-NetIPAddress -InterfaceIndex $iface.ifIndex -ErrorAction SilentlyContinue
                foreach ($ip in $ipConfig) {
                    if ($ip.AddressFamily -eq "IPv4") {
                        Write-StatusMessage "    IPv4: $($ip.IPAddress)/$($ip.PrefixLength)" -Status "Note"
                    }
                }
            }
        }
    }
    catch {
        Write-StatusMessage "Failed to retrieve network interface information: $_" -Status "Error"
    }
    
    # Check for DNS resolution issues
    Write-StatusMessage "Checking DNS resolution..." -Status "Info"
    try {
        $computerName = $env:COMPUTERNAME
        $dnsResult = Resolve-DnsName -Name $computerName -ErrorAction SilentlyContinue
        
        if ($dnsResult) {
            Write-StatusMessage "DNS resolution for $computerName successful" -Status "Success"
            Write-StatusMessage "  - Resolved to: $($dnsResult.IPAddress)" -Status "Note"
        } else {
            Write-StatusMessage "Failed to resolve own hostname via DNS" -Status "Warning"
        }
    }
    catch {
        Write-StatusMessage "DNS resolution error: $_" -Status "Error"
    }
}

function Measure-SMBPerformance {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,
        
        [Parameter(Mandatory = $false)]
        [int]$FileSizeMB = 100,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipWrite
    )
    
    if (-not (Test-Path -Path $Path -ErrorAction SilentlyContinue)) {
        Write-StatusMessage "The specified path does not exist: $Path" -Status "Error"
        return
    }
    
    $testFile = Join-Path -Path $Path -ChildPath "SMB_Performance_Test_$(Get-Random).dat"
    
    Write-StatusMessage "Testing SMB performance with $FileSizeMB MB file..." -Status "Info"
    
    # Write test
    if (-not $SkipWrite) {
        try {
            Write-StatusMessage "Testing write performance..." -Status "Info" -NoNewline
            $writeTimer = [System.Diagnostics.Stopwatch]::StartNew()
            
            $buffer = New-Object byte[] (1MB)
            (New-Object Random).NextBytes($buffer)
            
            $stream = [System.IO.File]::OpenWrite($testFile)
            for ($i = 0; $i -lt $FileSizeMB; $i++) {
                $stream.Write($buffer, 0, $buffer.Length)
                
                # Show progress every 10 MB
                if ($i % 10 -eq 0) {
                    Write-Host "." -NoNewline
                }
            }
            $stream.Close()
            $writeTimer.Stop()
            Write-Host ""
            
            $writeSpeed = [math]::Round($FileSizeMB / $writeTimer.Elapsed.TotalSeconds, 2)
            Write-StatusMessage "Write complete: $writeSpeed MB/s" -Status "Success"
        }
        catch {
            Write-StatusMessage "Write test failed: $_" -Status "Error"
            return
        }
    }
    
    # Read test
    if (Test-Path -Path $testFile -ErrorAction SilentlyContinue) {
        try {
            Write-StatusMessage "Testing read performance..." -Status "Info" -NoNewline
            $readTimer = [System.Diagnostics.Stopwatch]::StartNew()
            
            $buffer = New-Object byte[] (1MB)
            $stream = [System.IO.File]::OpenRead($testFile)
            
            $totalRead = 0
            $bytesRead = 0
            
            do {
                $bytesRead = $stream.Read($buffer, 0, $buffer.Length)
                $totalRead += $bytesRead
                
                # Show progress every 10 MB
                if ($totalRead % (10MB) -lt $buffer.Length) {
                    Write-Host "." -NoNewline
                }
                
            } while ($bytesRead -gt 0)
            
            $stream.Close()
            $readTimer.Stop()
            Write-Host ""
            
            $readSpeed = [math]::Round(($totalRead / 1MB) / $readTimer.Elapsed.TotalSeconds, 2)
            Write-StatusMessage "Read complete: $readSpeed MB/s" -Status "Success"
        }
        catch {
            Write-StatusMessage "Read test failed: $_" -Status "Error"
        }
        
        # Clean up test file
        try {
            Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
            Write-StatusMessage "Test file removed." -Status "Info"
        }
        catch {
            Write-StatusMessage "Could not remove test file: $_" -Status "Warning"
        }
    }
    else {
        Write-StatusMessage "Test file not found. Write test may have failed." -Status "Error"
    }
    
    # Quick concurrent connection test
    try {
        Write-StatusMessage "Testing SMB connection handling..." -Status "Info"
        
        $concurrentTestPath = Split-Path -Path $Path -Parent
        $maxConnections = 10
        $connections = @()
        
        for ($i = 1; $i -le $maxConnections; $i++) {
            Write-Host "." -NoNewline
            
            try {
                $newConn = New-PSDrive -Name "TestSMB$i" -PSProvider FileSystem -Root $concurrentTestPath -ErrorAction Stop
                $connections += "TestSMB$i"
            }
            catch {
                Write-StatusMessage "Failed after $($i-1) connections: $_" -Status "Warning"
                break
            }
        }
        
        Write-Host ""
        Write-StatusMessage "Successfully established $($connections.Count) concurrent connections." -Status "Success"
        
        # Clean up the drives
        foreach ($conn in $connections) {
            Remove-PSDrive -Name $conn -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-StatusMessage "Connection test failed: $_" -Status "Error"
    }
}

function Show-MainMenu {
    $continue = $true
    
    while ($continue) {
        Clear-Host
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host "           SMB DIAGNOSTICS TOOL                " -ForegroundColor Cyan
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host " 1. Check SMB Configuration" -ForegroundColor Green
        Write-Host " 2. Enumerate Local SMB Shares" -ForegroundColor Green
        Write-Host " 3. Test Remote SMB Share" -ForegroundColor Green
        Write-Host " 4. Check Network Configuration for SMB" -ForegroundColor Green
        Write-Host " 5. Measure SMB Performance" -ForegroundColor Green
        Write-Host " 6. Check Windows Search Indexer" -ForegroundColor Yellow
        Write-Host " 7. Reset Windows Search Index" -ForegroundColor Yellow
        Write-Host " 8. Run All Diagnostics (Local)" -ForegroundColor Magenta
        Write-Host ""
        Write-Host " Q. Quit" -ForegroundColor Red
        Write-Host ""
        
        $choice = Read-Host "Enter your choice"
        
        switch ($choice) {
            "1" {
                Clear-Host
                Test-SMBConfiguration
                Read-Host "Press Enter to continue"
            }
            "2" {
                Clear-Host
                Get-SMBShareInfo
                Read-Host "Press Enter to continue"
            }
            "3" {
                Clear-Host
                $computerName = Read-Host "Enter remote computer name"
                $shareName = Read-Host "Enter share name (optional, leave blank to enumerate shares)"
                Test-RemoteSMBShare -ComputerName $computerName -ShareName $shareName
                Read-Host "Press Enter to continue"
            }
            "4" {
                Clear-Host
                Test-SMBNetworking
                Read-Host "Press Enter to continue"
            }
            "5" {
                Clear-Host
                $sharePath = Read-Host "Enter SMB share path to test (e.g., \\server\share or local path)"
                $fileSize = Read-Host "Enter test file size in MB (default: 100)"
                
                if ([string]::IsNullOrEmpty($fileSize)) {
                    $fileSize = 100
                } else {
                    $fileSize = [int]$fileSize
                }
                
                Measure-SMBPerformance -Path $sharePath -FileSizeMB $fileSize
                Read-Host "Press Enter to continue"
            }
            "6" {
                Clear-Host
                Test-WindowsSearchIndexer
                Read-Host "Press Enter to continue"
            }
            "7" {
                Clear-Host
                Reset-WindowsSearchIndex
                Read-Host "Press Enter to continue"
            }
            "8" {
                Clear-Host
                Write-Host "Running all local diagnostics..." -ForegroundColor Cyan
                
                Test-SMBConfiguration
                Write-Host ""
                
                Get-SMBShareInfo
                Write-Host ""
                
                Test-SMBNetworking
                Write-Host ""
                
                Test-WindowsSearchIndexer
                Write-Host ""
                
                Read-Host "Press Enter to continue"
            }
            "Q" {
                $continue = $false
            }
            "q" {
                $continue = $false
            }
            default {
                Write-Host "Invalid choice. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    }
}

# Start the tool
Show-MainMenu
