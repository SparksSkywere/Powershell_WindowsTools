# Performance Profile - GUI tool to analyze system performance and hardware information

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "System Performance Profile"
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::DPI
$form.MinimumSize = New-Object System.Drawing.Size(1000, 700)

# Hide console window
function Hide-Console {
    Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0) | Out-Null
}
Hide-Console

# Create tab control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(1160, 680)
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

# Create tabs
$systemTab = New-Object System.Windows.Forms.TabPage
$systemTab.Text = "System Info"
$systemTab.UseVisualStyleBackColor = $true

$hardwareTab = New-Object System.Windows.Forms.TabPage
$hardwareTab.Text = "Hardware"
$hardwareTab.UseVisualStyleBackColor = $true

$performanceTab = New-Object System.Windows.Forms.TabPage
$performanceTab.Text = "Performance"
$performanceTab.UseVisualStyleBackColor = $true

$processTab = New-Object System.Windows.Forms.TabPage
$processTab.Text = "Processes"
$processTab.UseVisualStyleBackColor = $true

$networkTab = New-Object System.Windows.Forms.TabPage
$networkTab.Text = "Network"
$networkTab.UseVisualStyleBackColor = $true

$tabControl.TabPages.Add($systemTab)
$tabControl.TabPages.Add($hardwareTab)
$tabControl.TabPages.Add($performanceTab)
$tabControl.TabPages.Add($processTab)
$tabControl.TabPages.Add($networkTab)

# System Info Tab Components
$systemListView = New-Object System.Windows.Forms.ListView
$systemListView.View = [System.Windows.Forms.View]::Details
$systemListView.FullRowSelect = $true
$systemListView.GridLines = $true
$systemListView.Location = New-Object System.Drawing.Point(10, 50)
$systemListView.Size = New-Object System.Drawing.Size(1130, 580)
$systemListView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$systemListView.Columns.Add("Property", 300) | Out-Null
$systemListView.Columns.Add("Value", 830) | Out-Null

$systemRefreshButton = New-Object System.Windows.Forms.Button
$systemRefreshButton.Location = New-Object System.Drawing.Point(10, 15)
$systemRefreshButton.Size = New-Object System.Drawing.Size(120, 30)
$systemRefreshButton.Text = "Refresh System Info"

$systemTab.Controls.Add($systemListView)
$systemTab.Controls.Add($systemRefreshButton)

# Hardware Tab Components
$hardwareListView = New-Object System.Windows.Forms.ListView
$hardwareListView.View = [System.Windows.Forms.View]::Details
$hardwareListView.FullRowSelect = $true
$hardwareListView.GridLines = $true
$hardwareListView.Location = New-Object System.Drawing.Point(10, 50)
$hardwareListView.Size = New-Object System.Drawing.Size(1130, 580)
$hardwareListView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$hardwareListView.Columns.Add("Component", 200) | Out-Null
$hardwareListView.Columns.Add("Name", 400) | Out-Null
$hardwareListView.Columns.Add("Details", 530) | Out-Null

$hardwareRefreshButton = New-Object System.Windows.Forms.Button
$hardwareRefreshButton.Location = New-Object System.Drawing.Point(10, 15)
$hardwareRefreshButton.Size = New-Object System.Drawing.Size(120, 30)
$hardwareRefreshButton.Text = "Refresh Hardware"

$hardwareTab.Controls.Add($hardwareListView)
$hardwareTab.Controls.Add($hardwareRefreshButton)

# Performance Tab Components
$performanceListView = New-Object System.Windows.Forms.ListView
$performanceListView.View = [System.Windows.Forms.View]::Details
$performanceListView.FullRowSelect = $true
$performanceListView.GridLines = $true
$performanceListView.Location = New-Object System.Drawing.Point(10, 50)
$performanceListView.Size = New-Object System.Drawing.Size(1130, 580)
$performanceListView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$performanceListView.Columns.Add("Metric", 250) | Out-Null
$performanceListView.Columns.Add("Current Value", 200) | Out-Null
$performanceListView.Columns.Add("Details", 680) | Out-Null

$performanceRefreshButton = New-Object System.Windows.Forms.Button
$performanceRefreshButton.Location = New-Object System.Drawing.Point(10, 15)
$performanceRefreshButton.Size = New-Object System.Drawing.Size(120, 30)
$performanceRefreshButton.Text = "Refresh Performance"

$performanceTab.Controls.Add($performanceListView)
$performanceTab.Controls.Add($performanceRefreshButton)

# Process Tab Components
$processListView = New-Object System.Windows.Forms.ListView
$processListView.View = [System.Windows.Forms.View]::Details
$processListView.FullRowSelect = $true
$processListView.GridLines = $true
$processListView.Location = New-Object System.Drawing.Point(10, 50)
$processListView.Size = New-Object System.Drawing.Size(1130, 580)
$processListView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$processListView.Columns.Add("Process Name", 200) | Out-Null
$processListView.Columns.Add("PID", 80) | Out-Null
$processListView.Columns.Add("CPU %", 80) | Out-Null
$processListView.Columns.Add("Memory (MB)", 120) | Out-Null
$processListView.Columns.Add("Handles", 80) | Out-Null
$processListView.Columns.Add("Threads", 80) | Out-Null
$processListView.Columns.Add("Path", 470) | Out-Null

$processRefreshButton = New-Object System.Windows.Forms.Button
$processRefreshButton.Location = New-Object System.Drawing.Point(10, 15)
$processRefreshButton.Size = New-Object System.Drawing.Size(120, 30)
$processRefreshButton.Text = "Refresh Processes"

$processTab.Controls.Add($processListView)
$processTab.Controls.Add($processRefreshButton)

# Network Tab Components
$networkListView = New-Object System.Windows.Forms.ListView
$networkListView.View = [System.Windows.Forms.View]::Details
$networkListView.FullRowSelect = $true
$networkListView.GridLines = $true
$networkListView.Location = New-Object System.Drawing.Point(10, 50)
$networkListView.Size = New-Object System.Drawing.Size(1130, 580)
$networkListView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

$networkListView.Columns.Add("Interface", 250) | Out-Null
$networkListView.Columns.Add("Status", 100) | Out-Null
$networkListView.Columns.Add("IP Address", 150) | Out-Null
$networkListView.Columns.Add("MAC Address", 150) | Out-Null
$networkListView.Columns.Add("Speed", 120) | Out-Null
$networkListView.Columns.Add("Type", 120) | Out-Null
$networkListView.Columns.Add("Details", 240) | Out-Null

$networkRefreshButton = New-Object System.Windows.Forms.Button
$networkRefreshButton.Location = New-Object System.Drawing.Point(10, 15)
$networkRefreshButton.Size = New-Object System.Drawing.Size(120, 30)
$networkRefreshButton.Text = "Refresh Network"

$networkTab.Controls.Add($networkListView)
$networkTab.Controls.Add($networkRefreshButton)

# Create status bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusBar.Items.Add($statusLabel) | Out-Null

# Create main buttons
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(10, 700)
$exportButton.Size = New-Object System.Drawing.Size(120, 30)
$exportButton.Text = "Export Report"
$exportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left

$refreshAllButton = New-Object System.Windows.Forms.Button
$refreshAllButton.Location = New-Object System.Drawing.Point(140, 700)
$refreshAllButton.Size = New-Object System.Drawing.Size(120, 30)
$refreshAllButton.Text = "Refresh All"
$refreshAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left

# Function to add list view item safely
function Add-ListViewItem {
    param (
        [System.Windows.Forms.ListView]$ListView,
        [string]$Property,
        [string]$Value,
        [string]$Details = ""
    )
    
    try {
        $item = New-Object System.Windows.Forms.ListViewItem($Property)
        if ($ListView.Columns.Count -eq 2) {
            $item.SubItems.Add($Value)
        } else {
            $item.SubItems.Add($Value)
            $item.SubItems.Add($Details)
        }
        $ListView.Items.Add($item)
    }
    catch {
        Write-Host "Error adding item: $Property - $($_.Exception.Message)"
    }
}

# Function to get system information
function Get-SystemInfo {
    $statusLabel.Text = "Loading system information..."
    $form.Refresh()
    
    $systemListView.Items.Clear()
    
    try {
        # Computer System Information
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem
        $bios = Get-CimInstance -ClassName Win32_BIOS
        $processor = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
        $pageFile = Get-CimInstance -ClassName Win32_PageFileUsage
        
        # Basic System Info
        Add-ListViewItem -ListView $systemListView -Property "Computer Name" -Value $computerSystem.Name
        Add-ListViewItem -ListView $systemListView -Property "Domain/Workgroup" -Value $(if ($computerSystem.PartOfDomain) { $computerSystem.Domain } else { $computerSystem.Workgroup })
        Add-ListViewItem -ListView $systemListView -Property "Operating System" -Value $operatingSystem.Caption
        Add-ListViewItem -ListView $systemListView -Property "OS Version" -Value $operatingSystem.Version
        Add-ListViewItem -ListView $systemListView -Property "OS Build" -Value $operatingSystem.BuildNumber
        Add-ListViewItem -ListView $systemListView -Property "System Architecture" -Value $operatingSystem.OSArchitecture
        Add-ListViewItem -ListView $systemListView -Property "Kernel Version" -Value $operatingSystem.Version
        Add-ListViewItem -ListView $systemListView -Property "Install Date" -Value $operatingSystem.InstallDate.ToString("yyyy-MM-dd HH:mm:ss")
        Add-ListViewItem -ListView $systemListView -Property "Last Boot Time" -Value $operatingSystem.LastBootUpTime.ToString("yyyy-MM-dd HH:mm:ss")
        Add-ListViewItem -ListView $systemListView -Property "System Uptime" -Value $((Get-Date) - $operatingSystem.LastBootUpTime).ToString("dd\.hh\:mm\:ss")
        
        # Hardware Info
        Add-ListViewItem -ListView $systemListView -Property "System Manufacturer" -Value $computerSystem.Manufacturer
        Add-ListViewItem -ListView $systemListView -Property "System Model" -Value $computerSystem.Model
        Add-ListViewItem -ListView $systemListView -Property "System Type" -Value $computerSystem.SystemType
        Add-ListViewItem -ListView $systemListView -Property "BIOS Version" -Value $bios.SMBIOSBIOSVersion
        Add-ListViewItem -ListView $systemListView -Property "BIOS Date" -Value $bios.ReleaseDate.ToString("yyyy-MM-dd")
        Add-ListViewItem -ListView $systemListView -Property "BIOS Manufacturer" -Value $bios.Manufacturer
        Add-ListViewItem -ListView $systemListView -Property "Serial Number" -Value $bios.SerialNumber
        
        # Memory Information
        Add-ListViewItem -ListView $systemListView -Property "Total Physical Memory" -Value "$([math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)) GB"
        Add-ListViewItem -ListView $systemListView -Property "Available Memory" -Value "$([math]::Round($operatingSystem.FreePhysicalMemory / 1MB, 2)) GB"
        Add-ListViewItem -ListView $systemListView -Property "Virtual Memory Total" -Value "$([math]::Round($operatingSystem.TotalVirtualMemorySize / 1MB, 2)) GB"
        Add-ListViewItem -ListView $systemListView -Property "Virtual Memory Available" -Value "$([math]::Round($operatingSystem.FreeVirtualMemory / 1MB, 2)) GB"
        
        # Page File Information
        if ($pageFile) {
            Add-ListViewItem -ListView $systemListView -Property "Page File Location" -Value $pageFile.Name
            Add-ListViewItem -ListView $systemListView -Property "Page File Size" -Value "$([math]::Round($pageFile.AllocatedBaseSize / 1024, 2)) GB"
            Add-ListViewItem -ListView $systemListView -Property "Page File Usage" -Value "$([math]::Round($pageFile.CurrentUsage / 1024, 2)) GB"
            Add-ListViewItem -ListView $systemListView -Property "Page File Peak Usage" -Value "$([math]::Round($pageFile.PeakUsage / 1024, 2)) GB"
        }
        
        # Processor Information
        Add-ListViewItem -ListView $systemListView -Property "Processor" -Value $processor.Name
        Add-ListViewItem -ListView $systemListView -Property "Processor ID" -Value $processor.ProcessorId
        Add-ListViewItem -ListView $systemListView -Property "Processor Cores" -Value $processor.NumberOfCores
        Add-ListViewItem -ListView $systemListView -Property "Logical Processors" -Value $processor.NumberOfLogicalProcessors
        Add-ListViewItem -ListView $systemListView -Property "Processor Speed" -Value "$($processor.MaxClockSpeed) MHz"
        Add-ListViewItem -ListView $systemListView -Property "Processor Cache L2" -Value "$($processor.L2CacheSize) KB"
        Add-ListViewItem -ListView $systemListView -Property "Processor Cache L3" -Value "$($processor.L3CacheSize) KB"
        Add-ListViewItem -ListView $systemListView -Property "Processor Architecture" -Value $processor.Architecture
        Add-ListViewItem -ListView $systemListView -Property "Processor Family" -Value $processor.Family
        Add-ListViewItem -ListView $systemListView -Property "Processor Stepping" -Value $processor.Stepping
        
        # OS Details
        Add-ListViewItem -ListView $systemListView -Property "Time Zone" -Value "$($operatingSystem.CurrentTimeZone) hours from UTC"
        Add-ListViewItem -ListView $systemListView -Property "System Directory" -Value $operatingSystem.SystemDirectory
        Add-ListViewItem -ListView $systemListView -Property "Windows Directory" -Value $operatingSystem.WindowsDirectory
        Add-ListViewItem -ListView $systemListView -Property "Boot Device" -Value $operatingSystem.BootDevice
        Add-ListViewItem -ListView $systemListView -Property "System Device" -Value $operatingSystem.SystemDevice
        Add-ListViewItem -ListView $systemListView -Property "Locale" -Value $operatingSystem.Locale
        Add-ListViewItem -ListView $systemListView -Property "Code Set" -Value $operatingSystem.CodeSet
        Add-ListViewItem -ListView $systemListView -Property "Country Code" -Value $operatingSystem.CountryCode
        
        # System Performance Settings
        Add-ListViewItem -ListView $systemListView -Property "Number of Users" -Value $operatingSystem.NumberOfUsers
        Add-ListViewItem -ListView $systemListView -Property "Max Process Memory" -Value "$([math]::Round($operatingSystem.MaxProcessMemorySize / 1KB, 2)) MB"
        Add-ListViewItem -ListView $systemListView -Property "System Cache" -Value "$([math]::Round(($operatingSystem.TotalVirtualMemorySize - $operatingSystem.TotalVisibleMemorySize) / 1MB, 2)) GB"
        
        $statusLabel.Text = "System information loaded successfully."
    }
    catch {
        $statusLabel.Text = "Error loading system information: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to retrieve system information: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to get hardware information
function Get-HardwareInfo {
    $statusLabel.Text = "Loading hardware information..."
    $form.Refresh()
    
    $hardwareListView.Items.Clear()
    
    try {
        # Processors with detailed info
        Get-CimInstance -ClassName Win32_Processor | ForEach-Object {
            $item = New-Object System.Windows.Forms.ListViewItem("CPU")
            $item.SubItems.Add($_.Name)
            $details = "Cores: $($_.NumberOfCores), Threads: $($_.NumberOfLogicalProcessors), Speed: $($_.MaxClockSpeed) MHz"
            $details += ", L2: $($_.L2CacheSize)KB, L3: $($_.L3CacheSize)KB, Voltage: $($_.CurrentVoltage)V"
            $item.SubItems.Add($details)
            $hardwareListView.Items.Add($item)
        }
        
        # Memory with detailed specifications
        Get-CimInstance -ClassName Win32_PhysicalMemory | ForEach-Object {
            $item = New-Object System.Windows.Forms.ListViewItem("Memory")
            $item.SubItems.Add("$($_.Manufacturer) - $(if ($_.PartNumber) { $_.PartNumber.Trim() } else { 'Unknown' })")
            $details = "Capacity: $([math]::Round($_.Capacity / 1GB, 2)) GB, Speed: $($_.Speed) MHz, Location: $($_.DeviceLocator)"
            $details += ", Type: $($_.MemoryType), Form Factor: $($_.FormFactor), Data Width: $($_.DataWidth) bits"
            $item.SubItems.Add($details)
            $hardwareListView.Items.Add($item)
        }
        
        # Storage with SMART data
        Get-CimInstance -ClassName Win32_DiskDrive | ForEach-Object {
            $item = New-Object System.Windows.Forms.ListViewItem("Storage")
            $item.SubItems.Add($_.Model)
            $details = "Size: $([math]::Round($_.Size / 1GB, 2)) GB, Interface: $($_.InterfaceType)"
            $details += ", Partitions: $($_.Partitions), Sectors: $($_.TotalSectors), Status: $($_.Status)"
            $item.SubItems.Add($details)
            $hardwareListView.Items.Add($item)
        }
        
        # Graphics Cards with driver details
        Get-CimInstance -ClassName Win32_VideoController | Where-Object { $_.Name -notlike "*Basic*" } | ForEach-Object {
            $item = New-Object System.Windows.Forms.ListViewItem("Graphics")
            $item.SubItems.Add($_.Name)
            $details = "Driver: $($_.DriverVersion), Date: $($_.DriverDate)"
            if ($_.AdapterRAM -gt 0) {
                $details += ", VRAM: $([math]::Round($_.AdapterRAM / 1MB, 0)) MB"
            }
            $details += ", Resolution: $($_.CurrentHorizontalResolution)x$($_.CurrentVerticalResolution)"
            $item.SubItems.Add($details)
            $hardwareListView.Items.Add($item)
        }
        
        # Network Adapters with detailed specs
        Get-CimInstance -ClassName Win32_NetworkAdapter | Where-Object { $_.PhysicalAdapter -eq $true -and $_.AdapterType -notlike "*Loopback*" } | ForEach-Object {
            $item = New-Object System.Windows.Forms.ListViewItem("Network")
            $item.SubItems.Add($_.Name)
            $details = "MAC: $($_.MACAddress), Speed: $($_.Speed), Type: $($_.AdapterType)"
            $details += ", Manufacturer: $($_.Manufacturer), PNP Device ID: $($_.PNPDeviceID)"
            $item.SubItems.Add($details)
            $hardwareListView.Items.Add($item)
        }
        
        # Motherboard Information
        $motherboard = Get-CimInstance -ClassName Win32_BaseBoard
        if ($motherboard) {
            $item = New-Object System.Windows.Forms.ListViewItem("Motherboard")
            $item.SubItems.Add("$($motherboard.Manufacturer) $($motherboard.Product)")
            $details = "Version: $($motherboard.Version), Serial: $($motherboard.SerialNumber)"
            $item.SubItems.Add($details)
            $hardwareListView.Items.Add($item)
        }
        
        # Sound Devices
        Get-CimInstance -ClassName Win32_SoundDevice | ForEach-Object {
            $item = New-Object System.Windows.Forms.ListViewItem("Audio")
            $item.SubItems.Add($_.Name)
            $details = "Manufacturer: $($_.Manufacturer), Status: $($_.Status)"
            $item.SubItems.Add($details)
            $hardwareListView.Items.Add($item)
        }
        
        # USB Controllers
        Get-CimInstance -ClassName Win32_USBController | ForEach-Object {
            $item = New-Object System.Windows.Forms.ListViewItem("USB Controller")
            $item.SubItems.Add($_.Name)
            $details = "Manufacturer: $($_.Manufacturer), Protocol: $($_.ProtocolSupported)"
            $item.SubItems.Add($details)
            $hardwareListView.Items.Add($item)
        }
        
        $statusLabel.Text = "Hardware information loaded successfully."
    }
    catch {
        $statusLabel.Text = "Error loading hardware information: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to retrieve hardware information: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to get performance metrics
function Get-PerformanceInfo {
    $statusLabel.Text = "Loading performance metrics..."
    $form.Refresh()
    
    $performanceListView.Items.Clear()
    
    try {
        # CPU Usage with per-core breakdown
        $cpuUsage = (Get-CimInstance -ClassName Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average
        Add-ListViewItem -ListView $performanceListView -Property "CPU Usage (Average)" -Value "$cpuUsage%" -Details "Average across all cores"
        
        # Individual CPU core usage
        $cpuCores = Get-CimInstance -ClassName Win32_Processor
        for ($i = 0; $i -lt $cpuCores.Count; $i++) {
            Add-ListViewItem -ListView $performanceListView -Property "CPU Core $i Usage" -Value "$($cpuCores[$i].LoadPercentage)%" -Details "Individual core utilization"
        }
        
        # Memory Usage with detailed breakdown
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $totalMemory = $os.TotalVisibleMemorySize / 1MB
        $freeMemory = $os.FreePhysicalMemory / 1MB
        $usedMemory = $totalMemory - $freeMemory
        $memoryUsagePercent = [math]::Round(($usedMemory / $totalMemory) * 100, 2)
        Add-ListViewItem -ListView $performanceListView -Property "Physical Memory Usage" -Value "$memoryUsagePercent%" -Details "Used: $([math]::Round($usedMemory, 2)) GB / Total: $([math]::Round($totalMemory, 2)) GB"
        
        # Kernel Memory
        $kernelPaged = [math]::Round($os.SizeStoredInPagingFiles / 1MB, 2)
        $kernelNonPaged = [math]::Round(($os.TotalVirtualMemorySize - $os.TotalVisibleMemorySize) / 1MB, 2)
        Add-ListViewItem -ListView $performanceListView -Property "Kernel Memory (Paged)" -Value "$kernelPaged MB" -Details "Kernel memory that can be paged to disk"
        Add-ListViewItem -ListView $performanceListView -Property "Kernel Memory (Non-Paged)" -Value "$kernelNonPaged MB" -Details "Kernel memory that must stay in physical memory"
        
        # Cache Information
        $cacheBytes = Get-Counter '\Memory\Cache Bytes' -ErrorAction SilentlyContinue
        if ($cacheBytes) {
            $cacheSize = [math]::Round($cacheBytes.CounterSamples[0].CookedValue / 1MB, 2)
            Add-ListViewItem -ListView $performanceListView -Property "System Cache" -Value "$cacheSize MB" -Details "File system cache size"
        }
        
        # Disk Usage with I/O stats
        Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
            $usedSpace = ($_.Size - $_.FreeSpace) / 1GB
            $totalSpace = $_.Size / 1GB
            $freeSpace = $_.FreeSpace / 1GB
            $usagePercent = [math]::Round(($usedSpace / $totalSpace) * 100, 2)
            Add-ListViewItem -ListView $performanceListView -Property "Disk Usage ($($_.DeviceID))" -Value "$usagePercent%" -Details "Used: $([math]::Round($usedSpace, 2)) GB / Free: $([math]::Round($freeSpace, 2)) GB / Total: $([math]::Round($totalSpace, 2)) GB"
        }
        
        # Process and Thread counts with proper calculation
        $processCount = (Get-Process).Count
        Add-ListViewItem -ListView $performanceListView -Property "Running Processes" -Value $processCount -Details "Total number of running processes"
        
        # Fixed thread count calculation
        $threadCount = 0
        Get-Process | ForEach-Object {
            try {
                $threadCount += $_.Threads.Count
            }
            catch {
                # Skip processes that can't be accessed
            }
        }
        Add-ListViewItem -ListView $performanceListView -Property "Total Threads" -Value $threadCount -Details "Sum of all process threads"
        
        # Handle Count
        $handleCount = (Get-Process | Where-Object { $_.HandleCount } | Measure-Object -Property HandleCount -Sum).Sum
        Add-ListViewItem -ListView $performanceListView -Property "Total Handles" -Value $handleCount -Details "Sum of all process handles"
        
        # Context Switches
        $contextSwitches = Get-Counter '\System\Context Switches/sec' -ErrorAction SilentlyContinue
        if ($contextSwitches) {
            $switchRate = [math]::Round($contextSwitches.CounterSamples[0].CookedValue, 0)
            Add-ListViewItem -ListView $performanceListView -Property "Context Switches/sec" -Value $switchRate -Details "CPU context switches per second"
        }
        
        # System Calls
        $systemCalls = Get-Counter '\System\System Calls/sec' -ErrorAction SilentlyContinue
        if ($systemCalls) {
            $callRate = [math]::Round($systemCalls.CounterSamples[0].CookedValue, 0)
            Add-ListViewItem -ListView $performanceListView -Property "System Calls/sec" -Value $callRate -Details "System calls per second"
        }
        
        # Page Faults
        $pageFaults = Get-Counter '\Memory\Page Faults/sec' -ErrorAction SilentlyContinue
        if ($pageFaults) {
            $faultRate = [math]::Round($pageFaults.CounterSamples[0].CookedValue, 0)
            Add-ListViewItem -ListView $performanceListView -Property "Page Faults/sec" -Value $faultRate -Details "Memory page faults per second"
        }
        
        # Services
        $services = Get-Service
        $runningServices = ($services | Where-Object { $_.Status -eq "Running" }).Count
        $stoppedServices = ($services | Where-Object { $_.Status -eq "Stopped" }).Count
        Add-ListViewItem -ListView $performanceListView -Property "Windows Services" -Value "Running: $runningServices" -Details "Total: $($services.Count), Stopped: $stoppedServices"
        
        # Registry Information
        $registrySize = Get-Counter '\System\Registry Quota In Use (%)' -ErrorAction SilentlyContinue
        if ($registrySize) {
            $regUsage = [math]::Round($registrySize.CounterSamples[0].CookedValue, 2)
            Add-ListViewItem -ListView $performanceListView -Property "Registry Usage" -Value "$regUsage%" -Details "Registry quota utilization"
        }
        
        $statusLabel.Text = "Performance metrics loaded successfully."
    }
    catch {
        $statusLabel.Text = "Error loading performance metrics: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to retrieve performance metrics: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to get process information
function Get-ProcessInfo {
    $statusLabel.Text = "Loading process information..."
    $form.Refresh()
    
    $processListView.Items.Clear()
    
    try {
        $processes = Get-Process | Where-Object { $_.ProcessName -ne "Idle" } | Sort-Object -Property CPU -Descending | Select-Object -First 100
        
        foreach ($process in $processes) {
            try {
                $item = New-Object System.Windows.Forms.ListViewItem($process.ProcessName)
                $item.SubItems.Add($process.Id.ToString())
                $item.SubItems.Add($(if ($process.CPU) { [math]::Round($process.CPU, 2).ToString() } else { "0" }))
                $item.SubItems.Add([math]::Round($process.WorkingSet / 1MB, 2).ToString())
                $item.SubItems.Add($(if ($process.HandleCount) { $process.HandleCount.ToString() } else { "0" }))
                $item.SubItems.Add($(if ($process.Threads) { $process.Threads.Count.ToString() } else { "0" }))
                $item.SubItems.Add($(if ($process.Path) { $process.Path } else { "N/A" }))
                $processListView.Items.Add($item)
            }
            catch {
                # Skip processes that can't be accessed
                continue
            }
        }
        
        $statusLabel.Text = "Process information loaded successfully (Top 100 by CPU usage)."
    }
    catch {
        $statusLabel.Text = "Error loading process information: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to retrieve process information: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to get network information
function Get-NetworkInfo {
    $statusLabel.Text = "Loading network information..."
    $form.Refresh()
    
    $networkListView.Items.Clear()
    
    try {
        $networkAdapters = Get-NetAdapter
        $ipConfigs = Get-NetIPAddress
        
        foreach ($adapter in $networkAdapters) {
            try {
                $item = New-Object System.Windows.Forms.ListViewItem($adapter.Name)
                $item.SubItems.Add($adapter.Status)
                
                # Get IP address for this adapter
                $ipAddress = ($ipConfigs | Where-Object { $_.InterfaceAlias -eq $adapter.Name -and $_.AddressFamily -eq "IPv4" } | Select-Object -First 1).IPAddress
                if (-not $ipAddress) { $ipAddress = "N/A" }
                $item.SubItems.Add($ipAddress)
                
                $item.SubItems.Add($adapter.MacAddress)
                
                # Format speed
                $speed = if ($adapter.LinkSpeed -gt 1000000000) { 
                    "$([math]::Round($adapter.LinkSpeed / 1000000000, 1)) Gbps" 
                } elseif ($adapter.LinkSpeed -gt 1000000) { 
                    "$([math]::Round($adapter.LinkSpeed / 1000000, 0)) Mbps" 
                } else { 
                    "$($adapter.LinkSpeed) bps" 
                }
                $item.SubItems.Add($speed)
                
                $item.SubItems.Add($adapter.MediaType)
                
                $details = "Driver: $($adapter.DriverVersion), VLAN: $($adapter.VlanID)"
                $item.SubItems.Add($details)
                
                $networkListView.Items.Add($item)
            }
            catch {
                continue
            }
        }
        
        $statusLabel.Text = "Network information loaded successfully."
    }
    catch {
        $statusLabel.Text = "Error loading network information: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to retrieve network information: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to export all data to a report
function Export-Report {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text files (*.txt)|*.txt|HTML files (*.html)|*.html"
    $saveFileDialog.Title = "Export Performance Report"
    $saveFileDialog.FileName = "PerformanceReport_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $report = @()
            $report += "System Performance Report"
            $report += "Generated: $(Get-Date)"
            $report += "=" * 50
            $report += ""
            
            # Add system info
            $report += "SYSTEM INFORMATION"
            $report += "-" * 20
            foreach ($item in $systemListView.Items) {
                $report += "$($item.Text): $($item.SubItems[1].Text)"
            }
            $report += ""
            
            # Add hardware info
            $report += "HARDWARE INFORMATION"
            $report += "-" * 20
            foreach ($item in $hardwareListView.Items) {
                $report += "$($item.Text) - $($item.SubItems[1].Text): $($item.SubItems[2].Text)"
            }
            $report += ""
            
            # Add performance info
            $report += "PERFORMANCE METRICS"
            $report += "-" * 20
            foreach ($item in $performanceListView.Items) {
                $report += "$($item.Text): $($item.SubItems[1].Text) - $($item.SubItems[2].Text)"
            }
            $report += ""
            
            # Add network info
            $report += "NETWORK INFORMATION"
            $report += "-" * 20
            foreach ($item in $networkListView.Items) {
                $report += "$($item.Text): Status=$($item.SubItems[1].Text), IP=$($item.SubItems[2].Text), MAC=$($item.SubItems[3].Text), Speed=$($item.SubItems[4].Text)"
            }
            $report += ""
            
            # Add top processes
            $report += "TOP PROCESSES (by CPU usage)"
            $report += "-" * 30
            foreach ($item in $processListView.Items | Select-Object -First 20) {
                $report += "$($item.Text) (PID: $($item.SubItems[1].Text)) - CPU: $($item.SubItems[2].Text), Memory: $($item.SubItems[3].Text) MB"
            }
            
            $report | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Performance report exported successfully to:`n$($saveFileDialog.FileName)", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to export report: $($_.Exception.Message)", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Function to refresh all data
function Refresh-AllData {
    Get-SystemInfo
    Get-HardwareInfo
    Get-PerformanceInfo
    Get-ProcessInfo
    Get-NetworkInfo
}

# Event handlers
$systemRefreshButton.Add_Click({ Get-SystemInfo })
$hardwareRefreshButton.Add_Click({ Get-HardwareInfo })
$performanceRefreshButton.Add_Click({ Get-PerformanceInfo })
$processRefreshButton.Add_Click({ Get-ProcessInfo })
$networkRefreshButton.Add_Click({ Get-NetworkInfo })
$exportButton.Add_Click({ Export-Report })
$refreshAllButton.Add_Click({ Refresh-AllData })

# Add controls to the form
$form.Controls.Add($tabControl)
$form.Controls.Add($exportButton)
$form.Controls.Add($refreshAllButton)
$form.Controls.Add($statusBar)

# Load data when the form loads
$form.Add_Shown({
    Refresh-AllData
})

# Show the form
$form.ShowDialog() | Out-Null
