# WINDOWS ONLY

Write-Host "Please make sure you have edited this file and put your required port and description"

# Function to check and enable UPnP service
function Enable-UPnP {
    # Set the most basic values needed for this script
    $ssdpService = Get-Service -Name "SSDPSRV" -ErrorAction SilentlyContinue
    $upnpService = Get-Service -Name "upnphost" -ErrorAction SilentlyContinue

    # Checks the SSDP discovery service on the PC
    if ($null -eq $ssdpService) {
        Write-Output "SSDP Discovery service not found. UPnP might not be installed on this system."
        exit 1
    }

    # Checks the UPnP service host
    if ($null -eq $upnpService) {
        Write-Output "UPnP Device Host service not found. UPnP might not be installed on this system."
        exit 1
    }

    # If service is not running, start service
    if ($ssdpService.Status -ne 'Running') {
        Write-Output "Starting SSDP Discovery service..."
        Start-Service -Name "SSDPSRV"
    }

    # If service is not running, start service
    if ($upnpService.Status -ne 'Running') {
        Write-Output "Starting UPnP Device Host service..."
        Start-Service -Name "upnphost"
    }
}

# Function to add port mapping
function Add-PortMapping {
    param (
        [int]$externalPort,
        [int]$internalPort,
        [string]$internalClient,
        [string]$protocol,
        [string]$description
    )

    # Load the UPnP NAT interface
    try {
        $comObject = New-Object -ComObject HNetCfg.NATUPnP
        $natServices = $comObject.StaticPortMappingCollection
    } catch {
        Write-Output "Failed to create UPnP NAT object. Make sure UPnP is enabled on your router."
        exit 1
    }

    # If the UPnP is not supported on the router, create exception and close
    if ($null -eq $natServices) {
        Write-Output "UPnP not supported or enabled on your router."
        exit 1
    }

    # Add the port mapping
    try {
        $natServices.Add($externalPort, $protocol, $internalPort, $internalClient, $true, $description)
        Write-Output "Port $externalPort mapped to $internalClient successfully."
    } catch {
        Write-Output "Failed to map port $externalPort. Error: $_"
    }
}

# Function to remove the port mapping
function Remove-PortMapping {
    param (
        [int]$Port,
        [string]$Protocol
    )
    try {
        $natServices.Remove($Port, $Protocol)
        Write-Output "Port $Port unmapped successfully."
    } catch {
        Write-Output "Failed to unmap port $Port. Error: $_"
    }
}

# Define the port, protocol, and description, edit accordingly. (note, the firewall must be allowing inbound and outbound traffic for the port)
$externalPort = 000 # any port used for programs, eg: 80 for HTTP or 443 for HTTPS, 25565 for Minecraft
$internalPort = 000 # same as above
$internalClient = (Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress -ne $null }).IPAddress[0] # Leave this line alone
$protocol = "TCP"  # or "UDP" (note, the firewall must be set on the pc to allow this protocol)
$description = "Example Server" # Any program name can go here if you need

# Enable UPnP services
Enable-UPnP

# Infinite loop to keep the port open and refresh every hour
while ($true) {
    # Add the port mapping
    Add-PortMapping -externalPort $externalPort -internalPort $internalPort -internalClient $internalClient -protocol $protocol -description $description

    # Sleep for 1 hour
    Start-Sleep -Seconds 3600 # Adjust time accordingly, default is 1 hour (3600 seconds)

    # Remove the port mapping before re-adding it
    Remove-PortMapping -Port $externalPort -Protocol $protocol
}

# END SCRIPT BELOW

# Wait for user input to close the port after use, remove these lines below to loop until the script is closed forcibly
Write-Output "Press Enter to close the port..."
Read-Host

# Remove the port mapping
Remove-PortMapping -Port $externalPort -Protocol $protocol

# Script made by Sparks Skywere