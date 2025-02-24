<#
.SYNOPSIS
    Documents the complete network configuration of ESXi hosts managed by a vCenter server.

.DESCRIPTION
    This script connects to a vCenter server, collects detailed network configuration data 
    for all managed ESXi hosts, and generates individual HTML reports for each host. The 
    reports include standard and distributed vSwitch configurations, VMkernel interfaces, 
    DNS and routing details, firewall rules, NTP settings, and physical NIC information 
    with CDP/LLDP data.

.PARAMETER None
    This script does not accept parameters. It prompts for vCenter server details 
    and credentials interactively.

.EXAMPLE
    .\Get-ESXiNetworkConfig.ps1
    Runs the script, prompting for vCenter server hostname/IP and credentials, then 
    generates HTML reports in the current directory.

.OUTPUTS
    HTML files named "NetworkConfig_<hostname>_<timestamp>.html" for each ESXi host.

.NOTES
    - Requires VMware PowerCLI module to be installed. Install with: Install-Module -Name VMware.PowerCLI
    - Must be run with sufficient vCenter permissions to view host configurations.
    - Ignores invalid SSL certificates by default.
    - Current date used in script execution: February 24, 2025

    Version: 1.0
    Last Updated: February 24, 2025
#>

# Suppress PowerCLI welcome message for cleaner output
$PSDefaultParameterValues['Out-Default:Width'] = 200

# Set PowerCLI to ignore invalid certificates and suppress output
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null

# Generate timestamp for output file naming
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Prompt user for vCenter server details
$vCenterServer = Read-Host "Enter vCenter Server hostname or IP"
$credential = Get-Credential -Message "Enter vCenter credentials"

try {
    # Connect to vCenter with error handling
    Write-Host "Connecting to $vCenterServer..."
    Connect-VIServer -Server $vCenterServer -Credential $credential -ErrorAction Stop
}
catch {
    Write-Warning "Failed to connect to $vCenterServer : $_"
    exit
}

# Define CSS for HTML output styling
$css = @"
<style>
    body { font-family: Arial, sans-serif; margin: 20px; }
    h1 { color: #2c3e50; }
    h2 { color: #34495e; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background-color: #3498db; color: white; }
    tr:nth-child(even) { background-color: #f2f2f2; }
    .toc { margin-bottom: 20px; }
    .toc a { margin-right: 15px; }
</style>
"@

# Get all ESXi hosts managed by vCenter
$vmHosts = Get-VMHost | Sort-Object Name

foreach ($vmHost in $vmHosts) {
    Write-Host "Processing network configuration for host: $($vmHost.Name)"
    
    # Get current date and time for subheading
    $reportDateTime = Get-Date -Format "MMMM dd, yyyy HH:mm:ss"
    
    # Initialize HTML content array with main heading and italicized datetime subheading
    $htmlContent = [System.Collections.ArrayList]::new()
    $htmlContent.Add("<html><head>$css</head><body><h1>Network Configuration - $($vmHost.Name)</h1><h3><i>$reportDateTime</i></h3>") | Out-Null
    
    # Table of Contents
    $toc = '<div class="toc"><h2>Table of Contents</h2>'
    $toc += '<a href="#vSwitch">Standard vSwitches</a>'
    $toc += '<a href="#vmkernel">VMkernel Interfaces</a>'
    $toc += '<a href="#dvSwitch">Distributed vSwitches</a>'
    $toc += '<a href="#dnsRouting">DNS and Routing</a>'
    $toc += '<a href="#firewall">Firewall Rules</a>'
    $toc += '<a href="#ntp">NTP Settings</a>'
    $toc += '<a href="#physicalNics">Physical NICs</a>'
    $toc += '</div>'
    $htmlContent.Add($toc) | Out-Null

    try {
        # --- Standard vSwitch Configuration ---
        $htmlContent.Add('<h2 id="vSwitch">Standard vSwitches</h2>') | Out-Null
        $vSwitches = Get-VirtualSwitch -VMHost $vmHost -Standard
        $vmkAdapters = Get-VMHostNetworkAdapter -VMHost $vmHost -VMKernel
        
        $vSwitchData = foreach ($vSwitch in $vSwitches) {
            $security = $vSwitch | Get-SecurityPolicy
            $teaming = $vSwitch | Get-NicTeamingPolicy
            
            # Get VMkernel adapters and VLANs associated with this vSwitch
            $relatedVmk = $vmkAdapters | Where-Object { 
                $_.PortGroupName -in (Get-VirtualPortGroup -VirtualSwitch $vSwitch).Name 
            }
            $vmkList = if ($relatedVmk) {
                # Collect results into an array explicitly before joining
                $vmkArray = @($relatedVmk | ForEach-Object { 
                    $vlan = (Get-VirtualPortGroup -Name $_.PortGroupName -VMHost $vmHost).VLanId
                    "$($_.Name) (VLAN $vlan)"
                })
                [String]::Join(', ', $vmkArray)
            } else {
                'None'
            }

            [PSCustomObject]@{
                Name = $vSwitch.Name
                Ports = $vSwitch.NumPorts
                MTU = $vSwitch.MTU
                NICs = ($vSwitch.Nic -join ', ')
                Promiscuous = $security.AllowPromiscuous
                ForgedTransmits = $security.ForgedTransmits
                MacChanges = $security.MacChanges
                LoadBalancing = $teaming.LoadBalancingPolicy
                ActiveNICs = ($teaming.ActiveNic -join ', ')
                StandbyNICs = ($teaming.StandbyNic -join ', ')
                VMkernels_VLANs = $vmkList
            }
        }
        $htmlContent.Add(($vSwitchData | ConvertTo-Html -Fragment)) | Out-Null

        # --- VMkernel Interfaces ---
        $htmlContent.Add('<h2 id="vmkernel">VMkernel Interfaces</h2>') | Out-Null
        # Get host network config for VMkernel-specific gateway
        $hostNetwork = Get-VMHostNetwork -VMHost $vmHost
        $vmkData = foreach ($vmk in $vmkAdapters) {
            # Find the VMkernel-specific gateway from IpRouteConfig
            $vmkGateway = $hostNetwork.ExtensionData.IpRoute | Where-Object {
                $_.Network -eq '0.0.0.0/0' -and $_.Gateway -ne $hostNetwork.DefaultGateway
            } | Select-Object -First 1 -ExpandProperty Gateway
            if (-not $vmkGateway) { $vmkGateway = 'N/A' }  # Default to N/A if no specific gateway

            [PSCustomObject]@{
                Name = $vmk.Name
                IP = $vmk.IP
                SubnetMask = $vmk.SubnetMask
                MAC = $vmk.Mac
                PortGroup = $vmk.PortGroupName
                Gateway = $vmkGateway  # VMkernel-specific gateway
                VMotion = $vmk.VMotionEnabled
                FTLogging = $vmk.FaultToleranceLoggingEnabled
                Management = $vmk.ManagementTrafficEnabled
                VLAN = (Get-VirtualPortGroup -Name $vmk.PortGroupName -VMHost $vmHost).VLanId
                MTU = $vmk.MTU
            }
        }
        $htmlContent.Add(($vmkData | ConvertTo-Html -Fragment)) | Out-Null

        # --- Distributed vSwitch Configuration ---
        $htmlContent.Add('<h2 id="dvSwitch">Distributed vSwitches</h2>') | Out-Null
        $dvSwitches = Get-VDSwitch -VMHost $vmHost
        $dvSwitchData = foreach ($dvSwitch in $dvSwitches) {
            $security = $dvSwitch | Get-VDSecurityPolicy
            $teaming = $dvSwitch | Get-VDUplinkTeamingPolicy
            # Get physical NICs (uplinks) associated with this dvSwitch for this host
            $dvUplinks = Get-VMHostNetworkAdapter -VMHost $vmHost -DistributedSwitch $dvSwitch | Where-Object { $_.DeviceType -eq 'Physical' }
            $uplinkNames = $dvUplinks | ForEach-Object { $_.Name } -join ', '
            if (-not $uplinkNames) { $uplinkNames = 'None' }

            [PSCustomObject]@{
                Name = $dvSwitch.Name
                Ports = $dvSwitch.NumPorts
                MTU = $dvSwitch.Mtu
                NICs = $uplinkNames  # Physical NICs assigned as uplinks
                Promiscuous = $security.AllowPromiscuous
                ForgedTransmits = $security.ForgedTransmits
                MacChanges = $security.MacChanges
                LoadBalancing = $teaming.LoadBalancingPolicy
                ActiveNICs = ($teaming.ActiveUplink -join ', ')
                StandbyNICs = ($teaming.StandbyUplink -join ', ')
            }
        }
        $htmlContent.Add(($dvSwitchData | ConvertTo-Html -Fragment)) | Out-Null

        # --- DNS and Routing ---
        $htmlContent.Add('<h2 id="dnsRouting">DNS and Routing</h2>') | Out-Null
        $networkConfig = Get-VMHostNetwork -VMHost $vmHost
        $dnsRoutingData = [PSCustomObject]@{
            DNSServers = ($networkConfig.DnsAddress -join ', ')
            StaticRoutes = ($vmHost | Get-VMHostRoute | ForEach-Object { "$($_.Destination)/$($_.PrefixLength) via $($_.Gateway)" }) -join ', '
        }
        $htmlContent.Add(($dnsRoutingData | ConvertTo-Html -Fragment)) | Out-Null

        # --- Firewall Rules ---
        $htmlContent.Add('<h2 id="firewall">Firewall Rules</h2>') | Out-Null
        $firewallRules = Get-VMHostFirewallException -VMHost $vmHost | Where-Object { $_.Enabled }
        $firewallData = $firewallRules | Select-Object Name, Enabled, Protocol, @{N='Port';E={$_.Port}}
        $htmlContent.Add(($firewallData | ConvertTo-Html -Fragment)) | Out-Null

        # --- NTP Settings ---
        $htmlContent.Add('<h2 id="ntp">NTP Settings</h2>') | Out-Null
        $ntpConfig = Get-VMHostNtpServer -VMHost $vmHost
        $ntpService = Get-VMHostService -VMHost $vmHost | Where-Object { $_.Key -eq 'ntpd' }
        $ntpData = [PSCustomObject]@{
            NTPServers = ($ntpConfig -join ', ')
            Running = $ntpService.Running
            Policy = $ntpService.Policy
        }
        $htmlContent.Add(($ntpData | ConvertTo-Html -Fragment)) | Out-Null

        # --- Physical NIC Hardware Information with Network Hints ---
        $htmlContent.Add('<h2 id="physicalNics">Physical NICs</h2>') | Out-Null
        $physicalNics = Get-VMHostNetworkAdapter -VMHost $vmHost -Physical
        
        # Get Network Hints using current method
        $netSystem = Get-View -Id $vmHost.ExtensionData.ConfigManager.NetworkSystem
        $networkHints = $netSystem.QueryNetworkHint($null)
        $hintTable = @{}
        foreach ($hint in $networkHints) {
            $hintTable[$hint.Device] = $hint
        }

        $nicData = foreach ($nic in $physicalNics) {
            $hint = $hintTable[$nic.Name]
            $cdp = $hint.ConnectedSwitchPort
            $lldp = $hint.LLDPInfo
            [PSCustomObject]@{
                Name = $nic.Name
                MAC = $nic.Mac
                LinkSpeed = "$($nic.LinkSpeedMb) Mb/s"
                vSwitches = (($vSwitches | Where-Object { $_.Nic -contains $nic.Name }).Name -join ', ')
                CDP_Switch = if ($cdp) { $cdp.DevId } else { 'N/A' }
                CDP_Port = if ($cdp) { $cdp.PortId } else { 'N/A' }
                CDP_Hardware = if ($cdp) { $cdp.HardwarePlatform } else { 'N/A' }
                LLDP_Switch = if ($lldp) { $lldp.SystemName } else { 'N/A' }
                LLDP_Port = if ($lldp) { $lldp.PortId } else { 'N/A' }
                LLDP_Hardware = if ($lldp) { $lldp.ChassisId } else { 'N/A' }
            }
        }
        $htmlContent.Add(($nicData | ConvertTo-Html -Fragment)) | Out-Null

        # Close HTML
        $htmlContent.Add('</body></html>') | Out-Null

        # Write to file
        $fileName = "NetworkConfig_$($vmHost.Name)_$timestamp.html"
        $htmlContent | Out-File -FilePath $fileName -Encoding UTF8
        Write-Host "Generated report: $fileName"
    }
    catch {
        Write-Warning "Error processing $($vmHost.Name): $_"
    }
}

# Disconnect from vCenter without confirmation
Disconnect-VIServer -Server $vCenterServer -Confirm:$false
Write-Host "Disconnected from $vCenterServer"
