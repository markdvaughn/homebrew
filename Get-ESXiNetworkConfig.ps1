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
    - Current date used in script execution: February 25, 2025

    Version: 1.0.20
    Last Updated: February 25, 2025
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
        $vmkPortGroupNames = $vmkAdapters | ForEach-Object { $_.PortGroupName }
        
        $vSwitchData = foreach ($vSwitch in $vSwitches) {
            $security = $vSwitch | Get-SecurityPolicy
            $teaming = $vSwitch | Get-NicTeamingPolicy
            
            # Get VMkernel adapters and VLANs associated with this vSwitch
            $relatedVmk = $vmkAdapters | Where-Object { 
                $_.PortGroupName -in (Get-VirtualPortGroup -VirtualSwitch $vSwitch).Name 
            }
            $vmkList = if ($relatedVmk) {
                $vmkArray = @($relatedVmk | ForEach-Object { 
                    $vlan = (Get-VirtualPortGroup -Name $_.PortGroupName -VMHost $vmHost).VLanId
                    "$($_.Name) (VLAN $vlan)"
                })
                Write-Host "Joining vmkArray for $($vSwitch.Name): $($vmkArray -join ', ')"
                [String]::Join(', ', $vmkArray)
            } else {
                'None'
            }

            # Get virtual machine port groups (exclude VMkernel port groups)
            $vmPortGroups = Get-VirtualPortGroup -VirtualSwitch $vSwitch | Where-Object { $_.Name -notin $vmkPortGroupNames }
            $vmPgList = if ($vmPortGroups) {
                $vmPgArray = @($vmPortGroups | ForEach-Object { "$($_.Name) (VLAN $($_.VLanId))" })
                Write-Host "Joining vmPgArray for $($vSwitch.Name): $($vmPgArray -join ', ')"
                [String]::Join('<br>', $vmPgArray)  # Use <br> for line breaks in HTML
            } else {
                'None'
            }

            # Ensure arrays are not null before joining, with explicit $teaming check
            $nicList = if ($null -ne $vSwitch.Nic) { $vSwitch.Nic } else { @() }
            $activeNicList = @()
            $standbyNicList = @()
            $loadBalancing = 'N/A'
            if ($null -ne $teaming) {
                if ($null -ne $teaming.ActiveNic) { $activeNicList = $teaming.ActiveNic }
                if ($null -ne $teaming.StandbyNic) { $standbyNicList = $teaming.StandbyNic }
                $loadBalancing = $teaming.LoadBalancingPolicy
            }

            Write-Host "Joining nicList for $($vSwitch.Name): $($nicList -join ', ')"
            Write-Host "Joining activeNicList for $($vSwitch.Name): $($activeNicList -join ', ')"
            Write-Host "Joining standbyNicList for $($vSwitch.Name): $($standbyNicList -join ', ')"

            [PSCustomObject]@{
                Name = $vSwitch.Name
                Ports = $vSwitch.NumPorts
                MTU = $vSwitch.MTU
                NICs = [String]::Join(', ', $nicList)
                Promiscuous = $security.AllowPromiscuous
                ForgedTransmits = $security.ForgedTransmits
                MacChanges = $security.MacChanges
                LoadBalancing = $loadBalancing
                ActiveNICs = [String]::Join(', ', $activeNicList)
                StandbyNICs = [String]::Join(', ', $standbyNicList)
                VMkernels_VLANs = $vmkList
                VMPortGroups = $vmPgList
            }
        }
        $htmlContent.Add(($vSwitchData | ConvertTo-Html -Fragment)) | Out-Null

        # --- VMkernel Interfaces ---
        $htmlContent.Add('<h2 id="vmkernel">VMkernel Interfaces</h2>') | Out-Null
        $hostNetwork = Get-VMHostNetwork -VMHost $vmHost
        Write-Host "Debug: Host Default Gateway: $($hostNetwork.DefaultGateway)"
        $routes = $vmHost | Get-VMHostRoute
        $routeStrings = $routes | ForEach-Object { "$($_.Destination) via $($_.Gateway)" }
        Write-Host "Debug: Full Routes: $($routeStrings -join ', ')"
        $defaultGateway = ($routes | Where-Object { $_.Destination -eq '0.0.0.0' } | Select-Object -First 1).Gateway
        Write-Host "Debug: Route Default Gateway: $($defaultGateway)"
        $vmkData = foreach ($vmk in $vmkAdapters) {
            # Debug output to check raw values
            Write-Host "Debug: $($vmk.Name) IPGateway: '$($vmk.IPGateway)', ManagementEnabled: $($vmk.ManagementTrafficEnabled)"
            # Use IPGateway if set, otherwise fall back to host's default gateway from route for management
            $vmkGateway = if ($vmk.IPGateway -and $vmk.IPGateway -ne '0.0.0.0') { 
                $vmk.IPGateway 
            } elseif ($vmk.ManagementTrafficEnabled -and $defaultGateway -and $defaultGateway -ne '0.0.0.0') { 
                $defaultGateway 
            } else { 
                'N/A' 
            }

            [PSCustomObject]@{
                Name = $vmk.Name
                IP = $vmk.IP
                SubnetMask = $vmk.SubnetMask
                MAC = $vmk.Mac
                PortGroup = $vmk.PortGroupName
                Gateway = $vmkGateway
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
            
            # Get physical NICs (uplinks) from ProxySwitch
            $netSystem = Get-View -Id $vmHost.ExtensionData.ConfigManager.NetworkSystem
            $proxySwitch = $netSystem.NetworkInfo.ProxySwitch | Where-Object { $_.DvsUuid -eq $dvSwitch.ExtensionData.Uuid }
            Write-Host "Debug: ProxySwitch for $($dvSwitch.Name) found: $($null -ne $proxySwitch)"
            if ($proxySwitch) {
                Write-Host "Debug: Pnic for $($dvSwitch.Name): $($proxySwitch.Pnic -join ', ')"
            }
            
            $dvUplinks = if ($proxySwitch -and $proxySwitch.Pnic) {
                $proxySwitch.Pnic | ForEach-Object { $_.Split('-')[-1] }
            } else {
                @()
            }
            $uplinkNames = if ($dvUplinks) {
                Write-Host "Joining uplinkArray for $($dvSwitch.Name): $($dvUplinks -join ', ')"
                [String]::Join(', ', $dvUplinks)
            } else {
                'None'
            }

            # Get virtual machine port groups for distributed vSwitch
            $dvPortGroups = Get-VDPortgroup -VDSwitch $dvSwitch
            $dvPgList = if ($dvPortGroups) {
                $dvPgArray = @($dvPortGroups | ForEach-Object { "$($_.Name) (VLAN $($_.VlanConfiguration.VlanId))" })
                Write-Host "Joining dvPgArray for $($dvSwitch.Name): $($dvPgArray -join ', ')"
                [String]::Join('<br>', $dvPgArray)  # Use <br> for line breaks in HTML
            } else {
                'None'
            }

            # Get active and standby NICs from teaming policy
            $activeNicList = @()
            $standbyNicList = @()
            $loadBalancing = 'N/A'
            if ($null -ne $teaming) {
                $loadBalancing = $teaming.LoadBalancingPolicy
                if ($teaming.ActiveUplink -and $proxySwitch -and $proxySwitch.Pnic) {
                    # Map logical uplinks to physical NICs
                    $uplinkPorts = $dvSwitch.ExtensionData.Config.UplinkPortPolicy.UplinkPortName
                    $pnicList = $proxySwitch.Pnic | ForEach-Object { $_.Split('-')[-1] }
                    $uplinkMapping = @{}
                    for ($i = 0; $i -lt [Math]::Min($uplinkPorts.Count, $pnicList.Count); $i++) {
                        $uplinkMapping[$uplinkPorts[$i]] = $pnicList[$i]
                    }
                    $activeNicList = @($teaming.ActiveUplink | ForEach-Object { $uplinkMapping[$_] } | Where-Object { $_ })
                    $standbyNicList = @($teaming.StandbyUplink | ForEach-Object { $uplinkMapping[$_] } | Where-Object { $_ })
                }
                # If no explicit active/standby, assume all uplinks are active unless standby is specified
                if (-not $activeNicList -and -not $standbyNicList -and $dvUplinks) {
                    $activeNicList = $dvUplinks
                }
            }

            Write-Host "Joining activeNicList for $($dvSwitch.Name): $($activeNicList -join ', ')"
            Write-Host "Joining standbyNicList for $($dvSwitch.Name): $($standbyNicList -join ', ')"

            [PSCustomObject]@{
                Name = $dvSwitch.Name
                Ports = $dvSwitch.NumPorts
                MTU = $dvSwitch.Mtu
                NICs = $uplinkNames
                Promiscuous = $security.AllowPromiscuous
                ForgedTransmits = $security.ForgedTransmits
                MacChanges = $security.MacChanges
                LoadBalancing = $loadBalancing
                ActiveNICs = [String]::Join(', ', $activeNicList)
                StandbyNICs = [String]::Join(', ', $standbyNicList)
                VMPortGroups = $dvPgList
            }
        }
        $htmlContent.Add(($dvSwitchData | ConvertTo-Html -Fragment)) | Out-Null

        # --- DNS and Routing ---
        $htmlContent.Add('<h2 id="dnsRouting">DNS and Routing</h2>') | Out-Null
        $networkConfig = Get-VMHostNetwork -VMHost $vmHost
        $dnsServersList = if ($null -ne $networkConfig.DnsAddress) { $networkConfig.DnsAddress } else { @() }
        $staticRoutesList = if ($null -ne ($vmHost | Get-VMHostRoute)) { @($vmHost | Get-VMHostRoute | ForEach-Object { "$($_.Destination)/$($_.PrefixLength) via $($_.Gateway)" }) } else { @() }
        
        Write-Host "Joining dnsServersList: $($dnsServersList -join ', ')"
        Write-Host "Joining staticRoutesList: $($staticRoutesList -join ', ')"
        
        $dnsRoutingData = [PSCustomObject]@{
            DNSServers = [String]::Join(', ', $dnsServersList)
            StaticRoutes = [String]::Join(', ', $staticRoutesList)
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
        $ntpServersList = if ($null -ne $ntpConfig) { $ntpConfig } else { @() }
        
        Write-Host "Joining ntpServersList: $($ntpServersList -join ', ')"
        
        $ntpData = [PSCustomObject]@{
            NTPServers = [String]::Join(', ', $ntpServersList)
            Running = if ($null -ne $ntpService) { $ntpService.Running } else { 'N/A' }
            Policy = if ($null -ne $ntpService) { $ntpService.Policy } else { 'N/A' }
        }
        $htmlContent.Add(($ntpData | ConvertTo-Html -Fragment)) | Out-Null

        # --- Physical NIC Hardware Information with Network Hints ---
        $htmlContent.Add('<h2 id="physicalNics">Physical NICs</h2>') | Out-Null
        $physicalNics = Get-VMHostNetworkAdapter -VMHost $vmHost -Physical
        $standardSwitches = Get-VirtualSwitch -VMHost $vmHost -Standard
        $distributedSwitches = Get-VDSwitch -VMHost $vmHost
        
        # Pre-fetch ProxySwitch data for all distributed switches
        $netSystem = Get-View -Id $vmHost.ExtensionData.ConfigManager.NetworkSystem
        $proxySwitchMap = @{}
        foreach ($dvs in $distributedSwitches) {
            $proxySwitch = $netSystem.NetworkInfo.ProxySwitch | Where-Object { $_.DvsUuid -eq $dvs.ExtensionData.Uuid }
            if ($proxySwitch) {
                $proxySwitchMap[$dvs.ExtensionData.Uuid] = $proxySwitch.Pnic | ForEach-Object { $_.Split('-')[-1] }
            }
        }

        # Get Network Hints and Physical NIC speeds via Get-View
        $networkHints = $netSystem.QueryNetworkHint($null)
        $hintTable = @{}
        foreach ($hint in $networkHints) {
            $hintTable[$hint.Device] = $hint
        }
        $nicSpeeds = @{}
        $hostView = Get-View -Id $vmHost.Id
        foreach ($pnic in $hostView.Config.Network.Pnic) {
            $nicSpeeds[$pnic.Device] = $pnic.LinkSpeed.SpeedMb
        }

        $nicData = foreach ($nic in $physicalNics) {
            $hint = $hintTable[$nic.Name]
            $cdp = $hint.ConnectedSwitchPort
            $lldp = $hint.LLDPInfo
            
            # Combine standard and distributed vSwitches
            $vSwitchList = @(
                $standardSwitches | Where-Object { $_.Nic -contains $nic.Name } | ForEach-Object { $_.Name }
            ) + @(
                $distributedSwitches | Where-Object { 
                    $pnicList = $proxySwitchMap[$_.ExtensionData.Uuid]
                    $pnicList -and $pnicList -contains $nic.Name
                } | ForEach-Object { $_.Name }
            )
            
            Write-Host "Joining vSwitchList for $($nic.Name): $($vSwitchList -join ', ')"
            Write-Host "Debug: $($nic.Name) LinkSpeedMb: '$($nic.LinkSpeedMb)', SpeedMb from View: '$($nicSpeeds[$nic.Name])'"
            
            [PSCustomObject]@{
                Name = $nic.Name
                MAC = $nic.Mac
                LinkSpeed = if ($null -ne $nicSpeeds[$nic.Name]) { "$($nicSpeeds[$nic.Name]) Mb/s" } else { 'Unknown' }
                vSwitches = [String]::Join(', ', $vSwitchList)
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
