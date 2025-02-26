<#
.SYNOPSIS
    Documents the complete network configuration of ESXi hosts managed by a vCenter server.

.DESCRIPTION
    This script connects to a vCenter server, collects detailed network configuration data 
    for all managed ESXi hosts, and generates individual HTML reports for each host. The 
    reports include standard and distributed vSwitch configurations, VMkernel interfaces, 
    VM port groups, DNS and routing details, firewall rules, NTP settings, and physical NIC 
    information with CDP/LLDP data.

.PARAMETER None
    This script does not accept parameters. It prompts for vCenter server selection, 
    credentials, and output path interactively.

.EXAMPLE
    .\Get-ESXiNetworkConfig.ps1
    Runs the script, prompting for vCenter server selection with descriptions, credentials, 
    and output path options, then generates HTML reports in the specified directory.

.OUTPUTS
    HTML files named "NetworkConfig_<vCenterServer>_<hostname>_<timestamp>.html" for each ESXi host.

.NOTES
    - Requires VMware PowerCLI module to be installed. Install with: Install-Module -Name VMware.PowerCLI
    - Must be run with sufficient vCenter permissions to view host configurations.
    - Ignores invalid SSL certificates by default.
    - Current date used in script execution: February 26, 2025

    Version: 1.0.37
    Last Updated: February 26, 2025
#>

# --- Configuration Variables ---
# Pre-populated list of vCenter servers with descriptions (add your servers and descriptions here)
$vCenterList = @(
    @{ Server = "vc01.example.com"; Description = "Production Cluster - East Coast" },
    @{ Server = "vc02.example.com"; Description = "Production Cluster - West Coast" },
    @{ Server = "vc03.example.com"; Description = "Test/Dev Environment" }
)

# Default output path for HTML reports (modify as needed)
$defaultOutputPath = "C:\Reports\ESXiNetworkConfig"

# Script execution directory (determined at runtime)
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# --- End Configuration Variables ---

# Suppress PowerCLI welcome message for cleaner output
$PSDefaultParameterValues['Out-Default:Width'] = 200

# Set PowerCLI to ignore invalid certificates and suppress output
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null

# Generate timestamp for output file naming
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Function to validate numeric input and re-prompt if invalid
function Get-NumericChoice {
    param (
        [string]$Prompt,
        [int]$Min,
        [int]$Max
    )
    do {
        Write-Host $Prompt
        $input = Read-Host "Enter a number ($Min-$Max)"
        if ($input -match '^\d+$' -and [int]$input -ge $Min -and [int]$input -le $Max) {
            return [int]$input
        } else {
            Write-Host "Invalid input. Please enter a number between $Min and $Max."
        }
    } while ($true)
}

# Prompt user to select or enter a vCenter server
$vCenterPrompt = "Available vCenter Servers:`n"
for ($i = 0; $i -lt $vCenterList.Count; $i++) {
    $vCenterPrompt += "$($i + 1). $($vCenterList[$i].Server) - $($vCenterList[$i].Description)`n"
}
$vCenterPrompt += "$($vCenterList.Count + 1). Enter a custom vCenter server manually"

$selection = Get-NumericChoice -Prompt $vCenterPrompt -Min 1 -Max ($vCenterList.Count + 1)

if ($selection -ge 1 -and $selection -le $vCenterList.Count) {
    $vCenterServer = $vCenterList[$selection - 1].Server
    Write-Host "Selected vCenter: $vCenterServer ($($vCenterList[$selection - 1].Description))"
} else {
    $vCenterServer = Read-Host "Enter vCenter Server hostname or IP"
    Write-Host "Using custom vCenter: $vCenterServer"
}

# Prompt for credentials
$credential = Get-Credential -Message "Enter vCenter credentials for $vCenterServer"

# Prompt for output path
$outputPrompt = "Output Path Options:`n"
$outputPrompt += "1. Default path: $defaultOutputPath`n"
$outputPrompt += "2. Script execution directory: $scriptPath`n"
$outputPrompt += "3. Custom path"

$pathChoice = Get-NumericChoice -Prompt $outputPrompt -Min 1 -Max 3

switch ($pathChoice) {
    1 { 
        $outputPath = $defaultOutputPath
        Write-Host "Using default output path: $outputPath"
    }
    2 { 
        $outputPath = $scriptPath
        Write-Host "Using script execution directory: $outputPath"
    }
    3 { 
        $customPath = Read-Host "Enter custom output path"
        $outputPath = $customPath
        Write-Host "Using custom output path: $outputPath"
    }
}

# Ensure the output directory exists
if (-not (Test-Path -Path $outputPath)) {
    try {
        New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
        Write-Host "Created output directory: $outputPath"
    }
    catch {
        Write-Warning "Failed to create output directory $outputPath : $_"
        exit
    }
}

# Disconnect from any existing vCenter connections
if ($global:DefaultVIServers) {
    Disconnect-VIServer -Server * -Force -Confirm:$false
    Write-Host "Disconnected from all previous vCenter servers."
}

try {
    # Connect to the specified vCenter
    Write-Host "Connecting to $vCenterServer..."
    $viServer = Connect-VIServer -Server $vCenterServer -Credential $credential -ErrorAction Stop
    Write-Host "Connected to $vCenterServer successfully."
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
    .timestamp { font-size: 12px; font-style: italic; }
</style>
"@

# Get all ESXi hosts managed by the current vCenter
$vmHosts = Get-VMHost -Server $viServer | Sort-Object Name

foreach ($vmHost in $vmHosts) {
    Write-Host "Processing network configuration for host: $($vmHost.Name)"
    
    $reportDateTime = Get-Date -Format "MMMM dd, yyyy HH:mm:ss"
    
    $htmlContent = [System.Collections.ArrayList]::new()
    $htmlContent.Add("<html><head>$css</head><body><h1>Network Configuration - $($vmHost.Name)</h1><div class='timestamp'>$reportDateTime</div>") | Out-Null
    
    $toc = '<div class="toc"><h2>Table of Contents</h2>'
    $toc += '<a href="#vmkernel">VMkernel Interfaces</a>'
    $toc += '<a href="#vSwitch">Standard vSwitches</a>'
    $toc += '<a href="#dvSwitch">Distributed vSwitches</a>'
    $toc += '<a href="#vmPortGroups">VM Port Groups</a>'
    $toc += '<a href="#dnsRouting">DNS and Routing</a>'
    $toc += '<a href="#firewall">Firewall Rules</a>'
    $toc += '<a href="#ntp">NTP Settings</a>'
    $toc += '<a href="#physicalNics">Physical NICs</a>'
    $toc += '</div>'
    $htmlContent.Add($toc) | Out-Null

    try {
        # --- VMkernel Interfaces ---
        $htmlContent.Add('<h2 id="vmkernel">VMkernel Interfaces</h2>') | Out-Null
        $hostNetwork = Get-VMHostNetwork -VMHost $vmHost -Server $viServer
        $vmkAdapters = Get-VMHostNetworkAdapter -VMHost $vmHost -VMKernel -Server $viServer
        Write-Host "Debug: Host Default Gateway: $($hostNetwork.DefaultGateway)"
        $routes = $vmHost | Get-VMHostRoute -Server $viServer
        $routeStrings = $routes | ForEach-Object { "$($_.Destination) via $($_.Gateway)" }
        Write-Host "Debug: Full Routes: $([String]::Join(', ', $routeStrings))"
        $defaultGateway = ($routes | Where-Object { $_.Destination -eq '0.0.0.0' } | Select-Object -First 1).Gateway
        Write-Host "Debug: Route Default Gateway: $($defaultGateway)"
        $vmkData = foreach ($vmk in $vmkAdapters) {
            Write-Host "Debug: $($vmk.Name) IPGateway: '$($vmk.IPGateway)', ManagementEnabled: $($vmk.ManagementTrafficEnabled)"
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
                VLAN = (Get-VirtualPortGroup -Name $vmk.PortGroupName -VMHost $vmHost -Server $viServer).VLanId
                MTU = $vmk.MTU
            }
        }
        $htmlContent.Add(($vmkData | ConvertTo-Html -Fragment)) | Out-Null

        # --- Standard vSwitch Configuration ---
        $htmlContent.Add('<h2 id="vSwitch">Standard vSwitches</h2>') | Out-Null
        $vSwitches = Get-VirtualSwitch -VMHost $vmHost -Standard -Server $viServer
        $vmkPortGroupNames = $vmkAdapters | ForEach-Object { $_.PortGroupName }
        
        $vSwitchData = foreach ($vSwitch in $vSwitches) {
            try {
                $security = $vSwitch | Get-SecurityPolicy -Server $viServer
                $teaming = $vSwitch | Get-NicTeamingPolicy -Server $viServer
                
                $relatedVmk = $vmkAdapters | Where-Object { 
                    $_.PortGroupName -in (Get-VirtualPortGroup -VirtualSwitch $vSwitch -Server $viServer).Name 
                }
                $vmkList = if ($relatedVmk) {
                    $vmkArray = @($relatedVmk | ForEach-Object { 
                        $vlan = (Get-VirtualPortGroup -Name $_.PortGroupName -VMHost $vmHost -Server $viServer).VLanId
                        "$($_.Name) (VLAN $vlan)"
                    })
                    Write-Host "Joining vmkArray for $($vSwitch.Name): $([String]::Join(', ', $vmkArray))"
                    [String]::Join(', ', $vmkArray)
                } else {
                    'None'
                }

                Write-Host "Debug: Teaming for $($vSwitch.Name) is null: $($teaming -eq $null)"
                if ($teaming) {
                    Write-Host "Debug: ActiveNic is null: $($teaming.ActiveNic -eq $null)"
                    Write-Host "Debug: StandbyNic is null: $($teaming.StandbyNic -eq $null)"
                    Write-Host "Debug: LoadBalancingPolicy: $($teaming.LoadBalancingPolicy -or 'null')"
                }

                $nicList = if ($null -ne $vSwitch.Nic) { @($vSwitch.Nic) } else { @() }
                $activeNicList = if ($teaming -and $null -ne $teaming.ActiveNic) { @($teaming.ActiveNic) } else { @() }
                $standbyNicList = if ($teaming -and $null -ne $teaming.StandbyNic) { @($teaming.StandbyNic) } else { @() }
                $loadBalancing = if ($teaming -and $null -ne $teaming.LoadBalancingPolicy) { $teaming.LoadBalancingPolicy } else { 'N/A' }

                Write-Host "Debug: nicList before join: $(if ($nicList) { $nicList -join ', ' } else { 'empty' })"
                Write-Host "Debug: activeNicList before join: $(if ($activeNicList) { $activeNicList -join ', ' } else { 'empty' })"
                Write-Host "Debug: standbyNicList before join: $(if ($standbyNicList) { $standbyNicList -join ', ' } else { 'empty' })"

                $nicString = [String]::Join(', ', @($nicList))
                $activeNicString = [String]::Join(', ', @($activeNicList))
                $standbyNicString = [String]::Join(', ', @($standbyNicList))

                Write-Host "Joining nicList for $($vSwitch.Name): $nicString"
                Write-Host "Joining activeNicList for $($vSwitch.Name): $activeNicString"
                Write-Host "Joining standbyNicList for $($vSwitch.Name): $standbyNicString"

                [PSCustomObject]@{
                    Name = $vSwitch.Name
                    Ports = $vSwitch.NumPorts
                    MTU = $vSwitch.MTU
                    NICs = $nicString
                    Promiscuous = $security.AllowPromiscuous
                    ForgedTransmits = $security.ForgedTransmits
                    MacChanges = $security.MacChanges
                    LoadBalancing = $loadBalancing
                    ActiveNICs = $activeNicString
                    StandbyNICs = $standbyNicString
                    VMkernels_VLANs = $vmkList
                }
            }
            catch {
                Write-Warning "Error processing vSwitch $($vSwitch.Name) on $($vmHost.Name): $_"
                [PSCustomObject]@{
                    Name = $vSwitch.Name
                    Ports = 'N/A'
                    MTU = 'N/A'
                    NICs = 'N/A'
                    Promiscuous = 'N/A'
                    ForgedTransmits = 'N/A'
                    MacChanges = 'N/A'
                    LoadBalancing = 'N/A'
                    ActiveNICs = 'N/A'
                    StandbyNICs = 'N/A'
                    VMkernels_VLANs = 'N/A'
                }
            }
        }
        $htmlContent.Add(($vSwitchData | ConvertTo-Html -Fragment)) | Out-Null

        # --- Distributed vSwitch Configuration ---
        $htmlContent.Add('<h2 id="dvSwitch">Distributed vSwitches</h2>') | Out-Null
        $dvSwitches = Get-VDSwitch -VMHost $vmHost -Server $viServer
        $dvSwitchData = if ($dvSwitches.Count -gt 0) {
            foreach ($dvSwitch in $dvSwitches) {
                $security = $dvSwitch | Get-VDSecurityPolicy -Server $viServer
                $teaming = $dvSwitch | Get-VDUplinkTeamingPolicy -Server $viServer
                
                $netSystem = Get-View -Id $vmHost.ExtensionData.ConfigManager.NetworkSystem -Server $viServer
                $proxySwitch = $netSystem.NetworkInfo.ProxySwitch | Where-Object { $_.DvsUuid -eq $dvSwitch.ExtensionData.Uuid }
                Write-Host "Debug: ProxySwitch for $($dvSwitch.Name) found: $($null -ne $proxySwitch)"
                if ($proxySwitch) {
                    Write-Host "Debug: Pnic for $($dvSwitch.Name): $([String]::Join(', ', $proxySwitch.Pnic))"
                }
                
                $dvUplinks = if ($proxySwitch -and $proxySwitch.Pnic) {
                    $proxySwitch.Pnic | ForEach-Object { $_.Split('-')[-1] }
                } else {
                    @()
                }
                $uplinkNames = if ($dvUplinks) {
                    Write-Host "Joining uplinkArray for $($dvSwitch.Name): $([String]::Join(', ', $dvUplinks))"
                    [String]::Join(', ', $dvUplinks)
                } else {
                    'None'
                }

                $activeNicList = @()
                $standbyNicList = @()
                $loadBalancing = 'N/A'
                if ($null -ne $teaming) {
                    $loadBalancing = $teaming.LoadBalancingPolicy
                    if ($teaming.ActiveUplink -and $proxySwitch -and $proxySwitch.Pnic) {
                        $uplinkPorts = $dvSwitch.ExtensionData.Config.UplinkPortPolicy.UplinkPortName
                        $pnicList = $proxySwitch.Pnic | ForEach-Object { $_.Split('-')[-1] }
                        $uplinkMapping = @{}
                        for ($i = 0; $i -lt [Math]::Min($uplinkPorts.Count, $pnicList.Count); $i++) {
                            $uplinkMapping[$uplinkPorts[$i]] = $pnicList[$i]
                        }
                        $activeNicList = @($teaming.ActiveUplink | ForEach-Object { $uplinkMapping[$_] } | Where-Object { $_ })
                        $standbyNicList = @($teaming.StandbyUplink | ForEach-Object { $uplinkMapping[$_] } | Where-Object { $_ })
                    }
                    if (-not $activeNicList -and -not $standbyNicList -and $dvUplinks) {
                        $activeNicList = $dvUplinks
                    }
                }

                Write-Host "Joining activeNicList for $($dvSwitch.Name): $([String]::Join(', ', $activeNicList))"
                Write-Host "Joining standbyNicList for $($dvSwitch.Name): $([String]::Join(', ', $standbyNicList))"

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
                }
            }
        } else {
            [PSCustomObject]@{
                Name = 'N/A'
                Ports = ''
                MTU = ''
                NICs = ''
                Promiscuous = ''
                ForgedTransmits = ''
                MacChanges = ''
                LoadBalancing = ''
                ActiveNICs = ''
                StandbyNICs = ''
            }
        }
        $htmlContent.Add(($dvSwitchData | ConvertTo-Html -Fragment)) | Out-Null

        # --- VM Port Groups ---
        $htmlContent.Add('<h2 id="vmPortGroups">VM Port Groups</h2>') | Out-Null
        $vmPortGroupData = [System.Collections.ArrayList]::new()

        $standardPortGroups = [System.Collections.ArrayList]::new()
        foreach ($vSwitch in $vSwitches) {
            $vmPortGroups = Get-VirtualPortGroup -VirtualSwitch $vSwitch -Server $viServer | Where-Object { $_.Name -notin $vmkPortGroupNames }
            foreach ($pg in $vmPortGroups) {
                $standardPortGroups.Add([PSCustomObject]@{
                    Name = $pg.Name
                    'VLAN ID' = $pg.VLanId
                    'Associated vSwitch' = $vSwitch.Name
                    SwitchType = 'Standard'
                }) | Out-Null
            }
        }

        $distributedPortGroups = [System.Collections.ArrayList]::new()
        foreach ($dvSwitch in $dvSwitches) {
            $dvPortGroups = Get-VDPortgroup -VDSwitch $dvSwitch -Server $viServer
            foreach ($pg in $dvPortGroups) {
                $distributedPortGroups.Add([PSCustomObject]@{
                    Name = $pg.Name
                    'VLAN ID' = $pg.VlanConfiguration.VlanId
                    'Associated vSwitch' = $dvSwitch.Name
                    SwitchType = 'Distributed'
                }) | Out-Null
            }
        }

        if ($standardPortGroups.Count -gt 0) {
            $sortedStandard = @($standardPortGroups | Sort-Object 'Associated vSwitch', 'VLAN ID')
            $vmPortGroupData.AddRange($sortedStandard)
        }

        if ($distributedPortGroups.Count -gt 0) {
            $sortedDistributed = @($distributedPortGroups | Sort-Object 'Associated vSwitch', 'VLAN ID')
            $vmPortGroupData.AddRange($sortedDistributed)
        }

        if ($vmPortGroupData.Count -eq 0) {
            $vmPortGroupData.Add([PSCustomObject]@{
                Name = 'None'
                'VLAN ID' = 'N/A'
                'Associated vSwitch' = 'N/A'
                SwitchType = 'N/A'
            }) | Out-Null
        }

        $htmlContent.Add(($vmPortGroupData | Select-Object Name, 'VLAN ID', 'Associated vSwitch' | ConvertTo-Html -Fragment)) | Out-Null

        # --- DNS and Routing ---
        $htmlContent.Add('<h2 id="dnsRouting">DNS and Routing</h2>') | Out-Null
        $networkConfig = Get-VMHostNetwork -VMHost $vmHost -Server $viServer
        $dnsServersList = if ($null -ne $networkConfig.DnsAddress) { $networkConfig.DnsAddress } else { @() }
        $staticRoutesList = if ($null -ne ($vmHost | Get-VMHostRoute -Server $viServer)) { @($vmHost | Get-VMHostRoute -Server $viServer | ForEach-Object { "$($_.Destination)/$($_.PrefixLength) via $($_.Gateway)" }) } else { @() }
        
        Write-Host "Joining dnsServersList: $([String]::Join(', ', $dnsServersList))"
        Write-Host "Joining staticRoutesList: $([String]::Join(', ', $staticRoutesList))"
        
        $dnsRoutingData = [PSCustomObject]@{
            DNSServers = [String]::Join(', ', $dnsServersList)
            StaticRoutes = [String]::Join(', ', $staticRoutesList)
        }
        $htmlContent.Add(($dnsRoutingData | ConvertTo-Html -Fragment)) | Out-Null

        # --- Firewall Rules ---
        $htmlContent.Add('<h2 id="firewall">Firewall Rules</h2>') | Out-Null
        $firewallRules = Get-VMHostFirewallException -VMHost $vmHost -Server $viServer | Where-Object { $_.Enabled }
        $firewallData = $firewallRules | Select-Object Name, Enabled, Protocol, @{N='Port';E={$_.Port}}
        $htmlContent.Add(($firewallData | ConvertTo-Html -Fragment)) | Out-Null

        # --- NTP Settings ---
        $htmlContent.Add('<h2 id="ntp">NTP Settings</h2>') | Out-Null
        $ntpConfig = Get-VMHostNtpServer -VMHost $vmHost -Server $viServer
        $ntpService = Get-VMHostService -VMHost $vmHost -Server $viServer | Where-Object { $_.Key -eq 'ntpd' }
        $ntpServersList = if ($null -ne $ntpConfig) { $ntpConfig } else { @() }
        
        Write-Host "Joining ntpServersList: $([String]::Join(', ', $ntpServersList))"
        
        $ntpData = [PSCustomObject]@{
            NTPServers = [String]::Join(', ', $ntpServersList)
            Running = if ($null -ne $ntpService) { $ntpService.Running } else { 'N/A' }
            Policy = if ($null -ne $ntpService) { $ntpService.Policy } else { 'N/A' }
        }
        $htmlContent.Add(($ntpData | ConvertTo-Html -Fragment)) | Out-Null

        # --- Physical NIC Hardware Information with Network Hints ---
        $htmlContent.Add('<h2 id="physicalNics">Physical NICs</h2>') | Out-Null
        $physicalNics = Get-VMHostNetworkAdapter -VMHost $vmHost -Physical -Server $viServer
        $standardSwitches = Get-VirtualSwitch -VMHost $vmHost -Standard -Server $viServer
        $distributedSwitches = Get-VDSwitch -VMHost $vmHost -Server $viServer
        
        $netSystem = Get-View -Id $vmHost.ExtensionData.ConfigManager.NetworkSystem -Server $viServer
        $proxySwitchMap = @{}
        foreach ($dvs in $distributedSwitches) {
            $proxySwitch = $netSystem.NetworkInfo.ProxySwitch | Where-Object { $_.DvsUuid -eq $dvs.ExtensionData.Uuid }
            if ($proxySwitch) {
                $proxySwitchMap[$dvs.ExtensionData.Uuid] = $proxySwitch.Pnic | ForEach-Object { $_.Split('-')[-1] }
            }
        }

        $networkHints = $netSystem.QueryNetworkHint($null)
        $hintTable = @{}
        foreach ($hint in $networkHints) {
            $hintTable[$hint.Device] = $hint
        }
        $nicSpeeds = @{}
        $hostView = Get-View -Id $vmHost.Id -Server $viServer
        foreach ($pnic in $hostView.Config.Network.Pnic) {
            $nicSpeeds[$pnic.Device] = $pnic.LinkSpeed.SpeedMb
        }

        $nicData = foreach ($nic in $physicalNics) {
            $hint = $hintTable[$nic.Name]
            $cdp = $hint.ConnectedSwitchPort
            $lldp = $hint.LLDPInfo
            
            $vSwitchList = @(
                $standardSwitches | Where-Object { $_.Nic -contains $nic.Name } | ForEach-Object { $_.Name }
            ) + @(
                $distributedSwitches | Where-Object { 
                    $pnicList = $proxySwitchMap[$_.ExtensionData.Uuid]
                    $pnicList -and $pnicList -contains $nic.Name
                } | ForEach-Object { $_.Name }
            )
            
            Write-Host "Joining vSwitchList for $($nic.Name): $([String]::Join(', ', $vSwitchList))"
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

        # Write to file with vCenter server name included in the specified output path
        $fileName = Join-Path -Path $outputPath -ChildPath "NetworkConfig_${vCenterServer}_$($vmHost.Name)_$timestamp.html"
        $htmlContent | Out-File -FilePath $fileName -Encoding UTF8
        Write-Host "Generated report: $fileName"
    }
    catch {
        Write-Warning "Error processing $($vmHost.Name): $_"
    }
}

# Disconnect from the vCenter server
Disconnect-VIServer -Server $vCenterServer -Force -Confirm:$false
Write-Host "Disconnected from $vCenterServer"
