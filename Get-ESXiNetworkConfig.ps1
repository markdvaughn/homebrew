<#
.SYNOPSIS
    Documents the complete network configuration of ESXi hosts managed by a vCenter server.

.DESCRIPTION
    This script connects to a vCenter server, collects detailed network configuration data 
    for all managed ESXi hosts or hosts in a selected cluster, and generates individual HTML 
    reports for each host. The reports include standard and distributed vSwitch configurations, 
    VMkernel interfaces, VM port groups, DNS and routing details, firewall rules, NTP settings, 
    and physical NIC information with CDP/LLDP data.

.PARAMETER None
    This script does not accept parameters. It prompts for vCenter server selection, 
    credentials, output path, and cluster selection interactively.

.EXAMPLE
    .\Get-ESXiNetworkConfig.ps1
    Runs the script, prompting for vCenter server selection with descriptions, credentials, 
    output path, and cluster selection, then generates HTML reports in the specified directory.

.OUTPUTS
    HTML files named "NetworkConfig_<vCenterServer>_<hostname>_<timestamp>.html" for each ESXi host.

.NOTES
    - Requires VMware PowerCLI module to be installed. Install with: Install-Module -Name VMware.PowerCLI
    - Must be run with sufficient vCenter permissions to view host configurations.
    - Ignores invalid SSL certificates by default.
    - Current date used in script execution: February 27, 2025

    Version: 1.0.68
    Last Updated: February 27, 2025

.VERSION HISTORY
    1.0.24 - February 25, 2025
        - Initial version provided by user with detailed ESXi network configuration reporting.
    # [Previous versions 1.0.25 to 1.0.67 omitted for brevity, see prior script for full history]
    1.0.67 - February 27, 2025
        - Restored VLAN values for vDS VMkernels with debug output, but repetition persisted due to string conversion.
    1.0.68 - February 27, 2025
        - Fixed VLAN repetition for vDS VMkernels by avoiding direct string casting of $vlanConfig.VlanId and using the integer value directly.
#>

# --- Configuration Variables ---
$vCenterList = @(
    @{ Server = "vc01.example.com"; Description = "Production Cluster - East Coast" },
    @{ Server = "vc02.example.com"; Description = "Production Cluster - West Coast" },
    @{ Server = "vc03.example.com"; Description = "Test/Dev Environment" }
)
$defaultOutputPath = "C:\Reports\ESXiNetworkConfig"
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptName = $MyInvocation.MyCommand.Name
$scriptVersion = "1.0.68"
# --- End Configuration Variables ---

$PSDefaultParameterValues['Out-Default:Width'] = 200
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

function Get-NumericChoice {
    param (
        [string]$Prompt,
        [int]$Min,
        [int]$Max
    )
    do {
        $lines = $Prompt.Split("`n")
        Write-Host $lines[0] -ForegroundColor Cyan
        for ($i = 1; $i -lt $lines.Count; $i++) {
            if ($lines[$i]) { Write-Host $lines[$i] -ForegroundColor White }
        }
        $input = Read-Host "Enter a number ($Min-$Max)"
        if ($input -match '^\d+$' -and [int]$input -ge $Min -and [int]$input -le $Max) {
            return [int]$input
        } else {
            Write-Host "Invalid input. Please enter a number between $Min and $Max." -ForegroundColor Red
        }
    } while ($true)
}

$vCenterPrompt = "Available vCenter Servers:`n"
for ($i = 0; $i -lt $vCenterList.Count; $i++) {
    $vCenterPrompt += "$($i + 1). $($vCenterList[$i].Server) - $($vCenterList[$i].Description)`n"
}
$vCenterPrompt += "$($vCenterList.Count + 1). Enter a custom vCenter server manually"
$selection = Get-NumericChoice -Prompt $vCenterPrompt -Min 1 -Max ($vCenterList.Count + 1)

if ($selection -ge 1 -and $selection -le $vCenterList.Count) {
    $vCenterServer = $vCenterList[$selection - 1].Server
    Write-Host "Selected vCenter: $vCenterServer ($($vCenterList[$selection - 1].Description))" -ForegroundColor Green
} else {
    $vCenterServer = Read-Host "Enter vCenter Server hostname or IP"
    Write-Host "Using custom vCenter: $vCenterServer" -ForegroundColor Green
}

$credential = Get-Credential -Message "Enter vCenter credentials for $vCenterServer"
try {
    Write-Host "Connecting to $vCenterServer..." -ForegroundColor White
    $viServer = Connect-VIServer -Server $vCenterServer -Credential $credential -ErrorAction Stop
    Write-Host "Connected to $vCenterServer successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to $vCenterServer : $_" -ForegroundColor Red
    exit
}

$clusters = Get-Cluster -Server $viServer | Sort-Object Name
$clusterPrompt = "Available VMware Clusters:`n"
for ($i = 0; $i -lt $clusters.Count; $i++) {
    $clusterPrompt += "$($i + 1). $($clusters[$i].Name)`n"
}
$clusterPrompt += "$($clusters.Count + 1). ALL clusters"
$clusterSelection = Get-NumericChoice -Prompt $clusterPrompt -Min 1 -Max ($clusters.Count + 1)

if ($clusterSelection -eq ($clusters.Count + 1)) {
    Write-Host "Selected: ALL clusters" -ForegroundColor Green
    $vmHosts = Get-VMHost -Server $viServer | Sort-Object Name
} else {
    $selectedCluster = $clusters[$clusterSelection - 1]
    Write-Host "Selected cluster: $($selectedCluster.Name)" -ForegroundColor Green
    $vmHosts = Get-VMHost -Location $selectedCluster -Server $viServer | Sort-Object Name
}

$outputPrompt = "Output Path Options:`n1. Default path: $defaultOutputPath`n2. Script execution directory: $scriptPath`n3. Custom path"
$pathChoice = Get-NumericChoice -Prompt $outputPrompt -Min 1 -Max 3
switch ($pathChoice) {
    1 { $outputPath = $defaultOutputPath; Write-Host "Using default output path: $outputPath" -ForegroundColor Green }
    2 { $outputPath = $scriptPath; Write-Host "Using script execution directory: $outputPath" -ForegroundColor Green }
    3 { $outputPath = Read-Host "Enter custom output path"; Write-Host "Using custom output path: $outputPath" -ForegroundColor Green }
}

if (-not (Test-Path -Path $outputPath)) {
    try {
        New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
        Write-Host "Created output directory: $outputPath" -ForegroundColor Green
    } catch {
        Write-Host "Failed to create output directory $outputPath : $_" -ForegroundColor Red
        exit
    }
}

if ($global:DefaultVIServers | Where-Object { $_.Name -ne $vCenterServer }) {
    Disconnect-VIServer -Server ($global:DefaultVIServers | Where-Object { $_.Name -ne $vCenterServer }) -Force -Confirm:$false
    Write-Host "Disconnected from all previous vCenter servers except $vCenterServer." -ForegroundColor Green
}

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
    .footer { font-size: 10px; color: #666; text-align: center; margin-top: 20px; }
</style>
"@

foreach ($vmHost in $vmHosts) {
    Write-Host "Processing network configuration for host: $($vmHost.Name)" -ForegroundColor White
    
    $reportDateTime = Get-Date -Format "MMMM dd, yyyy HH:mm:ss"
    $htmlContent = [System.Collections.ArrayList]::new()
    $htmlContent.Add("<html><head>$css</head><body><h1>Network Configuration - $($vmHost.Name)</h1><div class='timestamp'>Generated on: $reportDateTime</div>") | Out-Null
    
    $toc = '<div class="toc"><h2>Table of Contents</h2>'
    $toc += '<a href="#vmkernel">VMkernel Interfaces</a><a href="#vSwitch">Standard vSwitches</a><a href="#dvSwitch">Distributed vSwitches</a><a href="#vmPortGroups">VM Port Groups</a><a href="#dnsRouting">DNS and Routing</a><a href="#firewall">Firewall Rules</a><a href="#ntp">NTP Settings</a><a href="#physicalNics">Physical NICs</a>'
    $htmlContent.Add($toc + '</div>') | Out-Null

    try {
        # --- VMkernel Interfaces ---
        $htmlContent.Add('<h2 id="vmkernel">VMkernel Interfaces</h2>') | Out-Null
        $hostNetwork = Get-VMHostNetwork -VMHost $vmHost -Server $viServer
        $vmkAdapters = Get-VMHostNetworkAdapter -VMHost $vmHost -VMKernel -Server $viServer
        $defaultGatewayDisplay = if ($hostNetwork.DefaultGateway -eq $null) { 'N/A' } else { $hostNetwork.DefaultGateway }
        Write-Host "Host Default Gateway: $defaultGatewayDisplay" -ForegroundColor White
        $routes = $vmHost | Get-VMHostRoute -Server $viServer
        $defaultGateway = ($routes | Where-Object { $_.Destination -eq '0.0.0.0' } | Select-Object -First 1).Gateway
        
        $vmkData = foreach ($vmk in $vmkAdapters) {
            $vmkGateway = if ($vmk.IPGateway -and $vmk.IPGateway -ne '0.0.0.0') { 
                $vmk.IPGateway 
            } elseif ($vmk.ManagementTrafficEnabled -and $defaultGateway -and $defaultGateway -ne '0.0.0.0') { 
                $defaultGateway 
            } else { 
                'N/A' 
            }
            
            $vlanId = 'N/A'
            if ($vmk.PortGroupName) {
                $portGroup = Get-VirtualPortGroup -Name $vmk.PortGroupName -VMHost $vmHost -Server $viServer -Standard -ErrorAction SilentlyContinue
                if ($portGroup) {
                    $vlanId = $portGroup.VLanId  # Integer for standard vSwitches
                } else {
                    $portGroup = Get-VDPortgroup -Name $vmk.PortGroupName -Server $viServer -ErrorAction SilentlyContinue
                    if ($portGroup) {
                        if ($portGroup.VlanConfiguration) {
                            $vlanConfig = $portGroup.VlanConfiguration
                            if ($vlanConfig.VlanId -ne $null) {
                                $vlanId = "$($vlanConfig.VlanId)"  # Direct integer to string, no cast
                            } elseif ($vlanConfig.VlanRange -and $vlanConfig.VlanRange.Count -gt 0) {
                                $vlanId = "Trunk ($($vlanConfig.VlanRange[0].Start)-$($vlanConfig.VlanRange[0].End))"
                            } else {
                                $vlanId = '0'
                            }
                        } else {
                            $pgView = Get-View -Id $portGroup.ExtensionData.MoRef -Server $viServer -ErrorAction SilentlyContinue
                            if ($pgView -and $pgView.Config.DefaultPortConfig.Vlan) {
                                $vlanConfig = $pgView.Config.DefaultPortConfig.Vlan
                                if ($vlanConfig -is [VMware.Vim.VmwareDistributedVirtualSwitchVlanIdSpec] -and $vlanConfig.VlanId -ne $null) {
                                    $vlanId = "$($vlanConfig.VlanId)"  # Direct integer to string, no cast
                                } elseif ($vlanConfig -is [VMware.Vim.VmwareDistributedVirtualSwitchTrunkVlanSpec] -and $vlanConfig.VlanId.Count -gt 0) {
                                    $vlanId = "Trunk ($($vlanConfig.VlanId[0].Start)-$($vlanConfig.VlanId[0].End))"
                                } else {
                                    $vlanId = '0'
                                }
                            }
                        }
                    }
                }
            }
            
            Write-Host "$($vmk.Name) IP: $($vmk.IP), Gateway: $vmkGateway, VLAN: $vlanId" -ForegroundColor White

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
                VLAN = $vlanId
                MTU = $vmk.MTU
            }
        }
        $htmlContent.Add(($vmkData | ConvertTo-Html -Fragment)) | Out-Null

        # --- Standard vSwitches ---
        $htmlContent.Add('<h2 id="vSwitch">Standard vSwitches</h2>') | Out-Null
        $vSwitches = Get-VirtualSwitch -VMHost $vmHost -Standard -Server $viServer
        $vmkPortGroupNames = $vmkAdapters | ForEach-Object { $_.PortGroupName }
        
        $vSwitchData = foreach ($vSwitch in $vSwitches) {
            try {
                $security = $vSwitch | Get-SecurityPolicy -Server $viServer
                $teaming = $vSwitch | Get-NicTeamingPolicy -Server $viServer
                
                # VMkernel list
                $relatedVmk = $vmkAdapters | Where-Object { $_.PortGroupName -in (Get-VirtualPortGroup -VirtualSwitch $vSwitch -Server $viServer).Name }
                $vmkArray = if ($relatedVmk) {
                    @($relatedVmk | ForEach-Object { 
                        $portGroup = Get-VirtualPortGroup -Name $_.PortGroupName -VMHost $vmHost -Server $viServer -Standard -ErrorAction SilentlyContinue
                        $vlanId = if ($portGroup) { $portGroup.VLanId } else { 'N/A' }
                        "$($_.Name) (VLAN $vlanId)"
                    })
                } else {
                    @('None')
                }
                $vmkList = if ($vmkArray) { [String]::Join(', ', $vmkArray) } else { 'None' }
                
                # NIC lists with explicit null handling
                $nicList = if ($vSwitch.Nic) { @($vSwitch.Nic) } else { @() }
                $activeNicList = if ($teaming -and $teaming.ActiveNic) { @($teaming.ActiveNic) } else { @() }
                $standbyNicList = if ($teaming -and $teaming.StandbyNic) { @($teaming.StandbyNic) } else { @() }
                $loadBalancing = if ($teaming -and $teaming.LoadBalancingPolicy) { $teaming.LoadBalancingPolicy } else { 'N/A' }

                # Pre-join NIC lists into strings with null safety
                $nicString = if ($nicList) { [String]::Join(', ', $nicList) } else { '' }
                $activeNicString = if ($activeNicList) { [String]::Join(', ', $activeNicList) } else { '' }
                $standbyNicString = if ($standbyNicList) { [String]::Join(', ', $standbyNicList) } else { '' }

                # Debug output before Write-Host
                Write-Host "Debug Before Write - $($vSwitch.Name): NICs=[$nicString], ActiveNICs=[$activeNicString], StandbyNICs=[$standbyNicString], VMkernels=[$vmkList]" -ForegroundColor Yellow
                
                Write-Host "$($vSwitch.Name) NICs: $nicString, VMkernels: $vmkList" -ForegroundColor White

                # Debug output right before PSCustomObject
                Write-Host "Debug Before PSCustomObject - $($vSwitch.Name): NICs=[$nicString], ActiveNICs=[$activeNicString], StandbyNICs=[$standbyNicString], VMkernels=[$vmkList]" -ForegroundColor Yellow

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
            } catch {
                Write-Host "Error processing vSwitch $($vSwitch.Name) on $($vmHost.Name): $_" -ForegroundColor Red
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

        # --- Distributed vSwitches ---
        $htmlContent.Add('<h2 id="dvSwitch">Distributed vSwitches</h2>') | Out-Null
        $dvSwitches = Get-VDSwitch -VMHost $vmHost -Server $viServer
        $dvSwitchData = if ($dvSwitches.Count -gt 0) {
            foreach ($dvSwitch in $dvSwitches) {
                $security = $dvSwitch | Get-VDSecurityPolicy -Server $viServer
                $teaming = $dvSwitch | Get-VDUplinkTeamingPolicy -Server $viServer
                
                $netSystem = Get-View -Id $vmHost.ExtensionData.ConfigManager.NetworkSystem -Server $viServer
                $proxySwitch = $netSystem.NetworkInfo.ProxySwitch | Where-Object { $_.DvsUuid -eq $dvSwitch.ExtensionData.Uuid }
                $dvUplinks = if ($proxySwitch -and $proxySwitch.Pnic) { $proxySwitch.Pnic | ForEach-Object { $_.Split('-')[-1] } } else { @() }
                $uplinkNames = if ($dvUplinks) { [String]::Join(', ', $dvUplinks) } else { 'None' }
                
                $dvPortGroups = Get-VDPortgroup -VDSwitch $dvSwitch -Server $viServer
                $relatedVmk = $vmkAdapters | Where-Object { $_.PortGroupName -in $dvPortGroups.Name }
                $vmkList = if ($relatedVmk) {
                    $vmkArray = @($relatedVmk | ForEach-Object { 
                        $portGroup = Get-VDPortgroup -Name $_.PortGroupName -Server $viServer -ErrorAction SilentlyContinue
                        $vlanId = 'N/A'
                        if ($portGroup) {
                            if ($portGroup.VlanConfiguration) {
                                $vlanConfig = $portGroup.VlanConfiguration
                                if ($vlanConfig.VlanId -ne $null) {
                                    $vlanId = "$($vlanConfig.VlanId)"  # Direct integer to string, no cast
                                } elseif ($vlanConfig.VlanRange -and $vlanConfig.VlanRange.Count -gt 0) {
                                    $vlanId = "Trunk ($($vlanConfig.VlanRange[0].Start)-$($vlanConfig.VlanRange[0].End))"
                                } else {
                                    $vlanId = '0'
                                }
                            } else {
                                $pgView = Get-View -Id $portGroup.ExtensionData.MoRef -Server $viServer -ErrorAction SilentlyContinue
                                if ($pgView -and $pgView.Config.DefaultPortConfig.Vlan) {
                                    $vlanConfig = $pgView.Config.DefaultPortConfig.Vlan
                                    if ($vlanConfig -is [VMware.Vim.VmwareDistributedVirtualSwitchVlanIdSpec] -and $vlanConfig.VlanId -ne $null) {
                                        $vlanId = "$($vlanConfig.VlanId)"  # Direct integer to string, no cast
                                    } elseif ($vlanConfig -is [VMware.Vim.VmwareDistributedVirtualSwitchTrunkVlanSpec] -and $vlanConfig.VlanId.Count -gt 0) {
                                        $vlanId = "Trunk ($($vlanConfig.VlanId[0].Start)-$($vlanConfig.VlanId[0].End))"
                                    } else {
                                        $vlanId = '0'
                                    }
                                }
                            }
                        }
                        "$($_.Name) (VLAN $vlanId)"
                    })
                    [String]::Join(', ', $vmkArray)
                } else {
                    'None'
                }
                Write-Host "$($dvSwitch.Name) NICs: $uplinkNames, VMkernels: $vmkList" -ForegroundColor White

                $activeNicList = @()
                $standbyNicList = @()
                $loadBalancing = 'N/A'
                if ($teaming) {
                    $loadBalancing = if ($teaming.LoadBalancingPolicy) { $teaming.LoadBalancingPolicy } else { 'N/A' }
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
                    VMkernels_VLANs = $vmkList
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
                VMkernels_VLANs = 'N/A'
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
        foreach ($dvSwitch in $vSwitches) {
            $dvPortGroups = Get-VDPortgroup -VDSwitch $dvSwitch -Server $viServer -ErrorAction SilentlyContinue
            foreach ($pg in $dvPortGroups) {
                $vlanId = $pg.VlanId
                if ($vlanId -eq $null -or $vlanId -eq 0) {
                    $pgView = Get-View -Id $pg.ExtensionData.MoRef -Server $viServer
                    $vlanConfig = $pgView.Config.DefaultPortConfig.Vlan
                    $vlanId = if ($vlanConfig -is [VMware.Vim.VmwareDistributedVirtualSwitchVlanIdSpec]) {
                        $vlanConfig.VlanId
                    } elseif ($vlanConfig -is [VMware.Vim.VmwareDistributedVirtualSwitchTrunkVlanSpec]) {
                        "Trunk ($($vlanConfig.VlanId[0].Start)-$($vlanConfig.VlanId[0].End))"
                    } else {
                        '0'
                    }
                }
                $distributedPortGroups.Add([PSCustomObject]@{
                    Name = $pg.Name
                    'VLAN ID' = $vlanId
                    'Associated vSwitch' = $dvSwitch.Name
                    SwitchType = 'Distributed'
                }) | Out-Null
            }
        }
        if ($standardPortGroups.Count -gt 0) {
            $sortedStandard = @($standardPortGroups | Sort-Object 'Associated vSwitch', 'VLAN ID')
            $vmPortGroupData.AddRange($sortedStandard)
            Write-Host "Standard Port Groups: $(($sortedStandard | ForEach-Object { $_.Name }) -join ', ')" -ForegroundColor White
        }
        if ($distributedPortGroups.Count -gt 0) {
            $sortedDistributed = @($distributedPortGroups | Sort-Object 'Associated vSwitch', 'VLAN ID')
            $vmPortGroupData.AddRange($sortedDistributed)
            Write-Host "Distributed Port Groups: $(($sortedDistributed | ForEach-Object { $_.Name }) -join ', ')" -ForegroundColor White
        }
        if ($vmPortGroupData.Count -eq 0) {
            $vmPortGroupData.Add([PSCustomObject]@{
                Name = 'None'
                'VLAN ID' = 'N/A'
                'Associated vSwitch' = 'N/A'
                SwitchType = 'N/A'
            }) | Out-Null
            Write-Host "No VM Port Groups found." -ForegroundColor White
        }
        $htmlContent.Add(($vmPortGroupData | Select-Object Name, 'VLAN ID', 'Associated vSwitch' | ConvertTo-Html -Fragment)) | Out-Null

        # --- DNS and Routing ---
        $htmlContent.Add('<h2 id="dnsRouting">DNS and Routing</h2>') | Out-Null
        $networkConfig = Get-VMHostNetwork -VMHost $vmHost -Server $viServer
        $dnsServersList = if ($null -ne $networkConfig.DnsAddress) { $networkConfig.DnsAddress } else { @() }
        $staticRoutesList = if ($null -ne ($vmHost | Get-VMHostRoute -Server $viServer)) { @($vmHost | Get-VMHostRoute -Server $viServer | ForEach-Object { "$($_.Destination)/$($_.PrefixLength) via $($_.Gateway)" }) } else { @() }
        Write-Host "DNS Servers: $([String]::Join(', ', $dnsServersList))" -ForegroundColor White
        if ($staticRoutesList.Count -gt 0) { Write-Host "Static Routes: $(($staticRoutesList | ForEach-Object { $_.Name }) -join ', ')" -ForegroundColor White }
        $dnsRoutingData = [PSCustomObject]@{
            DNSServers = [String]::Join(', ', $dnsServersList)
            StaticRoutes = [String]::Join(', ', $staticRoutesList)
        }
        $htmlContent.Add(($dnsRoutingData | ConvertTo-Html -Fragment)) | Out-Null

        # --- Firewall Rules ---
        $htmlContent.Add('<h2 id="firewall">Firewall Rules</h2>') | Out-Null
        $firewallRules = Get-VMHostFirewallException -VMHost $vmHost -Server $viServer | Where-Object { $_.Enabled }
        $firewallData = $firewallRules | Select-Object Name, Enabled, Protocol, @{N='Port';E={$_.Port}}
        if ($firewallRules.Count -gt 0) { Write-Host "Enabled Firewall Rules: $(($firewallRules | ForEach-Object { $_.Name }) -join ', ')" -ForegroundColor White }
        $htmlContent.Add(($firewallData | ConvertTo-Html -Fragment)) | Out-Null

        # --- NTP Settings ---
        $htmlContent.Add('<h2 id="ntp">NTP Settings</h2>') | Out-Null
        $ntpConfig = Get-VMHostNtpServer -VMHost $vmHost -Server $viServer
        $ntpService = Get-VMHostService -VMHost $vmHost -Server $viServer | Where-Object { $_.Key -eq 'ntpd' }
        $ntpServersList = if ($null -ne $ntpConfig) { $ntpConfig } else { @() }
        Write-Host "NTP Servers: $([String]::Join(', ', $ntpServersList))" -ForegroundColor White
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
            if ($proxySwitch) { $proxySwitchMap[$dvs.ExtensionData.Uuid] = $proxySwitch.Pnic | ForEach-Object { $_.Split('-')[-1] } }
        }
        $networkHints = $netSystem.QueryNetworkHint($null)
        $hintTable = @{}
        foreach ($hint in $networkHints) { $hintTable[$hint.Device] = $hint }
        $nicSpeeds = @{}
        $hostView = Get-View -Id $vmHost.Id -Server $viServer
        foreach ($pnic in $hostView.Config.Network.Pnic) { $nicSpeeds[$pnic.Device] = $pnic.LinkSpeed.SpeedMb }

        $nicData = foreach ($nic in $physicalNics) {
            $hint = $hintTable[$nic.Name]
            $cdp = $hint.ConnectedSwitchPort
            $lldp = $hint.LLDPInfo
            $vSwitchList = @($standardSwitches | Where-Object { $_.Nic -contains $nic.Name } | ForEach-Object { $_.Name }) + @($distributedSwitches | Where-Object { $proxySwitchMap[$_.ExtensionData.Uuid] -and $proxySwitchMap[$_.ExtensionData.Uuid] -contains $nic.Name } | ForEach-Object { $_.Name })
            $vSwitchesString = [String]::Join(', ', $vSwitchList)
            Write-Host "$($nic.Name) Speed: $(if ($nicSpeeds[$nic.Name]) { "$($nicSpeeds[$nic.Name]) Mb/s" } else { 'Unknown' }), vSwitches: $vSwitchesString" -ForegroundColor White
            
            [PSCustomObject]@{
                Name = $nic.Name
                MAC = $nic.Mac
                LinkSpeed = if ($null -ne $nicSpeeds[$nic.Name]) { "$($nicSpeeds[$nic.Name]) Mb/s" } else { 'Unknown' }
                vSwitches = $vSwitchesString
                CDP_Switch = if ($cdp) { $cdp.DevId } else { 'N/A' }
                CDP_Port = if ($cdp) { $cdp.PortId } else { 'N/A' }
                CDP_Hardware = if ($cdp) { $cdp.HardwarePlatform } else { 'N/A' }
                LLDP_Switch = if ($lldp) { $lldp.SystemName } else { 'N/A' }
                LLDP_Port = if ($lldp) { $lldp.PortId } else { 'N/A' }
                LLDP_Hardware = if ($lldp) { $lldp.ChassisId } else { 'N/A' }
            }
        }
        $htmlContent.Add(($nicData | ConvertTo-Html -Fragment)) | Out-Null

        # Add footer
        $footerTimestamp = Get-Date -Format "MMMM dd, yyyy HH:mm:ss"
        $footer = "<div class='footer'>Generated by $scriptName (Version $scriptVersion) on $footerTimestamp</div>"
        $htmlContent.Add($footer) | Out-Null
        $htmlContent.Add('</body></html>') | Out-Null

        $fileName = Join-Path -Path $outputPath -ChildPath "NetworkConfig_${vCenterServer}_$($vmHost.Name)_$timestamp.html"
        $htmlContent | Out-File -FilePath $fileName -Encoding UTF8
        Write-Host "Generated report: $fileName" -ForegroundColor Green
    } catch {
        Write-Host "Error processing $($vmHost.Name): $_" -ForegroundColor Red
    }
}

Disconnect-VIServer -Server $vCenterServer -Force -Confirm:$false
Write-Host "Disconnected from $vCenterServer" -ForegroundColor Green
