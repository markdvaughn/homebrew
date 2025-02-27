#Requires -Modules VMware.PowerCLI

# Pre-defined vCenter server list
$vCenterList = @(
    @{ Server = 'vc01.example.com'; Description = 'Production Cluster - East Coast' }
    @{ Server = 'vc02.example.com'; Description = 'Production Cluster - West Coast' }
    @{ Server = 'vc03.example.com'; Description = 'Test/Dev Environment' }
)

# Default output path
$defaultOutputPath = 'C:\Reports\ESXiHardwareOverview'
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# CSS styling
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

# Function to get valid numeric choice
function Get-NumericChoice {
    param (
        [string]$Prompt,
        [int]$Min,
        [int]$Max
    )
    do {
        Write-Host $Prompt
        $input = Read-Host "Enter your choice"
        if ($input -match '^\d+$' -and [int]$input -ge $Min -and [int]$input -le $Max) {
            return [int]$input
        }
        Write-Host "Error: Please enter a number between $Min and $Max" -ForegroundColor Red
    } while ($true)
}

# PowerCLI Configuration
$PSDefaultParameterValues['Out-Default:Width'] = 200
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null

# Disconnect existing connections
Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue

# vCenter selection
Write-Host "`nAvailable vCenter Servers:" -ForegroundColor Cyan
for ($i = 0; $i -lt $vCenterList.Count; $i++) {
    Write-Host "$($i + 1). $($vCenterList[$i].Server) - $($vCenterList[$i].Description)"
}
Write-Host "$($vCenterList.Count + 1). Enter a custom vCenter server manually"

$vCenterChoice = Get-NumericChoice -Prompt "`nSelect a vCenter server (1-$($vCenterList.Count + 1)):" -Min 1 -Max ($vCenterList.Count + 1)

if ($vCenterChoice -le $vCenterList.Count) {
    $vCenterServer = $vCenterList[$vCenterChoice - 1].Server
    $description = $vCenterList[$vCenterChoice - 1].Description
    Write-Host "Selected: $vCenterServer - $description" -ForegroundColor Green
} else {
    $vCenterServer = Read-Host "Enter custom vCenter server address"
    Write-Host "Selected custom server: $vCenterServer" -ForegroundColor Green
}

# Output path selection
Write-Host "`nOutput Path Options:" -ForegroundColor Cyan
Write-Host "1. Default path: $defaultOutputPath"
Write-Host "2. Script execution directory: $scriptPath"
Write-Host "3. Custom path"

$pathChoice = Get-NumericChoice -Prompt "`nSelect output path (1-3):" -Min 1 -Max 3

switch ($pathChoice) {
    1 { $outputPath = $defaultOutputPath }
    2 { $outputPath = $scriptPath }
    3 { $outputPath = Read-Host "Enter custom output path" }
}

Write-Host "Selected output path: $outputPath" -ForegroundColor Green

# Ensure output path exists
try {
    if (-not (Test-Path $outputPath)) {
        New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
        Write-Host "Created directory: $outputPath" -ForegroundColor Yellow
    }
} catch {
    Write-Error "Failed to create output directory: $_"
    exit
}

# Main execution
try {
    # Connect to vCenter
    $credential = Get-Credential -Message "Enter vCenter credentials for $vCenterServer"
    Write-Host "Connecting to $vCenterServer..." -ForegroundColor Yellow
    Connect-VIServer -Server $vCenterServer -Credential $credential -ErrorAction Stop

    # Get current timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Gather ESXi host information
    Write-Host "Collecting ESXi host information..." -ForegroundColor Yellow
    $esxiHosts = Get-VMHost | Sort-Object Name | ForEach-Object {
        $hardware = Get-VMHostHardware -VMHost $_
        $license = Get-VMHost $_ | Get-View | Select-Object -ExpandProperty LicenseKey
        
        [PSCustomObject]@{
            HostName = $_.Name
            Version = $_.Version
            Build = $_.Build
            ManagementIP = ($_.ExtensionData.Config.Network.Vnic | Where-Object {$_.Portgroup -eq "Management Network"}).Spec.Ip.IpAddress
            LicenseStatus = if ($license -match "evaluation") {"Evaluation"} else {"Licensed"}
            Manufacturer = $hardware.Manufacturer
            Motherboard = $hardware.Model
            SerialNumber = $hardware.SerialNumber
            ChassisSerial = $hardware.ChassisSerialNumber
            AssetTag = $hardware.AssetTag
            CPUModel = $hardware.CpuModel
            CPUSpeed = "$($hardware.CpuMhz) MHz"
            Sockets = $hardware.CpuPkg.Count
            Cores = $hardware.CpuCoreCountTotal
            MemoryGB = [math]::Round($hardware.MemorySize/1GB, 2)
            BIOSVersion = $hardware.BiosVersion
        }
    }

    # Create HTML content
    $htmlBody = @"
    $css
    <h1>ESXi Host Inventory Report - $vCenterServer</h1>
    <div class="timestamp">Generated on: $timestamp</div>
    <h2>Host Details</h2>
    <table>
        <tr>
            <th>Host Name</th>
            <th>Version</th>
            <th>Build</th>
            <th>Management IP</th>
            <th>License Status</th>
            <th>Manufacturer</th>
            <th>Motherboard</th>
            <th>Serial Number</th>
            <th>Chassis Serial</th>
            <th>Asset Tag</th>
            <th>CPU Model</th>
            <th>CPU Speed</th>
            <th>Sockets</th>
            <th>Cores</th>
            <th>Memory (GB)</th>
            <th>BIOS Version</th>
        </tr>
"@

    foreach ($host in $esxiHosts) {
        $htmlBody += @"
        <tr>
            <td>$($host.HostName)</td>
            <td>$($host.Version)</td>
            <td>$($host.Build)</td>
            <td>$($host.ManagementIP)</td>
            <td>$($host.LicenseStatus)</td>
            <td>$($host.Manufacturer)</td>
            <td>$($host.Motherboard)</td>
            <td>$($host.SerialNumber)</td>
            <td>$($host.ChassisSerial)</td>
            <td>$($host.AssetTag)</td>
            <td>$($host.CPUModel)</td>
            <td>$($host.CPUSpeed)</td>
            <td>$($host.Sockets)</td>
            <td>$($host.Cores)</td>
            <td>$($host.MemoryGB)</td>
            <td>$($host.BIOSVersion)</td>
        </tr>
"@
    }

    $htmlBody += "</table>"

    # Complete HTML
    $html = "<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body>$htmlBody</body></html>"

    # Save to file
    $outputFile = Join-Path $outputPath "ESXi_Host_Report_$vCenterServer_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $html | Out-File $outputFile
    Write-Host "Report generated successfully: $outputFile" -ForegroundColor Green

} catch {
    Write-Error "An error occurred: $_"
} finally {
    if ($global:DefaultVIServer) {
        Disconnect-VIServer -Server $vCenterServer -Confirm:$false
        Write-Host "Disconnected from $vCenterServer" -ForegroundColor Yellow
    }
}
