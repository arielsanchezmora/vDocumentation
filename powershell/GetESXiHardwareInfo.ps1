<#
.SYNOPSIS
    Retreives ESXi Hardware PCI/IO device information.
.DESCRIPTION
    This script will gather PCI/IO device information including HCL IDs for the following VMkernel name(s):
    Network Controller - vmnic*
    Storage Controller - vmhba*
    Graphic Device - vmgfx*
.NOTES
    File Name     : GetESXiHardwareInfo.ps1
    Author        : Edgar Sanchez - @edmsanchez13
    Version       : 1.0
    Contributor: Ariel Sanchez - @arielsanchezmor
.Link
  https://github.com/edmsanchez/vDocumentation
.INPUTS
   No inputs required
.OUTPUTS
   CSV file
.PARAMETER esxi
   The name(s) of the vSphere ESXi Host(s)
.EXAMPLE
    GetESXiHardwareInfo.ps1 -esxi devvm001.lab.local
.PARAMETER cluster
   The name(s) of the vSphere Cluster(s)
.EXAMPLE
    GetESXiHardwareInfo.ps1 -cluster production-cluster
.PARAMETER datacenter
   The name(s) of the vSphere Virtual DataCenter(s)
.EXAMPLE
    GetESXiHardwareInfo.ps1 -datacenter vDC001
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------

param(
    $esxi,
    $cluster,
    $datacenter
)

$outputCollection = @()
$outputFile = "HardwareIO.csv"

#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Check to see if there are any currently connected servers
if($Global:DefaultViServers.Count -gt 0) {
    Clear-Host
    Write-Host -ForegroundColor Green "`tConnected to " $Global:DefaultViServers
} else {
    Write-Host -ForegroundColor Red "`tError: You must be connected to a vCenter or a vSphere Host before running this script."
    break
}

# Check to make sure at least 1 parameter was used
if([string]::IsNullOrWhiteSpace($esxi) -and [string]::IsNullOrWhiteSpace($cluster) -and [string]::IsNullOrWhiteSpace($datacenter)) {
    Write-Host -ForegroundColor Red "`tError: You must at least use one parameter, run Get-Help " $MyInvocation.MyCommand.Name " for more information"
    break
}

# Gather host list
if([string]::IsNullOrWhiteSpace($esxi)) {
    # $Vmhost Parameter Empty

    if([string]::IsNullOrWhiteSpace($cluster)) {
        # $Cluster Parameter Empty

        if([string]::IsNullOrWhiteSpace($datacenter)) {
            # $Datacenter Parameter Empty

        } else {                
            # Processing by Datacenter
            Write-Host "`tGathering host list from the following DataCenter(s): " (@($datacenter) -join ',')
            foreach ($vDCname in $datacenter) {
                $tempList = Get-DataCenter $vDCname.Trim() | Get-VMHost 
                $vHostList += $tempList | Sort-Object -Property name
            }
        }
    } else {
        # Processing by Cluster
        Write-Host "`tGatehring host list from the following Cluster(s): " (@($cluster) -join ',')
        foreach ($vClusterName in $cluster) {
            $tempList = Get-Cluster $vClusterName.Trim() | Get-VMHost 
            $vHostList += $tempList | Sort-Object -Property name
        }
    }
} else {
    # Processing by ESXi Host
    Write-Host "`tGathering host list..."
    foreach($invidualHost in $esxi) {
        $tempList = $invidualHost.Trim()
        $vHostList += $tempList | Sort-Object -Property name
    }
}

# Main code execution
Foreach($esxihost in $vHostList) {
    $esxcli = Get-EsxCli -VMHost $esxihost
    $esxcli2 = Get-EsxCli -V2 -VMHost $esxihost
    $vmhost = Get-VMHost $esxihost
    Write-Host "`tGathering information from $vmhost ..."
    $pciDevices = $esxcli.hardware.pci.list() | where {$_.VMKernelName -like "vmhba*" -OR $_.VMKernelName -like "vmnic*" -OR $_.VMKernelName -like "vmgfx*" }
     
    Foreach ($pciDevice in $pciDevices) {
        $device = $vmhost | Get-VMHostPciDevice | Where { $pciDevice.Address -match $_.Id }
            
        # Get driver version
        $driverVersion = $esxcli.system.module.get($pciDevice.ModuleName) | Select -ExpandProperty Version

        # Get NIC firmware version
        if ($pciDevice.VMKernelName -like 'vmnic*') {
            $vmnicDetail = $esxcli2.network.nic.get.Invoke(@{nicname = $pciDevice.VMKernelName})
            $firmwareVersion = $vmnicDetail.DriverInfo.FirmwareVersion
            
            #Get NIC driver VIB package version
            $driverVib = $esxcli.software.vib.list() | Select Name, Version | Where {$_.Name -eq "net-"+$vmnicDetail.DriverInfo.Driver}
            $vibName = $driverVib.Name
            $vibVersion = $driverVib.Version
        } elseif ($pciDevice.VMKernelName -like 'vmhba*') {
            # Can't get HBA Firmware from Powercli ATM only through SSH or using Putty Plink+PowerCli
            $firmwareVersion = ""

            # Get HBA driver VIB package version
            $vibName = $pciDevice.ModuleName -replace "_", "-"
            $driverVib = $esxcli.software.vib.list() | Select Name,Version | Where {$_.Name -eq "scsi-"+$VibName -or $_.Name -eq "sata-"+$VibName -or $_.Name -eq $VibName}
            $vibName = $driverVib.Name
            $vibVersion = $driverVib.Version
        } else {
            $firmwareVersion = ""
            $vibName = ""
            $vibVersion = ""
        }
       
        # Make a combined object
        $hardwwareResults = New-Object -Type PSObject -Prop ([ordered]@{
            'Hostname' = $vmhost
            'Slot Description' = $pciDevice.SlotDescription
            'VMKernel Name' = $pciDevice.VMKernelName
            'Device Name' = $pciDevice.DeviceName
            'Vendor Name' = $pciDevice.VendorName
            'Device Class' = $pciDevice.DeviceClassName
            'PCI Address' = $pciDevice.Address
            'VID' = [String]::Format("{0:x4}", $device.VendorId)
            'DID' = [String]::Format("{0:x4}", $device.DeviceId)
            'SVID' = [String]::Format("{0:x4}", $device.SubVendorId)
            'SSID' = [String]::Format("{0:x4}", $device.SubDeviceId)
            'Driver' = $pciDevice.ModuleName
            'Driver Version' = $driverVersion
            'Firmware Version' = $firmwareVersion
            'VIB Name' = $vibName
            'VIB Version' = $vibVersion
        })
        #Add the object to the collection
        $outputCollection += $hardwwareResults
    }
}

# Display output on screen
Write-Host -ForegroundColor Green "`n" "Hardwaer PCI/IO Details:"
$outputCollection | Format-List 

# Export combined object
Write-Host -ForegroundColor Green "`tData was saved to" $outputFile "CSV file"
$outputCollection | Export-Csv $outputFile -NoTypeInformation
