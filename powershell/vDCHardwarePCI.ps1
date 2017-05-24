<#
.Synopsis
 Script to get PCI/IO Devices information including HCL IDs. It gathers the following VMKernel Name:
 Network Controller - vmnic*
 Storage Controller - vmhba*
 Graphic Device - vmgfx*
    
.Description
 Will Gather the following information for a Virtual Datacenter and export to CSV file: Hostname,	
 Slot Description, VMKernel Name, Device Name, Vendor Name, Device Class, PCI Address,
 VID, DID, SVID, SSID, Driver, Driver Version, Firmware Version, VIB Name, .VIB Version
 
 .Link
  https://github.com/edmsanchez/vDocumentation
 
 .Notes
  Script by: Edgar Sanchez
  Email: Ed.Sanchez@live.com
  Twitter: @edsanchez
  Contributor: Ariel Sanchez
  Twitter: @arielsanchezmor
  V1.0 - 5/12/2017     
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------
#Update $vDCName Variable to run in your environment

$outputCollection = @()
$vDCName = "Your vDC Name Here"
$vDCHosts = Get-DataCenter $vDCName | Get-VMHost | Sort-Object -Property Name

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Foreach($vmhost in $vDCHosts) {
    $esxcli = Get-EsxCli -VMHost $vmhost
    $esxcli2 = Get-EsxCli -V2 -VMHost $vmhost
    $pciDevices = $esxcli.hardware.pci.list() | where {$_.VMKernelName -like "vmhba*" -OR $_.VMKernelName -like "vmnic*" -OR $_.VMKernelName -like "vmgfx*" }
     
    Foreach ($pciDevice in $pciDevices) {
        $device = Get-VMHost $vmhost | Get-VMHostPciDevice | Where { $pciDevice.Address -match $_.Id }
            
        #Get driver version
        $driverVersion = $esxcli.system.module.get($pciDevice.ModuleName) | Select -ExpandProperty Version

        #Get NIC firmware version
        if ($pciDevice.VMKernelName -like 'vmnic*') {
            $vmnicDetail = $esxcli2.network.nic.get.Invoke(@{nicname = $pciDevice.VMKernelName})
            $firmwareVersion = $vmnicDetail.DriverInfo.FirmwareVersion
            
            #Get NIC driver VIB package version
            $driverVib = $esxcli.software.vib.list() | Select Name, Version | Where {$_.Name -eq "net-"+$vmnicDetail.DriverInfo.Driver}
            $vibName = $driverVib.Name
            $vibVersion = $driverVib.Version
        } elseif ($pciDevice.VMKernelName -like 'vmhba*') {
            #Can't get HBA Firmware from Powercli ATM only through SSH or using Putty Plink+PowerCli
            $firmwareVersion = ""

            #Get HBA driver VIB package version
            $vibName = $pciDevice.ModuleName -replace "_", "-"
            $driverVib = $esxcli.software.vib.list() | Select Name,Version | Where {$_.Name -eq "scsi-"+$VibName -or $_.Name -eq "sata-"+$VibName -or $_.Name -eq $VibName}
            $vibName = $driverVib.Name
            $vibVersion = $driverVib.Version
        } else {
            $firmwareVersion = ""
            $vibName = ""
            $vibVersion = ""
        }
       
        #Make a combined object
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
$outputCollection | Export-Csv HardwarePCI.csv -NoTypeInformation
