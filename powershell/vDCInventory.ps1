<#
.Synopsis
 Script to gather basic ESXi host information.
    
.Description
 Will Gather the following information for Virtual Datacenter and export to CSV file: Hostname, Management IP,
 RAC IP, ESXi Version, ESXi Build, Make, Model, Serial Number, CPU Model, Speed, Memory, Memory Slot Count, NIC Count.

 .Link
  https://github.com/edmsanchez/vDocumentation
 
 .Notes
  Script by: Edgar Sanchez
  Email: Ed.Sanchez@live.com
  Twitter: @edsanchez
  Contributor: Ariel Sanchez
  Twitter: @arielsanchezmor
  Get-VMHostWSManInstance Function by Carter Shanklin
  Twitter: @cshanklin.
  Downloaded from: http://poshcode.org/?show=928
  V1.0 - 04/10/2017     
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------
#Update $vDCName Variable to run in your environment

$outputCollection = @()
$vDCName = "Your virtual Datacenter name"
$vDCHosts = Get-DataCenter $vDCName | Get-VMHost | Sort-Object -Property Name

#-----------------------------------------------------------[Functions]------------------------------------------------------------


function Get-VMHostWSManInstance {
	param (
	[Parameter(Mandatory=$TRUE,HelpMessage="VMHosts to probe")]
	[VMware.VimAutomation.Client20.VMHostImpl[]]
	$vmhost,

	[Parameter(Mandatory=$TRUE,HelpMessage="Class Name")]
	[string]
	$class,

	[switch]
	$ignoreCertFailures,

	[System.Management.Automation.PSCredential]
	$credential=$null
	)

	$omcBase = "http://schema.omc-project.org/wbem/wscim/1/cim-schema/2/"
	$dmtfBase = "http://schemas.dmtf.org/wbem/wscim/1/cim-schema/2/"
	$vmwareBase = "http://schemas.vmware.com/wbem/wscim/1/cim-schema/2/"

	if ($ignoreCertFailures) {
		$option = New-WSManSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
	} else {
		$option = New-WSManSessionOption
	}
	foreach ($H in $vmhost) {
		if ($credential -eq $null) {
			$hView = $H | Get-View -property Value
			$ticket = $hView.AcquireCimServicesTicket()
			$password = convertto-securestring $ticket.SessionId -asplaintext -force
			$credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $ticket.SessionId, $password
		}
		$uri = "https`://" + $h.Name + "/wsman"
		if ($class -cmatch "^CIM") {
			$baseUrl = $dmtfBase
		} elseif ($class -cmatch "^OMC") {
			$baseUrl = $omcBase
		} elseif ($class -cmatch "^VMware") {
			$baseUrl = $vmwareBase
		} else {
			throw "Unrecognized class"
		}
		Get-WSManInstance -Authentication basic -ConnectionURI $uri -Credential $credential -Enumerate -Port 443 -UseSSL -SessionOption $option -ResourceURI "$baseUrl/$class"
	}
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Foreach($vmhost in $vDCHosts) {
    $esxcli = Get-EsxCli -VMHost $vmhost
    $hardware = Get-VMHost $vmhost | Get-VMHostHardware | Select Manufacturer, Model, CpuModel, CpuCount, CpuCoreCountTotal, MhzPerCpu, MemorySlotCount, NicCount  
    $vmInfo = Get-VMHost $vmhost | Select MemoryTotalGB, Version, Build

    #Get inventory info
    $mgmtIP = Get-VMHost $vmhost | Get-VMHostNetworkAdapter | Where-Object {$_.ManagementTrafficEnabled -eq 'True'} | Select -ExpandProperty IP
    $hardwarePlatfrom = $esxcli.hardware.platform.get()
    $vmRam = $vmInfo.MemoryTotalGB -as [int]
    $racIP = Get-VMHostWSManInstance -VMHost (Get-VMHost $vmhost) -class OMC_IPMIIPProtocolEndpoint -ignoreCertFailures | Select -ExpandProperty IPv4Address
    
    #Make a combined object
    $inventoryResults = New-Object -Type PSObject -Prop ([ordered]@{
        'Hostname' = $vmhost
        'Management IP' = $mgmtIP
        'RAC IP' = $racIP
        'OS/ESXi' = $vmInfo.Version
        'Build' = $vmInfo.Build
        'Make'= $hardware.Manufacturer
        'Model' = $hardware.Model
        'S/N' = $hardwarePlatfrom.serialNumber
        'CPU Model' = $hardware.CpuModel
        'CPU Count' = $hardware.CpuCount
        'CPU Core Total' = $hardware.CpuCoreCountTotal
        'Speed (MHz)' = $hardware.MhzPerCpu
        'Memory (GB)' = $vmRam
        'Memory Slot Count' = $hardware.MemorySlotCount
        'NIC Count' = $hardware.NicCount
    })
    #Add the object to the collection
    $outputCollection += $inventoryResults
}
$outputCollection | Export-Csv Inventory.csv -NoTypeInformation
