<#
.SYNOPSIS
    Retreives basic ESXi host information
.DESCRIPTION
    This script will gather the following information for a vSphere Cluster, DataCenter or individual ESXi host:
    Hostname, Management IP, RAC IP, ESXi Version, ESXi Build, Make, Model, Serial Number, CPU Model, Speed, 
    Memory, Memory Slot Count, NIC Count.
.NOTES
    File Name     : GetESXiInventory.ps1
    Author        : Edgar Sanchez - @edmsanchez13
    Version       : 1.0
    Contributor: Ariel Sanchez - @arielsanchezmor
    Get-VMHostWSManInstance Function by Carter Shanklin - @cshanklin
    Downloaded from: http://poshcode.org/?show=928
.Link
  https://github.com/edmsanchez/vDocumentation
.INPUTS
   No inputs required
.OUTPUTS
   CSV file
.PARAMETER esxi
   The name(s) of the vSphere ESXi Host(s)
.EXAMPLE
    GetESXiInventory.ps1 -esxi devvm001.lab.local
.PARAMETER cluster
   The name(s) of the vSphere Cluster(s)
.EXAMPLE
    GetESXiInventory.ps1 -cluster production-cluster
.PARAMETER datacenter
   The name(s) of the vSphere Virtual DataCenter(s)
.EXAMPLE
    GetESXiInventory.ps1 -datacenter vDC001
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------

param(
    $esxi,
    $cluster,
    $datacenter
)

$outputCollection = @()
$outputFile = "Inventory.csv"

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
    $vmhost = Get-VMHost $esxihost
    Write-Host "`tGathering information from $vmhost ..."
    $hardware = $vmhost | Get-VMHostHardware | Select Manufacturer, Model, CpuModel, CpuCount, CpuCoreCountTotal, MhzPerCpu, MemorySlotCount, NicCount  
    $vmInfo = $vmhost | Select MemoryTotalGB, Version, Build

    # Get inventory info
    $mgmtIP = $vmhost | Get-VMHostNetworkAdapter | Where-Object {$_.ManagementTrafficEnabled -eq 'True'} | Select -ExpandProperty IP
    $hardwarePlatfrom = $esxcli.hardware.platform.get()
    $vmRam = $vmInfo.MemoryTotalGB -as [int]
    $racIP = Get-VMHostWSManInstance -VMHost $vmhost -class OMC_IPMIIPProtocolEndpoint -ignoreCertFailures | Select -ExpandProperty IPv4Address
    
    # Make a combined object
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
    # Add the object to the collection
    $outputCollection += $inventoryResults
}

# Display output on screen
Write-Host -ForegroundColor Green "`n" "ESXi Inventory:"
$outputCollection | Format-List 

# Export combined object
Write-Host -ForegroundColor Green "`tData was saved to" $outputFile "CSV file"
$outputCollection | Export-Csv $outputFile -NoTypeInformation
