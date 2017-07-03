<#
.SYNOPSIS
    Retreives basic ESXi host information
.DESCRIPTION
    This script will gather the following information for a vSphere Cluster, DataCenter or individual ESXi host:
    Hostname, Management IP, RAC IP, ESXi Version information, Hardware information.
.NOTES
    File Name     : GetESXiInventory.ps1
    Author        : Edgar Sanchez - @edmsanchez13
    Version       : 1.0
    Contributor: Ariel Sanchez - @arielsanchezmor
    Get-VMHostWSManInstance Function by Carter Shanklin - @cshanklin
    Downloaded from: http://poshcode.org/?show=928
.Link
  https://github.com/edmsanchez/vDocumentation
  https://virtualcornerstone.com/GetESXiInventory
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
.PARAMETER ExportCSV
    Exports all data to CSV file. File is saved on current userpath from where the script was executed
.EXAMPLE
    GetESXiInventory.ps1 -cluster production-cluster -ExportCSV
.PARAMETER ExportExcel
    Exports all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
    ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
    the  File is saved on current userpath from where the script was executed.
.EXAMPLE
    GetESXiInventory.ps1 -cluster production-cluster -ExportExcel
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------

param (
    $esxi,
    $cluster,
    $datacenter,
    [switch]$ExportCSV,
    [switch]$ExportExcel
)

$outputCollection = @()
$skipCollection = @()
$date = Get-Date -format s
$date = $date -replace ":","-"
$outputFile = "Inventory" + $date

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
if ($Global:DefaultViServers.Count -gt 0) {
    Clear-Host
    Write-Host "`tConnected to " $Global:DefaultViServers -ForegroundColor Green
} else {
    Write-Host "`tError: You must be connected to a vCenter or a vSphere Host before running this script." -ForegroundColor Red
    break
}

# Check to make sure at least 1 parameter was used
if ([string]::IsNullOrWhiteSpace($esxi) -and [string]::IsNullOrWhiteSpace($cluster) -and [string]::IsNullOrWhiteSpace($datacenter)) {
    Write-Host "`tError: You must at least use one parameter, run Get-Help " $MyInvocation.MyCommand.Name " for more information" -ForegroundColor Red
    break
}

# Gather host list
if ([string]::IsNullOrWhiteSpace($esxi)) {
    # $Vmhost Parameter Empty

    if ([string]::IsNullOrWhiteSpace($cluster)) {
        # $Cluster Parameter Empty

        if ([string]::IsNullOrWhiteSpace($datacenter)) {
            # $Datacenter Parameter Empty

        } else {                
            # Processing by Datacenter
            if ($datacenter -eq "all vdc") {
                Write-Host "`tGathering all hosts from the following vCenter(s): " $Global:DefaultViServers
                $vHostList = Get-VMHost | Sort-Object -Property Name
                
            } else {
                Write-Host "`tGathering host list from the following DataCenter(s): " (@($datacenter) -join ',')
                foreach ($vDCname in $datacenter) {
                    $tempList = Get-Datacenter -Name $vDCname.Trim() | Get-VMHost 
                    $vHostList += $tempList | Sort-Object -Property Name
                }
            }
        }
    } else {
        # Processing by Cluster
        Write-Host "`tGatehring host list from the following Cluster(s): " (@($cluster) -join ',')
        foreach ($vClusterName in $cluster) {
            $tempList = Get-Cluster -Name $vClusterName.Trim() | Get-VMHost 
            $vHostList += $tempList | Sort-Object -Property Name
        }
    }
} else {
    # Processing by ESXi Host
    Write-Host "`tGathering host list..."
    foreach ($invidualHost in $esxi) {
        $tempList = $invidualHost.Trim()
        $vHostList += $tempList | Sort-Object -Property Name
    }
}

# Main code execution
Foreach ($esxihost in $vHostList) {
    $esxcli = Get-EsxCli -VMHost $esxihost
    $vmhost = Get-VMHost -Name $esxihost
    # Skip if ESXi host is not in Connected or Maintenance ConnectionState
    if ($vmhost.ConnectionState -eq "Connected" -or $vmhost.ConnectionState -eq "Maintenance") {
        # Do nothing - ESXi host is reachable
    } else {
        # Make a combined object
        $skiphosts = New-Object -TypeName PSObject -Property ([ordered]@{
            'Hostname' = $esxihost
            'Connection State' = $esxihost.ConnectionState
            })
        #Add the object to the collection
        $skipCollection += $skiphosts
        continue
    }

    # Get inventory info
    Write-Host "`tGathering information from $vmhost ..."
    $hardware = $vmhost | Get-VMHostHardware -SkipAllSslCertificateChecks -WaitForAllData
    $vmInfo = $vmhost | Select-Object -Property MemoryTotalGB, Build
    $mgmtIP = $vmhost | Get-VMHostNetworkAdapter | Where-Object {$_.ManagementTrafficEnabled -eq 'True'} | Select-Object -ExpandProperty IP
    $hardwarePlatfrom = $esxcli.hardware.platform.get()
    $vmRam = $vmInfo.MemoryTotalGB -as [int]

    # Get RAC IP
    # Try with -class OMC_IPMIIPProtocolEndpoint First
    $rac = Get-VMHostWSManInstance -VMHost $vmhost -class OMC_IPMIIPProtocolEndpoint -ignoreCertFailures -ErrorAction SilentlyContinue
    if ($rac.Name) {
        $racIP = $rac.IPv4Address
    } else { # Else try with -class CIM_IPProtocolEndpoint
        $rac = Get-VMHostWSManInstance -VMHost $vmhost -class CIM_IPProtocolEndpoint -ignoreCertFailures -ErrorAction SilentlyContinue
        if ($rac.Name) {
            $racIP = $rac | Where-Object {$_.Name -match "Management Controller IP"} | Select-Object -ExpandProperty IPv4Address
        }
    }

    #Get ESXi version details
    $vmhostView = $vmhost | Get-View
    $esxiVersion = $esxcli.system.version.get()
    
    # Make a combined object
    $inventoryResults = New-Object -TypeName PSObject -Property ([ordered]@{
        'Hostname' = $vmhost
        'Management IP' = $mgmtIP
        'RAC IP' = $racIP
        'Product' = $vmhostView.Config.Product.Name
        'Version' = $vmhostView.Config.Product.Version
        'Build' = $Vminfo.Build
        'Update' = $esxiVersion.Update
        'Patch' = $esxiVersion.Patch
        'Make'= $hardware.Manufacturer
        'Model' = $hardware.Model
        'S/N' = $hardwarePlatfrom.serialNumber
        'BIOS Version' = $hardware.BiosVersion
        'CPU Model' = $hardware.CpuModel
        'CPU Count' = $hardware.CpuCount
        'CPU Core Total' = $hardware.CpuCoreCountTotal
        'Speed (MHz)' = $hardware.MhzPerCpu
        'Memory (GB)' = $vmRam
        'Memory Slots Count' = $hardware.MemorySlotCount
        'Memory Slots Used' = $hardware.MemoryModules.Count
        'Power Supplies' = $hardware.PowerSupplies.Count
        'NIC Count' = $hardware.NicCount
    })
    # Add the object to the collection
    $outputCollection += $inventoryResults
}

# Display output on screen
Write-Host "`n" "ESXi Inventory:" -ForegroundColor Green
$outputCollection | Format-List 

# Display list of Hosts that were skipped
If ($skipCollection.count -gt 0) {
    Write-Host "`tSkipped hosts: " -ForegroundColor Yellow
    $skipCollection | Format-Table -AutoSize
}

# Export combined object
# Export to CSV
if ($ExportCSV) {
    Write-Host "`tData exported to" ($outputFile + ".csv") "file" -ForegroundColor Green
    $outputCollection | Export-Csv ($outputFile + ".csv") -NoTypeInformation
}

# Export to Excel
if ($ExportExcel) {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        $outputCollection | Export-Excel ($outputFile + ".xlsx") -BoldTopRow -WorkSheetname Inventory
    } else {
        Write-Host "ImportExcel Module missing, see help for more information" -ForegroundColor Red
    }
}
