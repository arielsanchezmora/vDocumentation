function Get-ESXInventory {
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
    Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
.EXAMPLE
    GetESXiInventory.ps1 -cluster production-cluster -ExportCSV
.PARAMETER ExportExcel
    Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
    ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
    File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
.EXAMPLE
    GetESXiInventory.ps1 -cluster production-cluster -ExportExcel
.PARAMETER folderPath
    Specificies an alternate folder path of where the exported file should be saved.
.EXAMPLE
    GetESXiInventory.ps1 -cluster production-cluster -ExportExcel -folderPath C:\temp
#>

<#
----------------------------------------------------------[Declarations]----------------------------------------------------------
#>

    [CmdletBinding()]
    param (
        $esxi,
        $cluster,
        $datacenter,
        [switch]$ExportCSV,
        [switch]$ExportExcel,
        $folderPath
    )

    $outputCollection = @()
    $skipCollection = @()
    $date = Get-Date -format s
    $date = $date -replace ":","-"
    $outputFile = "Inventory" + $date

<#
----------------------------------------------------------[Functions]----------------------------------------------------------
#>    

    function Get-VMHostWSManInstance {
        param (
	        [Parameter(Mandatory=$TRUE,HelpMessage="VMHosts to probe")]
	        [VMware.VimAutomation.Client20.VMHostImpl[]]$vmhost,

	        [Parameter(Mandatory=$TRUE,HelpMessage="Class Name")]
	        [string]$class,
            [switch]$ignoreCertFailures,
            [System.Management.Automation.PSCredential]$credential=$null
        )
        Write-Verbose ((get-date -Format G) + "`tGet-VMHostWSManInstance    Started execution")
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
    Write-Verbose ((get-date -Format G) + "`tGet-VMHostWSManInstance    Finished execution")
    }

<#
----------------------------------------------------------[Execution]----------------------------------------------------------
#>

    <#
        Check to see if there are any currently connected servers
    #>
    Write-Verbose ((get-date -Format G) + "`tValidate connection to to a vSphere server")
    if ($Global:DefaultViServers.Count -gt 0) {
        Clear-Host
        Write-Host "`tConnected to $Global:DefaultViServers" -ForegroundColor Green
    } else {
        Write-Error "You must be connected to a vCenter or vSphere Host before running this cmdlet."
        break
    }

    <#
        Check to make sure at least one parameter was used (esxi, cluster, or datacenter
        Although all 3 can be specified, only the first one is taken
        Example -esxi "host001" -cluster "test-cluster" : esxi parameter will be used
    #>
    Write-Verbose ((get-date -Format G) + "`tValidate parameters used")
    if ([string]::IsNullOrWhiteSpace($esxi) -and [string]::IsNullOrWhiteSpace($cluster) -and [string]::IsNullOrWhiteSpace($datacenter)) {
        Write-Error "You must use a parameter (-esxi, -cluster, -datacenter). Use Get-Help for more information"
        break
    }

    <#
        Gather host list
        will check against esxi, cluster, and datacenter parameter
        processed as first come
    #>
    Write-Verbose ((get-date -Format G) + "`tGather host list")
    if ([string]::IsNullOrWhiteSpace($esxi)) {
        Write-Verbose ((get-date -Format G) + "`t-esxi parameter is Null or Empty")

        if ([string]::IsNullOrWhiteSpace($cluster)) {
            Write-Verbose ((get-date -Format G) + "`t-cluster parameter is Null or Empty")

            if ([string]::IsNullOrWhiteSpace($datacenter)) {
                Write-Verbose ((get-date -Format G) + "`t-datacenter parameter is Null or Empty")

            } else {
                Write-Verbose ((get-date -Format G) + "`tExecuting cmdlet using datacenter parameter")
                if ($datacenter -eq "all vdc") {
                    Write-Host "`tGathering all hosts from the following vCenter(s): " $Global:DefaultViServers
                    $vHostList = Get-VMHost | Sort-Object -Property Name
                
                } else {
                    Write-Host "`tGathering host list from the following DataCenter(s): " (@($datacenter) -join ',')
                    foreach ($vDCname in $datacenter) {
                        $tempList = Get-Datacenter -Name $vDCname.Trim() -ErrorAction SilentlyContinue | Get-VMHost 
                        if ([string]::IsNullOrWhiteSpace($tempList)) {
                            Write-Warning "`tDatacenter with name $vDCname was not found in $Global:DefaultViServers"
                        } else {
                            $vHostList += $tempList | Sort-Object -Property Name
                        }
                    }
                }
            }
        } else {
            Write-Verbose ((get-date -Format G) + "`tExecuting cmdlet using cluster parameter")
            Write-Host "`tGathering host list from the following Cluster(s): " (@($cluster) -join ',')
            foreach ($vClusterName in $cluster) {
                $tempList = Get-Cluster -Name $vClusterName.Trim() -ErrorAction SilentlyContinue | Get-VMHost 
                if ([string]::IsNullOrWhiteSpace($tempList)) {
                    Write-Warning "`tCluster with name $vClusterName was not found in $Global:DefaultViServers"
                } else {
                    $vHostList += $tempList | Sort-Object -Property Name
                }
            }
        }
    } else {
        Write-Verbose ((get-date -Format G) + "`tExecuting cmdlet using esxi parameter")
        Write-Host "`tGathering host list..."
        foreach ($invidualHost in $esxi) {
            $tempList = $invidualHost.Trim()
            $vHostList += $tempList | Sort-Object -Property Name
        }
    }

    <#
        Validate export switches,
        folder path and dependencies
    #>
    Write-Verbose ((get-date -Format G) + "`tValidate export switches and folder path")
    if ($ExportCSV -or $ExportExcel) {
        $currentLocation = (Get-Location).Path
        if ([string]::IsNullOrWhiteSpace($folderPath)) {
            Write-Verbose ((get-date -Format G) + "`t-folderPath parameter is Null or Empty")
            Write-Warning "`tFolder Path (-folderPath) was not specified for saving exported data. The current location: '$currentLocation' will be used"
            $outputFile = $currentLocation + "\" + $outputFile
        } else {
            if (Test-Path $folderPath) {
                Write-Verbose ((get-date -Format G) + "`t'$folderPath' path found")
                $lastCharsOfFolderPath = $folderPath.Substring($folderPath.Length - 1)
                if ($lastCharsOfFolderPath -eq "\" -or $lastCharsOfFolderPath -eq "/") {
                    $outputFile = $folderPath + $outputFile
                } else {
                    $outputFile = $folderPath + "\" + $outputFile
                }
                Write-Verbose ((get-date -Format G) + "`t$outputFile")
            } else {
                Write-Warning "`t'$folderPath' path not found. The current location: '$currentLocation' will be used instead"
                $outputFile = $currentLocation + "\" + $outputFile
            }
        }
    }

    if ($ExportExcel) {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Write-Verbose ((get-date -Format G) + "`tImportExcel Module available")
        } else {
            Write-Warning "`tImportExcel Module missing. Will export data to CSV file instead"
            Write-Warning "`tImportExcel Module can be installed directly from the PowerShell Gallery"
            Write-Warning "`tSee https://github.com/dfinke/ImportExcel for more information"
            $ExportExcel = $false
            $ExportCSV = $true
        }
    }

    <#
        Main code execution
    #>
    Foreach ($esxihost in $vHostList) {
        $esxcli = Get-EsxCli -VMHost $esxihost
        $vmhost = Get-VMHost -Name $esxihost
        
        <#
            Skip if ESXi host is not in a Connected
            or Maintenance ConnectionState
        #>
        Write-Verbose ((get-date -Format G) + "`t$vmhost Connection State: " + $vmhost.ConnectionState)
        if ($vmhost.ConnectionState -eq "Connected" -or $vmhost.ConnectionState -eq "Maintenance") {
            <#
                Do nothing - ESXi host is reachable
            #>
        } else {
            <#
                Use a custom object and array
                to keep track of all skipped
                hosts and continue to the next
                back to the Foreach loop
            #>
            $skiphosts = New-Object -TypeName PSObject -Property ([ordered]@{
                'Hostname' = $esxihost
                'Connection State' = $esxihost.ConnectionState
                })
            
            $skipCollection += $skiphosts
            continue
        }

        <#
            Get inventory info
        #>
        Write-Host "`tGathering information from $vmhost ..."
        $hardware = $vmhost | Get-VMHostHardware -SkipAllSslCertificateChecks -WaitForAllData -ErrorAction SilentlyContinue
        $vmInfo = $vmhost | Select-Object -Property MemoryTotalGB, Build
        $mgmtIP = $vmhost | Get-VMHostNetworkAdapter | Where-Object {$_.ManagementTrafficEnabled -eq 'True'} | Select-Object -ExpandProperty IP
        $hardwarePlatfrom = $esxcli.hardware.platform.get()
        $vmRam = $vmInfo.MemoryTotalGB -as [int]

        <#
            Get RAC IP
            Try with -class OMC_IPMIIPProtocolEndpoint First
            Else try with -class CIM_IPProtocolEndpoint
        #>
        Write-Verbose ((get-date -Format G) + "`tGet-VMHostWSManInstance using OMC_IPMIIPProtocolEndpoint class")
        $rac = Get-VMHostWSManInstance -VMHost $vmhost -class OMC_IPMIIPProtocolEndpoint -ignoreCertFailures -ErrorAction SilentlyContinue
        if ($rac.Name) {
            $racIP = $rac.IPv4Address
        } else { 
            Write-Verbose ((get-date -Format G) + "`tGet-VMHostWSManInstance using CIM_IPProtocolEndpoint class")
            $rac = Get-VMHostWSManInstance -VMHost $vmhost -class CIM_IPProtocolEndpoint -ignoreCertFailures -ErrorAction SilentlyContinue
            if ($rac.Name) {
                $racIP = $rac | Where-Object {$_.Name -match "Management Controller IP"} | Select-Object -ExpandProperty IPv4Address
                Write-Verbose ((get-date -Format G) + "`tRAC IP gathered using CIM_IPProtocolEndpoint class")
            }
        }

        <#
            Get ESXi version details
        #>
        $vmhostView = $vmhost | Get-View
        $esxiVersion = $esxcli.system.version.get()
    
        <#
            Use a custom object and array
            to store inventory information
        #>
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
        
        $outputCollection += $inventoryResults
    }

    <#
        Display skipped hosts and their connection status
    #>
    If ($skipCollection.count -gt 0) {
        Write-Warning "`tSkipped hosts: "
        $skipCollection | Format-Table -AutoSize
    }

    <#
        Output to screen
        Export data to CSV, Excel
    #>
    if ($outputCollection.count -gt 0) {
        Write-Host "`n" "ESXi Inventory:" -ForegroundColor Green
        $outputCollection | Format-List 

        if ($ExportCSV) {
            $outputCollection | Export-Csv ($outputFile + ".csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + ".csv") "file" -ForegroundColor Green
        }

        if ($ExportExcel) {
            $outputCollection | Export-Excel ($outputFile + ".xlsx") -BoldTopRow -WorkSheetname Inventory
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
    } else {
        Write-Verbose ((get-date -Format G) + "`tNo information gathered")
    }

}
