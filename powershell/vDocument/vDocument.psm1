function Get-ESXInventory {
<#
.SYNOPSIS
    Get basic ESXi host information
.DESCRIPTION
    Will get inventory information for a vSphere Cluster, Datacenter or individual ESXi host
    The following is gathered:
    Hostname, Management IP, RAC IP, ESXi Version information, Hardware information
.NOTES
    Author     : Edgar Sanchez - @edmsanchez13
    Contributor: Ariel Sanchez - @arielsanchezmor
    Get-VMHostWSManInstance Function by Carter Shanklin - @cshanklin
    Downloaded from: http://poshcode.org/?show=928
.Link
  https://github.com/edmsanchez/vDocumentation
.INPUTS
   No inputs required
.OUTPUTS
   CSV file
   Excel file
.PARAMETER esxi
   The name(s) of the vSphere ESXi Host(s)
.EXAMPLE
    Get-ESXInventory -esxi devvm001.lab.local
.PARAMETER cluster
   The name(s) of the vSphere Cluster(s)
.EXAMPLE
    Get-ESXInventory -cluster production-cluster
.PARAMETER datacenter
   The name(s) of the vSphere Virtual Datacenter(s)
.EXAMPLE
    Get-ESXInventory -datacenter vDC001
.PARAMETER ExportCSV
    Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
.EXAMPLE
    Get-ESXInventory -cluster production-cluster -ExportCSV
.PARAMETER ExportExcel
    Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
    ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
    File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
.EXAMPLE
    Get-ESXInventory -cluster production-cluster -ExportExcel
.PARAMETER folderPath
    Specify an alternate folder path where the exported data should be saved.
.EXAMPLE
    Get-ESXInventory -cluster production-cluster -ExportExcel -folderPath C:\temp
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
    $vHostList = @()
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
    } #END function

<#
----------------------------------------------------------[Execution]----------------------------------------------------------
#>

    <#
        Validate if any connected to a VIServer
    #>
    Write-Verbose ((get-date -Format G) + "`tValidate connection to to a vSphere server")

    if ($Global:DefaultViServers.Count -gt 0) {
        Clear-Host
        Write-Host "`tConnected to $Global:DefaultViServers" -ForegroundColor Green
    } else {
        Write-Error "You must be connected to a vCenter or vSphere Host before running this cmdlet."
        break
    } #END if/else

    <#
        Validate that at least one parameter was specified (-esxi, -cluster, or -datacenter
        Although all 3 can be specified, only the first is used
        Example: -esxi "host001" -cluster "test-cluster". -esxi is the first parameter
        and what will be used.
    #>
    Write-Verbose ((get-date -Format G) + "`tValidate parameters used")

    if ([string]::IsNullOrWhiteSpace($esxi) -and [string]::IsNullOrWhiteSpace($cluster) -and [string]::IsNullOrWhiteSpace($datacenter)) {
        Write-Error "You must use a parameter (-esxi, -cluster, -datacenter). Use Get-Help for more information"
        break
    } #END if

    <#
        Gather host list based on parameter used
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
                        } #END if/else
                    } #END foreach
                } #END if/else
            } #END if/else
        } else {
            Write-Verbose ((get-date -Format G) + "`tExecuting cmdlet using cluster parameter")
            Write-Host "`tGathering host list from the following Cluster(s): " (@($cluster) -join ',')

            foreach ($vClusterName in $cluster) {
                $tempList = Get-Cluster -Name $vClusterName.Trim() -ErrorAction SilentlyContinue | Get-VMHost 

                if ([string]::IsNullOrWhiteSpace($tempList)) {
                    Write-Warning "`tCluster with name $vClusterName was not found in $Global:DefaultViServers"
                } else {
                    $vHostList += $tempList | Sort-Object -Property Name
                } #END if/else
            } #END foreach
        } #END if/else
    } else { 
        Write-Verbose ((get-date -Format G) + "`tExecuting cmdlet using esxi parameter")
        Write-Host "`tGathering host list..."

        foreach($invidualHost in $esxi) {
            $vHostList += $invidualHost.Trim() | Sort-Object -Property Name
        } #END foreach
    } #END if/else

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
                } #END if/else
                Write-Verbose ((get-date -Format G) + "`t$outputFile")
            } else {
                Write-Warning "`t'$folderPath' path not found. The current location: '$currentLocation' will be used instead"
                $outputFile = $currentLocation + "\" + $outputFile
            } #END if/else
        } #END if/else
    } #END if

    if ($ExportExcel) {

        if (Get-Module -ListAvailable -Name ImportExcel) {
            Write-Verbose ((get-date -Format G) + "`tImportExcel Module available")
        } else {
            Write-Warning "`tImportExcel Module missing. Will export data to CSV file instead"
            Write-Warning "`tImportExcel Module can be installed directly from the PowerShell Gallery"
            Write-Warning "`tSee https://github.com/dfinke/ImportExcel for more information"
            $ExportExcel = $false
            $ExportCSV = $true
        } #END if/else
    } #END if

    <#
        Main code execution
    #>
    foreach ($esxihost in $vHostList) {
        $vmhost = Get-VMHost -Name $esxihost -ErrorAction SilentlyContinue

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
                Use a custom object and array to keep track of skipped
                hosts and continue to the next foreach loop
            #>
            $skiphosts = [pscustomobject]@{
                'Hostname' = $esxihost
                'Connection State' = $esxihost.ConnectionState
            } #END [PSCustomObject]
            $skipCollection += $skiphosts
            continue
        } #END if/else

        $esxcli2 = Get-EsxCli -VMHost $esxihost -V2

        <#
            Get inventory info
        #>
        Write-Host "`tGathering information from $vmhost ..."
        $hardware = $vmhost | Get-VMHostHardware -SkipAllSslCertificateChecks -WaitForAllData -ErrorAction SilentlyContinue
        $vmInfo = $vmhost | Select-Object -Property MemoryTotalGB, Build
        $mgmtIP = $vmhost | Get-VMHostNetworkAdapter | Where-Object {$_.ManagementTrafficEnabled -eq 'True'} | Select-Object -ExpandProperty IP
        $hardwarePlatfrom = $esxcli2.hardware.platform.get.Invoke()
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
            } #END if
        } #END if/ese

        <#
            Get ESXi version details
        #>
        $vmhostView = $vmhost | Get-View
        $esxiVersion = $esxcli2.system.version.get.Invoke()
    
        <#
            Use a custom object to store
	        collected data
        #>
        $outputCollection += [PSCustomObject]@{
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
        } #END [PSCustomObject]
    } #END foreach

    <#
        Display skipped hosts and their connection status
    #>
    If ($skipCollection) {
        Write-Warning "`tCheck Connection State or Host name "
        Write-Warning "`tSkipped hosts: "
        $skipCollection | Format-Table -AutoSize
    } #END if


    <#
        Output to screen
        Export data to CSV, Excel
    #>
    if ($outputCollection) {
        Write-Host "`n" "ESXi Inventory:" -ForegroundColor Green
        $outputCollection | Format-List 

        if ($ExportCSV) {
            $outputCollection | Export-Csv ($outputFile + ".csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + ".csv") "file" -ForegroundColor Green
        } #END if

        if ($ExportExcel) {
            $outputCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Inventory -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        } #END if
    } else {
        Write-Verbose ((get-date -Format G) + "`tNo information gathered")
    } #END if/else
} #END function

function Get-ESXIODevice {
<#
.SYNOPSIS
    Get ESXi vmnic* and vmhba* VMKernel device information
.DESCRIPTION
    Will get PCI/IO Device information including HCL IDs for the below VMkernel name(s): 
    Network Controller - vmnic*
    Storage Controller - vmhba*
    Graphic Device - vmgfx*
    All this can be gathered for a vSphere Cluster, Datacenter or individual ESXi host
.NOTES
    Author     : Edgar Sanchez - @edmsanchez13
    Contributor: Ariel Sanchez - @arielsanchezmor
.Link
  https://github.com/edmsanchez/vDocumentation
.INPUTS
   No inputs required
.OUTPUTS
   CSV file
   Excel file
.PARAMETER esxi
   The name(s) of the vSphere ESXi Host(s)
.EXAMPLE
    Get-ESXIODevice -esxi devvm001.lab.local
.PARAMETER cluster
   The name(s) of the vSphere Cluster(s)
.EXAMPLE
    Get-ESXIODevice -cluster production-cluster
.PARAMETER datacenter
   The name(s) of the vSphere Virtual DataCenter(s)
.EXAMPLE
    Get-ESXIODevice -datacenter vDC001
.PARAMETER ExportCSV
    Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
.EXAMPLE
   Get-ESXIODevice -cluster production-cluster -ExportCSV
.PARAMETER ExportExcel
    Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
    ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
    File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
.EXAMPLE
    Get-ESXIODevice -cluster production-cluster -ExportExcel
.PARAMETER folderPath
    Specificies an alternate folder path of where the exported file should be saved.
.EXAMPLE
    Get-ESXIODevice -cluster production-cluster -ExportExcel -folderPath C:\temp
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
    $vHostList = @()
    $date = Get-Date -format s
    $date = $date -replace ":","-"
    $outputFile = "IODevice" + $date

<#
----------------------------------------------------------[Execution]----------------------------------------------------------
#>

    <#
        Validate if any connected to a VIServer
    #>
    Write-Verbose ((get-date -Format G) + "`tValidate connection to to a vSphere server")

    if ($Global:DefaultViServers.Count -gt 0) {
        Clear-Host
        Write-Host "`tConnected to $Global:DefaultViServers" -ForegroundColor Green
    } else {
        Write-Error "You must be connected to a vCenter or vSphere Host before running this cmdlet."
        break
    } #END if/else

    <#
        Validate that at least one parameter was specified (-esxi, -cluster, or -datacenter
        Although all 3 can be specified, only the first is used
        Example: -esxi "host001" -cluster "test-cluster". -esxi is the first parameter
        and what will be used.
    #>
    Write-Verbose ((get-date -Format G) + "`tValidate parameters used")

    if ([string]::IsNullOrWhiteSpace($esxi) -and [string]::IsNullOrWhiteSpace($cluster) -and [string]::IsNullOrWhiteSpace($datacenter)) {
        Write-Error "You must use a parameter (-esxi, -cluster, -datacenter). Use Get-Help for more information"
        break
    } #END if

    <#
        Gather host list based on parameter used
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
                        } #END if/else
                    } #END foreach
                } #END if/else
            } #END if/else
        } else {
            Write-Verbose ((get-date -Format G) + "`tExecuting cmdlet using cluster parameter")
            Write-Host "`tGathering host list from the following Cluster(s): " (@($cluster) -join ',')

            foreach ($vClusterName in $cluster) {
                $tempList = Get-Cluster -Name $vClusterName.Trim() -ErrorAction SilentlyContinue | Get-VMHost 

                if ([string]::IsNullOrWhiteSpace($tempList)) {
                    Write-Warning "`tCluster with name $vClusterName was not found in $Global:DefaultViServers"
                } else {
                    $vHostList += $tempList | Sort-Object -Property Name
                } #END if/else
            } #END foreach
        } #END if/else
    } else { 
        Write-Verbose ((get-date -Format G) + "`tExecuting cmdlet using esxi parameter")
        Write-Host "`tGathering host list..."

        foreach($invidualHost in $esxi) {
            $vHostList += $invidualHost.Trim() | Sort-Object -Property Name
        } #END foreach
    } #END if/else

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
                } #END if/else
                Write-Verbose ((get-date -Format G) + "`t$outputFile")
            } else {
                Write-Warning "`t'$folderPath' path not found. The current location: '$currentLocation' will be used instead"
                $outputFile = $currentLocation + "\" + $outputFile
            } #END if/else
        } #END if/else
    } #END if

    if ($ExportExcel) {

        if (Get-Module -ListAvailable -Name ImportExcel) {
            Write-Verbose ((get-date -Format G) + "`tImportExcel Module available")
        } else {
            Write-Warning "`tImportExcel Module missing. Will export data to CSV file instead"
            Write-Warning "`tImportExcel Module can be installed directly from the PowerShell Gallery"
            Write-Warning "`tSee https://github.com/dfinke/ImportExcel for more information"
            $ExportExcel = $false
            $ExportCSV = $true
        } #END if/else
    } #END if

    <#
        Main code execution
    #>
    foreach ($esxihost in $vHostList) {
        $vmhost = Get-VMHost -Name $esxihost -ErrorAction SilentlyContinue

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
                Use a custom object and array to keep track of skipped
                hosts and continue to the next foreach loop
            #>
            $skiphosts = [pscustomobject]@{
                'Hostname' = $esxihost
                'Connection State' = $esxihost.ConnectionState
            } #END [PSCustomObject]
            $skipCollection += $skiphosts
            continue
        } #END if/else

        $esxcli2 = Get-EsxCli -VMHost $esxihost -V2

        <#
            Get IO Device info
        #>
        Write-Host "`tGathering information from $vmhost ..."
        $pciDevices = $esxcli2.hardware.pci.list.Invoke() | Where-Object {$_.VMKernelName -like "vmhba*" -OR $_.VMKernelName -like "vmnic*" -OR $_.VMKernelName -like "vmgfx*" }
     
        foreach ($pciDevice in $pciDevices) {
            $device = $vmhost | Get-VMHostPciDevice | Where-Object { $pciDevice.Address -match $_.Id }
        
            Write-Verbose ((get-date -Format G) + "`tGet driver version for: " + $pciDevice.ModuleName)
            $driverVersion = $esxcli2.system.module.get.Invoke(@{module = $pciDevice.ModuleName}) | Select-Object -ExpandProperty Version

            <#
                Get NIC Firmware version
            #>
            if ($pciDevice.VMKernelName -like 'vmnic*') {
                Write-Verbose ((get-date -Format G) + "`tGet Firmware version for: " + $pciDevice.VMKernelName)
                $vmnicDetail = $esxcli2.network.nic.get.Invoke(@{nicname = $pciDevice.VMKernelName})
                $firmwareVersion = $vmnicDetail.DriverInfo.FirmwareVersion
            
                <#
                    Get NIC driver VIB package version
                #>
                Write-Verbose ((get-date -Format G) + "`tGet VIB details for: " + $pciDevice.ModuleName)
                $driverVib = $esxcli2.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object {$_.Name -eq "net-"+$vmnicDetail.DriverInfo.Driver}
                $vibName = $driverVib.Name
                $vibVersion = $driverVib.Version

            <#
                Skip if VMkernnel is vmhba* 
                Can't get HBA Firmware from Powercli at the moment
                only through SSH or using Putty Plink+PowerCli
            #>
            } elseif ($pciDevice.VMKernelName -like 'vmhba*') {
                Write-Verbose ((get-date -Format G) + "`tSkip Firmware version check for: " + $pciDevice.DeviceName)
                $firmwareVersion = $null

                <#
                    Get HBA driver VIB package version
                #>
                Write-Verbose ((get-date -Format G) + "`tGet VIB deatils for: " + $pciDevice.ModuleName)
                $vibName = $pciDevice.ModuleName -replace "_", "-"
                $driverVib = $esxcli2.software.vib.list.Invoke() | Select-Object -Property Name,Version | Where-Object {$_.Name -eq "scsi-"+$VibName -or $_.Name -eq "sata-"+$VibName -or $_.Name -eq $VibName}
                $vibName = $driverVib.Name
                $vibVersion = $driverVib.Version
            } else {
                Write-Verbose ((get-date -Format G) + "`tSkipping: " + $pciDevice.DeviceName)
                $firmwareVersion = $null
                $vibName = $null
                $vibVersion = $null
            } #END if/else
       
            <#
	            Use a custom object to store
	            collected data
            #>
            $outputCollection += [PSCustomObject]@{
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
            } #END [PSCustomObject]
        } #END foreach
    } #END foreach

    <#
        Display skipped hosts and their connection status
    #>
    If ($skipCollection) {
        Write-Warning "`tCheck Connection State or Host name "
        Write-Warning "`tSkipped hosts: "
        $skipCollection | Format-Table -AutoSize
    } #END if

    <#
        Output to screen
        Export data to CSV, Excel
    #>
    if ($outputCollection) {
        Write-Host "`n" "ESXi IO Device:" -ForegroundColor Green
        $outputCollection | Format-List 

        if ($ExportCSV) {
            $outputCollection | Export-Csv ($outputFile + ".csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + ".csv") "file" -ForegroundColor Green
        } #END if

        if ($ExportExcel) {
            $outputCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname IO_Device -NoNumberConversion VID,DID,SVID,SSID -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        } #END if
    } else {
        Write-Verbose ((get-date -Format G) + "`tNo information gathered")
    } #END if/else
} #END function

function Get-ESXNetworking {
<#
.SYNOPSIS
    Get ESXi Networking Details.
.DESCRIPTION
    Will get Physical Adapters, Virtual Switches, and Port Groups
    All this can be gathered for a vSphere Cluster, Datacenter or individual ESXi host
.NOTES
    Author     : Edgar Sanchez - @edmsanchez13
    Contributor: Ariel Sanchez - @arielsanchezmor
.Link
  https://github.com/edmsanchez/vDocumentation
.INPUTS
   No inputs required
.OUTPUTS
   CSV file
   Excel file
.PARAMETER esxi
   The name(s) of the vSphere ESXi Host(s)
.EXAMPLE
    Get-ESXNetworking -esxi devvm001.lab.local
.PARAMETER cluster
   The name(s) of the vSphere Cluster(s)
.EXAMPLE
    Get-ESXNetworking -cluster production-cluster
.PARAMETER datacenter
   The name(s) of the vSphere Virtual DataCenter(s)
.EXAMPLE
    Get-ESXNetworking -datacenter vDC001
.PARAMETER ExportCSV
    Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
.EXAMPLE
   Get-ESXNetworking -cluster production-cluster -ExportCSV
.PARAMETER ExportExcel
    Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
    ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
    File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
.EXAMPLE
    Get-ESXNetworking -cluster production-cluster -ExportExcel
.PARAMETER PhysicalAdapters
    Default switch to get Physical Adapter details including uplinks to vswitch and CDP/LLDP Information
    This is default option that will get processed if no switch parameter is provided.
.EXAMPLE
    Get-ESXNetworking -cluster production-cluster -PhysicalAdapters
.PARAMETER VMkernelAdapters
    Switch to get VMkernel Adapter details including Enabled services
.EXAMPLE
    Get-ESXNetworking -cluster production-cluster -VMkernelAdapters
.PARAMETER VirtualSwitches
    Switch to get Virtual switches details including port groups.
.EXAMPLE
    Get-ESXNetworking -cluster production-cluster -VirtualSwitches
.PARAMETER folderPath
    Specificies an alternate folder path of where the exported file should be saved.
.EXAMPLE
    Get-ESXNetworking -cluster production-cluster -ExportExcel -folderPath C:\temp
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
        [switch]$VirtualSwitches,
        [switch]$VMkernelAdapters,
        [switch]$PhysicalAdapters,
        $folderPath
    )

    $PhysicalAdapterCollection = @()
    $VMkernelAdapterCollection = @()
    $VirtualSwitchesCollection = @()
    $skipCollection = @()
    $vHostList = @()
    $date = Get-Date -format s
    $date = $date -replace ":","-"
    $outputFile = "Networking" + $date

<#
----------------------------------------------------------[Execution]----------------------------------------------------------
#>

    <#
        Validate if any connected to a VIServer
    #>
    Write-Verbose ((get-date -Format G) + "`tValidate connection to to a vSphere server")

    if ($Global:DefaultViServers.Count -gt 0) {
        Clear-Host
        Write-Host "`tConnected to $Global:DefaultViServers" -ForegroundColor Green
    } else {
        Write-Error "You must be connected to a vCenter or vSphere Host before running this cmdlet."
        break
    } #END if/else

    <#
        Validate that at least one parameter was specified (-esxi, -cluster, or -datacenter
        Although all 3 can be specified, only the first is used
        Example: -esxi "host001" -cluster "test-cluster". -esxi is the first parameter
        and what will be used.
    #>
    Write-Verbose ((get-date -Format G) + "`tValidate parameters used")

    if ([string]::IsNullOrWhiteSpace($esxi) -and [string]::IsNullOrWhiteSpace($cluster) -and [string]::IsNullOrWhiteSpace($datacenter)) {
        Write-Error "You must use a parameter (-esxi, -cluster, -datacenter). Use Get-Help for more information"
        break
    } #END if

    <#
        Gather host list based on parameter used
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
                        } #END if/else
                    } #END foreach
                } #END if/else
            } #END if/else
        } else {
            Write-Verbose ((get-date -Format G) + "`tExecuting cmdlet using cluster parameter")
            Write-Host "`tGathering host list from the following Cluster(s): " (@($cluster) -join ',')

            foreach ($vClusterName in $cluster) {
                $tempList = Get-Cluster -Name $vClusterName.Trim() -ErrorAction SilentlyContinue | Get-VMHost 

                if ([string]::IsNullOrWhiteSpace($tempList)) {
                    Write-Warning "`tCluster with name $vClusterName was not found in $Global:DefaultViServers"
                } else {
                    $vHostList += $tempList | Sort-Object -Property Name
                } #END if/else
            } #END foreach
        } #END if/else
    } else { 
        Write-Verbose ((get-date -Format G) + "`tExecuting cmdlet using esxi parameter")
        Write-Host "`tGathering host list..."

        foreach($invidualHost in $esxi) {
            $vHostList += $invidualHost.Trim() | Sort-Object -Property Name
        } #END foreach
    } #END if/else

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
                } #END if/else
                Write-Verbose ((get-date -Format G) + "`t$outputFile")
            } else {
                Write-Warning "`t'$folderPath' path not found. The current location: '$currentLocation' will be used instead"
                $outputFile = $currentLocation + "\" + $outputFile
            } #END if/else
        } #END if/else
    } #END if

    if ($ExportExcel) {

        if (Get-Module -ListAvailable -Name ImportExcel) {
            Write-Verbose ((get-date -Format G) + "`tImportExcel Module available")
        } else {
            Write-Warning "`tImportExcel Module missing. Will export data to CSV file instead"
            Write-Warning "`tImportExcel Module can be installed directly from the PowerShell Gallery"
            Write-Warning "`tSee https://github.com/dfinke/ImportExcel for more information"
            $ExportExcel = $false
            $ExportCSV = $true
        } #END if/else
    } #END if

     <#
        Validate that a cmdlet switch was used. Options are
        -PhysicalAdapters, -VMkernelAdapters, -VirtualSwitches. 
        If none was specified then it will default to -PhysicalAdapters
    #>
    Write-Verbose ((get-date -Format G) + "`tValidate cmdlet switches")

    if ($PhysicalAdapters -or $VMkernelAdapters -or $VirtualSwitches) {
        Write-Verbose ((get-date -Format G) + "`tA cmdlet switch was specified")
    } else {
        Write-Warning "`tA cmdlet switch was not specified"
        Write-Warning "`tYou must use one of the following cmdlet swicth: -PhysicalAdapters, -VMkernelAdapters, -VirtualSwitches. Use Get-Help for more information"
        Write-Warning "`tWill proceed with default of -PhysicalAdapters cmdlet switch"
        $PhysicalAdapters = $true
    } #END if/else
   
    <#
        Main code execution
    #>
    foreach ($esxihost in $vHostList) {
        $vmhost = Get-VMHost -Name $esxihost -ErrorAction SilentlyContinue

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
                Use a custom object and array to keep track of skipped
                hosts and continue to the next foreach loop
            #>
            $skiphosts = [pscustomobject]@{
                'Hostname' = $esxihost
                'Connection State' = $esxihost.ConnectionState
            } #END [PSCustomObject]
            $skipCollection += $skiphosts
            continue
        } #END if/else

        $esxcli2 = Get-EsxCli -VMHost $esxihost -V2

        <#
            Get physical adapter details
            Default whether the switch is specified or not
        #>
        if ($PhysicalAdapters) {
            Write-Host "`tGathering physical adapter details from $vmhost ..."
            $vmnics = $vmhost | Get-VMHostNetworkAdapter -Physical | Select-Object Name, Mac, Mtu

            foreach($nic in $vmnics) {
                Write-Verbose ((get-date -Format G) + "`tGet device details for: " + $nic.Name)
                $pciList = $esxcli2.hardware.pci.list.Invoke() | Where-Object {$_.VMKernelName -eq $nic.Name}
                $nicList = $esxcli2.network.nic.list.Invoke() | Where-Object {$_.Name -eq $nic.Name}

                <#
                    Get uplink vSwitch, check standard
                    vSwitch first then Distributed.
                #>
                $vSwitch = $esxcli2.network.vswitch.standard.list.Invoke() | Where-Object {$_.uplinks -contains $nic.Name}

                if ($vSwitch) {
                    Write-Verbose ((get-date -Format G) + "`tUplinks to vswitch: " + $vSwitch.Name)
                } else {
                    $vSwitch = $esxcli2.network.vswitch.dvs.vmware.list.Invoke() | Where-Object {$_.uplinks -contains $nic.Name}
                    Write-Verbose ((get-date -Format G) + "`tUplinks to vswitch: " + $vSwitch.Name)
                } #END if/else

                <#
                    Get Device Discovery Protocol CDP/LLDP
                #>
                Write-Verbose ((get-date -Format G) + "`tGet Device Discovery Protocol for: " + $nic.Name)
                $esxiHostView = $vmhost | Get-View 
                $networkSystem = $esxiHostView.Configmanager.Networksystem
                $networkView = Get-View $networkSystem
                $networkViewInfo = $networkView.QueryNetworkHint($nic.Name)

                If ($networkViewInfo.connectedswitchport -ne $null) {
                    Write-Verbose ((get-date -Format G) + "`tDevice Discovery Protocol: CDP")
                    $ddp = "CDP"
                    $ddpExtended = $networkViewInfo.connectedswitchport
                    $ddpDevID = $ddpExtended.DevId
                    $ddpDevIP = $ddpExtended.Address
                    $ddpDevPortId = $ddpExtended.PortId
                } else {
                    Write-Verbose ((get-date -Format G) + "`tCDP not found")

                    if ($networkViewInfo.lldpinfo -ne $null) {
                        Write-Verbose ((get-date -Format G) + "`tDevice Discovery Protocol: LLDP")
                        $ddp = "LLDP"
                        $ddpDevID = $networkViewInfo.lldpinfo.Parameter | Where-Object {$_.Key -eq "System Name"} | Select-Object -ExpandProperty Value  
                        $ddpDevIP = $networkViewInfo.lldpinfo.Parameter | Where-Object {$_.Key -eq "Management Address"} | Select-Object -ExpandProperty Value  
                        $ddpDevPortId = $networkViewInfo.lldpinfo.Portid
                    } else {
                        Write-Verbose ((get-date -Format G) + "`tLLDP not found")
                        $ddp = ""
                        $ddpDevID = ""
                        $ddpDevIP = ""
                        $ddpDevPortId = ""
                    } #END if/else
                } #END if/else

                <#
	                Use a custom object to store
	                collected data
                #>
                $PhysicalAdapterCollection += [PSCustomObject]@{
                    'Hostname' = $vmhost
                    'Name' = $nic.Name
                    'Slot Description' = $pciList.SlotDescription
                    'Device' = $nicList.Description
                    'Duplex' = $nicList.Duplex
                    'Link' = $nicList.Link
                    'MAC' = $nic.Mac
                    'MTU' = $nicList.MTU
                    'Speed' = $nicList.Speed
                    'vSwitch' = $vSwitch.Name
                    'vSwitch MTU' = $vSwitch.MTU
                    'Discovery Protocol' = $ddp
                    'Device ID' = $ddpDevID
                    'Device IP' = $ddpDevIP
                    'Port' = $ddpDevPortId
                } #END [PSCustomObject]
            } #END foreach
        } #END if

        <#
            Get VMkernel adapter details
            if $VMkernelAdapters switch was used
        #>
        if ($VMkernelAdapters) {
            Write-Host "`tGathering VMkernel adapter details from $vmhost ..."
            $vmnics = $vmhost | Get-VMHostNetworkAdapter -VMKernel

            foreach($nic in $vmnics) {
                Write-Verbose ((get-date -Format G) + "`tGet device details for: " + $nic.Name)

                <#
                    Get VMkernel adapter enabled services
                #>
                $enabledServices = @()
                Write-Verbose ((get-date -Format G) + "`tGet Enabled Services for: " + $nic.Name)

                if ($nic.VMotionEnabled) {
                    $enabledServices += "vMotion"
                } #END if

                if ($nic.FaultToleranceLoggingEnabled) {
                    $enabledServices += "Fault Tolerance logging"
                } #END if

                if ($nic.ManagementTrafficEnabled) {
                    $enabledServices += "Management"
                } #END if

                if ($nic.VsanTrafficEnabled) {
                    $enabledServices += "vSAN"
                } #END if

                <#
                    Get VMkernel adapter associated vSwitch,
                    VLAN ID, and NIC Teaming Policy
                #>
                Write-Verbose ((get-date -Format G) + "`tGet Port Group details for: " + $nic.PortGroupName)
                $interfaceList = $esxcli2.network.ip.interface.list.Invoke() | Where-Object {$_.Name -eq $nic.Name}
                $portVLanId = Get-VirtualPortGroup -VMhost $vmhost -Name $nic.PortGroupName | Where-Object {$_.VirtualSwitchName -eq $interfaceList.Portset} | Select-Object -ExpandProperty VLanId
                $portGroupTeam = Get-VirtualPortGroup -VMhost $vmhost -Name $nic.PortGroupName | Where-Object {$_.VirtualSwitchName -eq $interfaceList.Portset} | Get-NicTeamingPolicy

                <#
                    Get vSwitch MTU using Active adapter associated with the VMKernel Port
                    test against both Standard and Distributed Switch.
                #>
                Write-Verbose ((get-date -Format G) + "`tGet vSwitch MTU for: " + $interfaceList.Portset)
                $vSwitch = $esxcli2.network.vswitch.standard.list.Invoke() | Where-Object {$_.uplinks -contains ($portGroupTeam.ActiveNic | Select-Object -First 1)}

                if ($vSwitch) {
                    Write-Verbose ((get-date -Format G) + "`tStandard vSwitch")
                } else {
                    $vSwitch = $esxcli2.network.vswitch.dvs.vmware.list.Invoke() | Where-Object {$_.uplinks -contains ($portGroupTeam.ActiveNic | Select-Object -First 1)}
                    Write-Verbose ((get-date -Format G) + "`tDistributed vSwitch")
                } #END if/else

                <#
                    Get TCP/IP Stack details
                #>
                Write-Verbose ((get-date -Format G) + "`tGet VMkernel TCP/IP configuration...")
                $tcpipConfig = $vmhost | Get-VMHostNetwork

                if ($tcpipConfig.VirtualNic.Name -contains $nic.Name) {
                    $vmkGateway = $tcpipConfig.VMKernelGateway
                    $dnsAddress = $tcpipConfig.DnsAddress
                } else {
                    $vmkGateway = ""
                    $dnsAddress = ""
                } #END if/else
                                    
                <#
                    Use a custom object to store
	                collected data
                #>
                $VMkernelAdapterCollection += [PSCustomObject]@{
                    'Hostname' = $vmhost
                    'Name' = $nic.Name
                    'MAC' = $nic.Mac
                    'MTU' = $nic.MTU
                    'IP' = $nic.IP
                    'Subnet Mask' = $nic.SubnetMask
                    'TCP/IP Stack' = $interfaceList.NetstackInstance
                    'Default Gateway' = $vmkGateway
                    'DNS' = (@($dnsAddress) -join ',')
                    'PortGroup Name' = $nic.PortGroupName
                    'VLAN ID' = $portVLanId
                    'Enabled Services' = (@($enabledServices) -join ',')
                    'vSwitch' = $interfaceList.Portset
                    'vSwitch MTU' = $vSwitch.MTU
                    'Active adapters' = (@($PortGroupTeam.ActiveNic) -join ',')
                    'Standby adapters' = (@($PortGroupTeam.StandbyNic) -join ',')
                    'Unused adapters' = (@($PortGroupTeam.UnusedNic) -join ',')
                } #END [PSCustomObject]
            } #END foreach
        } #END if

        <#
            Get virtual switches details
            if -VirtualSwitches switch was used
        #>
        if ($VirtualSwitches) {
            Write-Host "`tGathering virtual switches details from $vmhost ..."

            <#
                Get standard switch details
            #>
            $StdvSwitch = Get-VirtualSwitch -VMHost $vmhost -Standard

            foreach ($vSwitch in $StdvSwitch) {
                Write-Verbose ((get-date -Format G) + "`tGet standard switch details for: " + $vSwitch.Name)

                <#
                    Get PortGroup details,
                    Security Policy, and Teaming Policy
                #>
                $portGroups = $vSwitch | Get-VirtualPortGroup

                if ($portGroups) {

                    foreach ($port in $portGroups) {
                    
                        Write-Verbose ((get-date -Format G) + "`tGet Port Group details for: " + $port.Name)
                        $portGroupTeam = $port | Get-NicTeamingPolicy
                        $portGroupSecurity = $port | Get-SecurityPolicy
                        $portGroupPolicy = @()

                        if ($portGroupSecurity.AllowPromiscuous) {
                            $portGroupPolicy += "Accept"
                        } else {
                            $portGroupPolicy += "Reject"
                        } #END if/else

                        if ($portGroupSecurity.MacChanges) {
                            $portGroupPolicy += "Accept"
                        } else {
                            $portGroupPolicy += "Reject"
                        } #END if/else

                        if ($portGroupSecurity.ForgedTransmits) {
                            $portGroupPolicy += "Accept"
                        } else {
                            $portGroupPolicy += "Reject"
                        } #END if/else

                        <#
                            Use a custom object to store
                            collected data
                        #>
                        $VirtualSwitchesCollection += [PSCustomObject]@{
                            'Hostname' = $vmhost
                            'Type' = "Standard"
                            'Version' = $null
                            'Name' = $vSwitch.Name
                            'Uplink/ConnectedAdapters' = (@($vSwitch.Nic) -join ',')
                            'PortGroup' = $port.Name
                            'VLAN ID' = $port.VLanId
                            'Active adapters' = (@($PortGroupTeam.ActiveNic) -join ',')
                            'Standby adapters' = (@($PortGroupTeam.StandbyNic) -join ',')
                            'Unused adapters' = (@($PortGroupTeam.UnusedNic) -join ',')
                            'Security Promiscuous/MacChanges/ForgedTransmits' = (@($portGroupPolicy) -join '/')                        
                        } #END [PSCustomObject]
                    } #END foreach
                } else {
                    <#
                        Use a custom object to store
                        collected data
                    #>
                    $VirtualSwitchesCollection += [PSCustomObject]@{
                        'Hostname' = $vmhost
                        'Type' = "Standard"
                        'Version' = $null
                        'Name' = $vSwitch.Name
                        'Uplink/ConnectedAdapters' = (@($vSwitch.Nic) -join ',')
                        'PortGroup' = $null
                        'VLAN ID' = $null
                        'Active adapters' = $null
                        'Standby adapters' = $null
                        'Unused adapters' = $null
                        'Security Promiscuous/MacChanges/ForgedTransmits' = $null
                    } #END [PSCustomObject]
                } #END if/else
            } #END foreach

            <#
                Get distributed switch details
            #>
            $dVSwitch = Get-VDSwitch -VMHost $vmhost

            foreach($vSwitch in $dVSwitch) {
                Write-Verbose ((get-date -Format G) + "`tGet standard switch details for: " + $vSwitch.Name)
                
                <#
                    Get PortGroup details,
                    Security Policy, and Teaming Policy
                #>
                $portGroups = $vSwitch | Get-VDPortgroup
                $vSwitchUplink = @()
                if ($portGroups) {

                    <#
                        Get Distribute Switch Uplinks
                    #>
                    $dvUplinkPG = $portGroups | Where-Object {$_.IsUplink}
                    $vSwitchUplink = $vSwitch | Get-VDPort | Where-Object {$_.Portgroup -eq $dvUplinkPG}

                    $uplinkConnected = @()
                    foreach ($uplink in $vSwitchUplink) {
                        $uplinkConnected += $uplink.Name + "," + $uplink.ConnectedEntity
                    } #END foreach
                    
                    foreach($port in $portGroups | Where-Object {!$_.IsUplink}) {

                        Write-Verbose ((get-date -Format G) + "`tGet Port Group details for: " + $port.Name)
                        $portGroupTeam = $port | Get-VDUplinkTeamingPolicy
                        $portGroupSecurity = $port | Get-VDSecurityPolicy
                        $portGroupPolicy = @()
                        
                        if ($portGroupSecurity.AllowPromiscuous) {
                            $portGroupPolicy += "Accept"
                        } else {
                            $portGroupPolicy += "Reject"
                        } #END if/else

                        if ($portGroupSecurity.MacChanges) {
                            $portGroupPolicy += "Accept"
                        } else {
                            $portGroupPolicy += "Reject"
                        } #END if/else

                        if ($portGroupSecurity.ForgedTransmits) {
                            $portGroupPolicy += "Accept"
                        } else {
                            $portGroupPolicy += "Reject"
                        } #END if/else

                        <#
                            Use a custom object to store
                            collected data
                        #>
                        $VirtualSwitchesCollection += [PSCustomObject]@{
                            'Hostname' = $vmhost
                            'Type' = "Distributed"
                            'Version' = $vSwitch.Version
                            'Name' = $vSwitch.Name
                            'Uplink/ConnectedAdapters' = (@($uplinkConnected) -join '/')
                            'PortGroup' = $port.Name
                            'VLAN ID' = $port.VlanConfiguration.VlanId
                            'Active adapters' = (@($PortGroupTeam.ActiveUplinkPort) -join ',')
                            'Standby adapters' = (@($PortGroupTeam.StandbyUplinkPort) -join ',')
                            'Unused adapters' = (@($PortGroupTeam.UnusedUplinkPort) -join ',')
                            'Security Promiscuous/MacChanges/ForgedTransmits' = (@($portGroupPolicy) -join '/')
                        } #END [PSCustomObject]
                    } #END foreach
                } else {
                    <#
                        Use a custom object to store
                        collected data
                    #>
                    $VirtualSwitchesCollection += [PSCustomObject]@{
                        'Hostname' = $vmhost
                        'Type' = "Distributed"
                        'Version' = $vSwitch.Version
                        'Name' = $vSwitch.Name
                        'Uplink/ConnectedAdapters' = (@($uplinkConnected) -join '/')
                        'PortGroup' = $port.Name
                        'VLAN ID' = $port.VlanConfiguration.VlanId
                        'Active adapters' = (@($PortGroupTeam.ActiveUplinkPort) -join ',')
                        'Standby adapters' = (@($PortGroupTeam.StandbyUplinkPort) -join ',')
                        'Unused adapters' = (@($PortGroupTeam.UnusedUplinkPort) -join ',')
                        'Security Promiscuous/MacChanges/ForgedTransmits' = (@($portGroupPolicy) -join '/')
                    } #END [PSCustomObject]
                } #END if/else
            } #END foreach
        } #END if
    } #END foreach       

    <#
        Display skipped hosts and their connection status
    #>
    If ($skipCollection) {
        Write-Warning "`tCheck Connection State or Host name "
        Write-Warning "`tSkipped hosts: "
        $skipCollection | Format-Table -AutoSize
    } #END if

    <#
        Validate output arrays
    #>
    if ($PhysicalAdapterCollection -or $VMkernelAdapterCollection -or $VirtualSwitchesCollection) {
        Write-Verbose ((get-date -Format G) + "`tInformation gathered")
    } else {
        Write-Verbose ((get-date -Format G) + "`tNo information gathered")
    } #END if/else

    <#
        Output to screen
        Export data to CSV, Excel
    #>
    if ($PhysicalAdapterCollection) {
        Write-Host "`n" "ESXi Physical Adapters:" -ForegroundColor Green
        $PhysicalAdapterCollection | Format-List

        if ($ExportCSV) {
            $PhysicalAdapterCollection | Export-Csv ($outputFile + "PhysicalAdapters.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "PhysicalAdapters.csv") "file" -ForegroundColor Green
        } #END if

        if ($ExportExcel) {
            $PhysicalAdapterCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Physical_Adapters -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        } #END if
    } #END if

    if ($VMkernelAdapterCollection) {
        Write-Host "`n" "ESXi VMkernel Adapters:" -ForegroundColor Green
        $VMkernelAdapterCollection | Format-List

        if ($ExportCSV) {
            $VMkernelAdapterCollection | Export-Csv ($outputFile + "VMkernelAdapters.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "VMkernelAdapters.csv") "file" -ForegroundColor Green
        } #END if

        if ($ExportExcel) {
            $VMkernelAdapterCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname VMkernel_Adapters -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        } #END if
    } #END if

    if ($VirtualSwitchesCollection) {
        Write-Host "`n" "ESXi Virtual Switches:" -ForegroundColor Green
        $VirtualSwitchesCollection | Format-List

        if ($ExportCSV) {
            $VirtualSwitchesCollection | Export-Csv ($outputFile + "VirtualSwitches.csv") -Force -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "VirtualSwitches.csv") "file" -ForegroundColor Green
        } #END if

        if ($ExportExcel) {
            $VirtualSwitchesCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Virtual_Switches -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        } #END if
    } #END if $VirtualSwitchesCollection.Count
} #END function
