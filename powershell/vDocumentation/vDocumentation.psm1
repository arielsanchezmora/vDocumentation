function Get-ESXInventory {
    <#
    .SYNOPSIS
        Get basic ESXi host information
    .DESCRIPTION
        Will get inventory information for a vSphere Cluster, Datacenter or individual ESXi host
        The following is gathered:
        Hostname, Management IP, RAC IP, ESXi Version information, Hardware information
        and Host configuration
    .NOTES
        Author     : Edgar Sanchez - @edmsanchez13
        Contributor: Ariel Sanchez - @arielsanchezmor
        Get-VMHostWSManInstance Function by Carter Shanklin - @cshanklin
        Downloaded from: http://poshcode.org/?show=928
    .Link
      https://github.com/arielsanchezmora/vDocumentation
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
       The name(s) of the vSphere Virtual Datacenter(s).
    .EXAMPLE
        Get-ESXInventory -datacenter vDC001
        Get-ESXInventory -datacenter "all vdc" will gather all hosts in vCenter(s). This is the default if no Parameter (-esxi, -cluster, or -datacenter) is specified. 
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
    .PARAMETER Hardware
        Switch to get Hardware inventory
    .EXAMPLE
        Get-ESXInventory -cluster production-cluster -Hardware
    .PARAMETER Configuration
        Switch to get system configuration details
    .EXAMPLE
        Get-ESXInventory -cluster production-cluster -Configuration
    .PARAMETER folderPath
        Specify an alternate folder path where the exported data should be saved.
    .EXAMPLE
        Get-ESXInventory -cluster production-cluster -ExportExcel -folderPath C:\temp
    .PARAMETER PassThru
        Switch to return object to command line
    .EXAMPLE
        Get-ESXInventory -esxi 192.168.1.100 -Hardware -PassThru
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
            [switch]$Hardware,
            [switch]$Configuration,
            [switch]$PassThru,
            $folderPath
        )
    
        $hardwareCollection = @()
        $configurationCollection = @()
        $outputCollection = @()
        $skipCollection = @()
        $vHostList = @()
        $ReturnCollection = @()
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
            Write-Verbose ((Get-Date -Format G) + "`tGet-VMHostWSManInstance    Started execution")
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
        Write-Verbose ((Get-Date -Format G) + "`tGet-VMHostWSManInstance    Finished execution")
        } #END function
    
    <#
    ----------------------------------------------------------[Execution]----------------------------------------------------------
    #>
    
        <#
            Validate if any connected to a VIServer
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate connection to to a vSphere server")
    
        if ($Global:DefaultViServers.Count -gt 0) {
            Write-Host "`tConnected to $Global:DefaultViServers" -ForegroundColor Green
        } else {
            Write-Error "You must be connected to a vSphere server before running this cmdlet."
            break
        } #END if/else
    
        <#
            Validate if a parameter was specified (-esxi, -cluster, or -datacenter)
            Although all 3 can be specified, only the first is used
            Example: -esxi "host001" -cluster "test-cluster". -esxi is the first parameter
            and what will be used.
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate parameters used")
    
        if ([string]::IsNullOrWhiteSpace($esxi) -and [string]::IsNullOrWhiteSpace($cluster) -and [string]::IsNullOrWhiteSpace($datacenter)) {
            Write-Verbose ((Get-Date -Format G) + "`tA parameter (-esxi, -cluster, -datacenter) was not specified. Will gather all hosts")
            $datacenter = "all vdc"
        } #END if
    
        <#
            Gather host list based on parameter used
        #>
        Write-Verbose ((Get-Date -Format G) + "`tGather host list")
    
        if ([string]::IsNullOrWhiteSpace($esxi)) {
            Write-Verbose ((Get-Date -Format G) + "`t-esxi parameter is Null or Empty")
    
            if ([string]::IsNullOrWhiteSpace($cluster)) {
                Write-Verbose ((Get-Date -Format G) + "`t-cluster parameter is Null or Empty")
    
                if ([string]::IsNullOrWhiteSpace($datacenter)) {
                    Write-Verbose ((Get-Date -Format G) + "`t-datacenter parameter is Null or Empty")
    
                } else {
                    Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using datacenter parameter")
    
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
                Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using cluster parameter")
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
            Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using esxi parameter")
            Write-Host "`tGathering host list..."
    
            foreach($invidualHost in $esxi) {
                $vHostList += $invidualHost.Trim() | Sort-Object -Property Name
            } #END foreach
        } #END if/else
    
        <#
            Validate export switches,
            folder path and dependencies
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate export switches and folder path")
    
        if ($ExportCSV -or $ExportExcel) {
            $currentLocation = (Get-Location).Path
    
            if ([string]::IsNullOrWhiteSpace($folderPath)) {
                Write-Verbose ((Get-Date -Format G) + "`t-folderPath parameter is Null or Empty")
                Write-Warning "`tFolder Path (-folderPath) was not specified for saving exported data. The current location: '$currentLocation' will be used"
                $outputFile = $currentLocation + "\" + $outputFile
            } else {
    
                if (Test-Path $folderPath) {
                    Write-Verbose ((Get-Date -Format G) + "`t'$folderPath' path found")
                    $lastCharsOfFolderPath = $folderPath.Substring($folderPath.Length - 1)
    
                    if ($lastCharsOfFolderPath -eq "\" -or $lastCharsOfFolderPath -eq "/") {
                        $outputFile = $folderPath + $outputFile
                    } else {
                        $outputFile = $folderPath + "\" + $outputFile
                    } #END if/else
                    Write-Verbose ((Get-Date -Format G) + "`t$outputFile")
                } else {
                    Write-Warning "`t'$folderPath' path not found. The current location: '$currentLocation' will be used instead"
                    $outputFile = $currentLocation + "\" + $outputFile
                } #END if/else
            } #END if/else
        } #END if
    
        if ($ExportExcel) {
    
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Write-Verbose ((Get-Date -Format G) + "`tImportExcel Module available")
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
            -Hardware, -Configuration. By default all are executed
            unless one is specified. 
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate cmdlet switches")
    
        if ($Hardware -or $Configuration) {
            Write-Verbose ((Get-Date -Format G) + "`tA cmdlet switch was specified")
        } else {
            Write-Verbose ((Get-Date -Format G) + "`tA cmdlet switch was not specified")
            Write-Verbose ((Get-Date -Format G) + "`tWill execute all (-Hardware -Configuration)")
            $Hardware = $true
            $Configuration = $true
        } #END if/else
    
        <#
            Initialize varibles used for -Configuraiton switch
        #>
        if ($Configuration) {
            Write-Verbose ((Get-Date -Format G) + "`tInitializing -Configuration cmdlet switch variables...")
            $serviceInstance = Get-View ServiceInstance
            $licenseManager = Get-View $ServiceInstance.Content.LicenseManager
            $licenseManagerAssign = Get-View $LicenseManager.LicenseAssignmentManager
        }
    
        <#
            Main code execution
        #>
        foreach ($esxihost in $vHostList) {
            $vmhost = Get-VMHost -Name $esxihost -ErrorAction SilentlyContinue
    
            <#
                Skip if ESXi host is not in a Connected
                or Maintenance ConnectionState
            #>
            Write-Verbose ((Get-Date -Format G) + "`t$vmhost Connection State: " + $vmhost.ConnectionState)
    
            if ($vmhost.ConnectionState -eq "Connected" -or $vmhost.ConnectionState -eq "Maintenance") {
                <#
                    Do nothing - ESXi host is reachable
                #>
            } else {
                <#
                    Use a custom object and array to keep track of skipped
                    hosts and continue to the next foreach loop
                #>
                $skipCollection += [pscustomobject]@{
                    'Hostname' = $esxihost
                    'Connection State' = $esxihost.ConnectionState
                } #END [PSCustomObject]
                continue
            } #END if/else
    
            $esxcli2 = Get-EsxCli -VMHost $esxihost -V2
            $hostHardware = $vmhost | Get-VMHostHardware -WaitForAllData -SkipAllSslCertificateChecks -ErrorAction SilentlyContinue
    
            <#
                Get ESXi version details
            #>
            $vmhostView = $vmhost | Get-View
            $esxiVersion = $esxcli2.system.version.get.Invoke()
                    
            <#
                Get Hardware invetory details
            #>
            if ($Hardware) {
                Write-Host "`tGathering Hardware inventory from $vmhost ..."
                $mgmtIP = $vmhost | Get-VMHostNetworkAdapter | Where-Object {$_.ManagementTrafficEnabled -eq 'True'} | Select-Object -ExpandProperty IP
                $hardwarePlatfrom = $esxcli2.hardware.platform.get.Invoke()
    
                <#
                    Get RAC IP, and Firmware
                    Try with -class OMC_IPMIIPProtocolEndpoint First
                    Else try with -class CIM_IPProtocolEndpoint
                #>
                Write-Verbose ((Get-Date -Format G) + "`tGet-VMHostWSManInstance using OMC_IPMIIPProtocolEndpoint class")
                $rac = Get-VMHostWSManInstance -VMHost $vmhost -class OMC_IPMIIPProtocolEndpoint -ignoreCertFailures -ErrorAction SilentlyContinue
    
                if ($rac.Name) {
                    $racIP = $rac.IPv4Address
                } else { 
                    Write-Verbose ((Get-Date -Format G) + "`tGet-VMHostWSManInstance using CIM_IPProtocolEndpoint class")
                    $rac = Get-VMHostWSManInstance -VMHost $vmhost -class CIM_IPProtocolEndpoint -ignoreCertFailures -ErrorAction SilentlyContinue
                    
                    if ($rac.Name) {
                        $racIP = $rac | Where-Object {$_.Name -match "Management Controller IP"} | Select-Object -ExpandProperty IPv4Address
                        Write-Verbose ((Get-Date -Format G) + "`tRAC IP gathered using CIM_IPProtocolEndpoint class")
                    } #END if
                } #END if/ese
    
                if ($bmc = $vmhost.ExtensionData.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo | Where-Object {$_.Name -match "BMC Firmware"}) {
                    $bmcFirmware = (($bmc.Name -split "firmware")[1]) -split " " | Select-Object -Last 1
                } else {
                    $bmcFirmware = $null
                } #END if/else
    
                <#
                    Use a custom object to store
                    collected data
                #>
                $hardwareCollection += [PSCustomObject]@{
                    'Hostname' = $vmhost
                    'Management IP' = $mgmtIP
                    'RAC IP' = $racIP
                    'RAC Firmware' = $bmcFirmware
                    'Product' = $vmhostView.Config.Product.Name
                    'Version' = $vmhostView.Config.Product.Version
                    'Build' = $vmhost.Build
                    'Update' = $esxiVersion.Update
                    'Patch' = $esxiVersion.Patch
                    'Make'= $hostHardware.Manufacturer
                    'Model' = $hostHardware.Model
                    'S/N' = $hardwarePlatfrom.serialNumber
                    'BIOS' = $hostHardware.BiosVersion
                    'BIOS Release Date' = (($vmhost.ExtensionData.Hardware.BiosInfo.ReleaseDate -split " ")[0])
                    'CPU Model' = $hostHardware.CpuModel
                    'CPU Count' = $hostHardware.CpuCount
                    'CPU Core Total' = $hostHardware.CpuCoreCountTotal
                    'Speed (MHz)' = $hostHardware.MhzPerCpu
                    'Memory (GB)' = $vmhost.MemoryTotalGB -as [int]
                    'Memory Slots Count' = $hostHardware.MemorySlotCount
                    'Memory Slots Used' = $hostHardware.MemoryModules.Count
                    'Power Supplies' = $hostHardware.PowerSupplies.Count
                    'NIC Count' = $hostHardware.NicCount
                } #END [PSCustomObject]
            } #END if
    
            <#
                Get Host Configuration details
            #>
            if ($Configuration) {
                Write-Host "`tGathering configuration details from $vmhost ..."
                $vibList = @()
                $vmhostID = $vmhostView.Config.Host.Value
                $vmhostLM = $licenseManagerAssign.QueryAssignedLicenses($vmhostID)
                $vmhostSoftware = $esxcli2.software.vib.list.Invoke()
                $vmhostPatch = $vmhostSoftware | Sort-Object InstallDate -Descending | Select-Object -First 1
                $vmhostvDC = $vmhost | Get-Datacenter | Select-Object -ExpandProperty Name
                $vmhostCluster = $vmhost | Get-Cluster | Select-Object -ExpandProperty Name
    
                <#
                    Get List of software/patches last installed
                #>
                foreach($vibName in $vmhostSoftware | Where-Object {$_.InstallDate -eq $vmhostPatch.InstallDate}) {
                    $vibList += $vibName.Name
                }
    
                <#
                    Get NTP Configuraiton
                #>
                $ntpServerList = $vmhost | Get-VMHostNtpServer
                $ntpService = $vmhost | Get-VMHostService | Where-Object {$_.key -eq "ntpd"}
                $vmhostFireWall = $vmhost | Get-VMHostFirewallException
                $ntpFWException = $vmhostFireWall | Select-Object Name, Enabled | Where-Object {$_.Name -eq "NTP Client"}
    
                <#
                    Get syslog Configuration
                #>
                $syslogList = @()
                $syslogFWException =  $vmhostFireWall | Select-Object Name, Enabled | Where-Object {$_.Name -eq "syslog"}
                foreach($syslog in  $vmhost | Get-VMHostSysLogServer) {
                    $syslogList += $syslog.Host + ":" + $syslog.Port
                }
    
                <#
                    Use a custom object to store
                    collected data
                #>
                $configurationCollection += [PSCustomObject]@{
                    'Hostname' = $vmhost
                    'Make'= $hostHardware.Manufacturer
                    'Model' = $hostHardware.Model
                    'CPU Model' = $hostHardware.CpuModel
                    'Hyper-Threading' = $vmhost.HyperthreadingActive
                    'Max EVC Mode' = $vmhost.MaxEVCMode
                    'Product' = $vmhostView.Config.Product.Name
                    'Version' = $vmhostView.Config.Product.Version
                    'Build' = $vmhost.Build
                    'Update' = $esxiVersion.Update
                    'Patch' = $esxiVersion.Patch
                    'License Version' = $vmhostLM.AssignedLicense.Name | Select-Object -Unique
                    'License Key' = $vmhostLM.AssignedLicense.LicenseKey | Select-Object -Unique
                    'Connection State' = $vmhost.ConnectionState
                    'Standalone' = $vmhost.IsStandalone
                    'Cluster' = $vmhostCluster
                    'Virtual Datacenter' = $vmhostvDC
                    'vCenter' = $vmhost.ExtensionData.CLient.ServiceUrl.Split('/')[2]
                    'Software/Patch Last Installed' = $vmhostPatch.InstallDate
                    'Software/Patch Name(s)' = (@($vibList | Sort) -join ',')
                    'Service' = $ntpService.Label
                    'Service Running' = $ntpService.Running
                    'Startup Policy' = $ntpService.Policy
                    'NTP Client Enabled' = $ntpFWException.Enabled
                    'NTP Server' = (@($ntpServerList) -join ',')
                    'Syslog Server' = (@($syslogList) -join ',')
                    'Syslog Client Enabled' = $syslogFWException.Enabled
                } #END [PSCustomObject]
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
        if ($iSCSICollection -or $FibreChannelCollection -or $DatastoresCollection) {
            Write-Verbose ((Get-Date -Format G) + "`tInformation gathered")
        } else {
            Write-Verbose ((Get-Date -Format G) + "`tNo information gathered")
        } #END if/else
    
        <#
            Output to screen
            Export data to CSV, Excel
        #>
        if ($hardwareCollection) {
            Write-Host "`n" "ESXi Hardware Inventory:" -ForegroundColor Green
    
            if ($ExportCSV) {
                $hardwareCollection | Export-Csv ($outputFile + "Hardware.csv") -NoTypeInformation
                Write-Host "`tData exported to" ($outputFile + "Hardware.csv") "file" -ForegroundColor Green
            } elseif ($ExportExcel) {
                $hardwareCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Hardware_Inventory -NoNumberConversion * -AutoSize -BoldTopRow
                Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
            } elseif ($PassThru) {
                $ReturnCollection += $hardwareCollection 
                $ReturnCollection 
            } else {
                $hardwareCollection | Format-List
            }#END if
        } #END if
    
        if ($configurationCollection) {
            Write-Host "`n" "ESXi Host Configuration:" -ForegroundColor Green
    
            if ($ExportCSV) {
                $configurationCollection | Export-Csv ($outputFile + "Configuration.csv") -NoTypeInformation
                Write-Host "`tData exported to" ($outputFile + "Configuration.csv") "file" -ForegroundColor Green
            } elseif ($ExportExcel) {
                $configurationCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Host_Configuration -NoNumberConversion * -AutoSize -BoldTopRow
                Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
            } elseif ($PassThru) {
                $ReturnCollection += $configurationCollection
                $ReturnCollection  
            } else {
                $configurationCollection | Format-List
            }#END if
        } #END if
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
      https://github.com/arielsanchezmora/vDocumentation
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
        Get-ESXInventory -datacenter "all vdc" will gather all hosts in vCenter(s). This is the default if no Parameter (-esxi, -cluster, or -datacenter) is specified. 
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
    .PARAMETER PassThru
       Returns the object to console
    .EXAMPLE
        Get-ESXIODevice -esxi devvm001.lab.local -PassThru
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
            [switch]$PassThru,
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
        Write-Verbose ((Get-Date -Format G) + "`tValidate connection to to a vSphere server")
    
        if ($Global:DefaultViServers.Count -gt 0) {
            
            Write-Host "`tConnected to $Global:DefaultViServers" -ForegroundColor Green
        } else {
            Write-Error "You must be connected to a vSphere server before running this cmdlet."
            break
        } #END if/else
    
        <#
            Validate if a parameter was specified (-esxi, -cluster, or -datacenter)
            Although all 3 can be specified, only the first is used
            Example: -esxi "host001" -cluster "test-cluster". -esxi is the first parameter
            and what will be used.
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate parameters used")
    
        if ([string]::IsNullOrWhiteSpace($esxi) -and [string]::IsNullOrWhiteSpace($cluster) -and [string]::IsNullOrWhiteSpace($datacenter)) {
            Write-Verbose ((Get-Date -Format G) + "`tA parameter (-esxi, -cluster, -datacenter) was not specified. Will gather all hosts")
            $datacenter = "all vdc"
        } #END if
    
        <#
            Gather host list based on parameter used
        #>
        Write-Verbose ((Get-Date -Format G) + "`tGather host list")
    
        if ([string]::IsNullOrWhiteSpace($esxi)) {
            Write-Verbose ((Get-Date -Format G) + "`t-esxi parameter is Null or Empty")
    
            if ([string]::IsNullOrWhiteSpace($cluster)) {
                Write-Verbose ((Get-Date -Format G) + "`t-cluster parameter is Null or Empty")
    
                if ([string]::IsNullOrWhiteSpace($datacenter)) {
                    Write-Verbose ((Get-Date -Format G) + "`t-datacenter parameter is Null or Empty")
    
                } else {
                    Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using datacenter parameter")
    
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
                Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using cluster parameter")
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
            Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using esxi parameter")
            Write-Host "`tGathering host list..."
    
            foreach($invidualHost in $esxi) {
                $vHostList += $invidualHost.Trim() | Sort-Object -Property Name
            } #END foreach
        } #END if/else
    
        <#
            Validate export switches,
            folder path and dependencies
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate export switches and folder path")
    
        if ($ExportCSV -or $ExportExcel) {
            $currentLocation = (Get-Location).Path
    
            if ([string]::IsNullOrWhiteSpace($folderPath)) {
                Write-Verbose ((Get-Date -Format G) + "`t-folderPath parameter is Null or Empty")
                Write-Warning "`tFolder Path (-folderPath) was not specified for saving exported data. The current location: '$currentLocation' will be used"
                $outputFile = $currentLocation + "\" + $outputFile
            } else {
    
                if (Test-Path $folderPath) {
                    Write-Verbose ((Get-Date -Format G) + "`t'$folderPath' path found")
                    $lastCharsOfFolderPath = $folderPath.Substring($folderPath.Length - 1)
    
                    if ($lastCharsOfFolderPath -eq "\" -or $lastCharsOfFolderPath -eq "/") {
                        $outputFile = $folderPath + $outputFile
                    } else {
                        $outputFile = $folderPath + "\" + $outputFile
                    } #END if/else
                    Write-Verbose ((Get-Date -Format G) + "`t$outputFile")
                } else {
                    Write-Warning "`t'$folderPath' path not found. The current location: '$currentLocation' will be used instead"
                    $outputFile = $currentLocation + "\" + $outputFile
                } #END if/else
            } #END if/else
        } #END if
    
        if ($ExportExcel) {
    
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Write-Verbose ((Get-Date -Format G) + "`tImportExcel Module available")
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
            Write-Verbose ((Get-Date -Format G) + "`t$vmhost Connection State: " + $vmhost.ConnectionState)
    
            if ($vmhost.ConnectionState -eq "Connected" -or $vmhost.ConnectionState -eq "Maintenance") {
                <#
                    Do nothing - ESXi host is reachable
                #>
            } else {
                <#
                    Use a custom object and array to keep track of skipped
                    hosts and continue to the next foreach loop
                #>
                $skipCollection += [pscustomobject]@{
                    'Hostname' = $esxihost
                    'Connection State' = $esxihost.ConnectionState
                } #END [PSCustomObject]
                continue
            } #END if/else
    
            $esxcli2 = Get-EsxCli -VMHost $esxihost -V2
    
            <#
                Get IO Device info
            #>
            Write-Host "`tGathering information from $vmhost ..."
            $pciDevices = $esxcli2.hardware.pci.list.Invoke() | Where-Object {$_.VMKernelName -like "vmhba*" -or $_.VMKernelName -like "vmnic*" -or $_.VMKernelName -like "vmgfx*" } | Sort-Object -Property VMKernelName
         
            foreach ($pciDevice in $pciDevices) {
                $device = $vmhost | Get-VMHostPciDevice | Where-Object { $pciDevice.Address -match $_.Id }
            
                Write-Verbose ((Get-Date -Format G) + "`tGet driver version for: " + $pciDevice.ModuleName)
                $driverVersion = $esxcli2.system.module.get.Invoke(@{module = $pciDevice.ModuleName}) | Select-Object -ExpandProperty Version
    
                <#
                    Get NIC Firmware version
                #>
                if ($pciDevice.VMKernelName -like 'vmnic*') {
                    Write-Verbose ((Get-Date -Format G) + "`tGet Firmware version for: " + $pciDevice.VMKernelName)
                    $vmnicDetail = $esxcli2.network.nic.get.Invoke(@{nicname = $pciDevice.VMKernelName})
                    $firmwareVersion = $vmnicDetail.DriverInfo.FirmwareVersion
                
                    <#
                        Get NIC driver VIB package version
                    #>
                    Write-Verbose ((Get-Date -Format G) + "`tGet VIB details for: " + $pciDevice.ModuleName)
                    $driverVib = $esxcli2.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object {$_.Name -eq $vmnicDetail.DriverInfo.Driver -or $_.Name -eq "net-"+$vmnicDetail.DriverInfo.Driver -or $_.Name -eq "net55-"+$vmnicDetail.DriverInfo.Driver}
                    $vibName = $driverVib.Name
                    $vibVersion = $driverVib.Version
    
                <#
                    If HP Smart Array vmhba* (scsi-hpsa driver) then can get Firmware version
                    elese skip if VMkernnel is vmhba* 
                    Can't get HBA Firmware from Powercli at the moment
                    only through SSH or using Putty Plink+PowerCli.
                #>
                } elseif ($pciDevice.VMKernelName -like 'vmhba*') {
                    
                    if ($pciDevice.DeviceName -match "smart array") {
                        Write-Verbose ((Get-Date -Format G) + "`tGet Firmware version for: " + $pciDevice.VMKernelName)
                        $hpsa = $vmhost.ExtensionData.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo | Where-Object {$_.Name -match "HP Smart Array"}
                        $firmwareVersion = (($hpsa.Name -split "firmware")[1]).Trim()
                    } else {
                        Write-Verbose ((Get-Date -Format G) + "`tSkip Firmware version check for: " + $pciDevice.DeviceName)
                        $firmwareVersion = $null    
                    } #END if/ese
                        
                    <#
                        Get HBA driver VIB package version
                    #>
                    Write-Verbose ((Get-Date -Format G) + "`tGet VIB deatils for: " + $pciDevice.ModuleName)
                    $vibName = $pciDevice.ModuleName -replace "_", "-"
                    $driverVib = $esxcli2.software.vib.list.Invoke() | Select-Object -Property Name,Version | Where-Object {$_.Name -eq "scsi-"+$VibName -or $_.Name -eq "sata-"+$VibName -or $_.Name -eq $VibName}
                    $vibName = $driverVib.Name
                    $vibVersion = $driverVib.Version
                } else {
                    Write-Verbose ((Get-Date -Format G) + "`tSkipping: " + $pciDevice.DeviceName)
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
    
            if ($ExportCSV) {
                $outputCollection | Export-Csv ($outputFile + ".csv") -NoTypeInformation
                Write-Host "`tData exported to" ($outputFile + ".csv") "file" -ForegroundColor Green
            } elseif ($ExportExcel) {
                $outputCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname IO_Device -NoNumberConversion * -AutoSize -BoldTopRow
                Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
            } elseif ($PassThru) {
                $outputCollection
            } else {
                $outputCollection | Format-List
            }#END if
        } else {
            Write-Verbose ((Get-Date -Format G) + "`tNo information gathered")
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
      https://github.com/arielsanchezmora/vDocumentation
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
        Get-ESXInventory -datacenter "all vdc" will gather all hosts in vCenter(s). This is the default if no Parameter (-esxi, -cluster, or -datacenter) is specified. 
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
        Switch to get Physical Adapter details including uplinks to vswitch and CDP/LLDP Information
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
    .PARAMETER PassThru
       Returns the object to console
    .EXAMPLE
        Get-ESXNetworking -esxi devvm001.lab.local -PassThru
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
            [switch]$PassThru,
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
        Write-Verbose ((Get-Date -Format G) + "`tValidate connection to to a vSphere server")
    
        if ($Global:DefaultViServers.Count -gt 0) {
            
            Write-Host "`tConnected to $Global:DefaultViServers" -ForegroundColor Green
        } else {
            Write-Error "You must be connected to a vSphere server before running this cmdlet."
            break
        } #END if/else
    
        <#
            Validate if a parameter was specified (-esxi, -cluster, or -datacenter)
            Although all 3 can be specified, only the first is used
            Example: -esxi "host001" -cluster "test-cluster". -esxi is the first parameter
            and what will be used.
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate parameters used")
    
        if ([string]::IsNullOrWhiteSpace($esxi) -and [string]::IsNullOrWhiteSpace($cluster) -and [string]::IsNullOrWhiteSpace($datacenter)) {
            Write-Verbose ((Get-Date -Format G) + "`tA parameter (-esxi, -cluster, -datacenter) was not specified. Will gather all hosts")
            $datacenter = "all vdc"
        } #END if
    
        <#
            Gather host list based on parameter used
        #>
        Write-Verbose ((Get-Date -Format G) + "`tGather host list")
    
        if ([string]::IsNullOrWhiteSpace($esxi)) {
            Write-Verbose ((Get-Date -Format G) + "`t-esxi parameter is Null or Empty")
    
            if ([string]::IsNullOrWhiteSpace($cluster)) {
                Write-Verbose ((Get-Date -Format G) + "`t-cluster parameter is Null or Empty")
    
                if ([string]::IsNullOrWhiteSpace($datacenter)) {
                    Write-Verbose ((Get-Date -Format G) + "`t-datacenter parameter is Null or Empty")
    
                } else {
                    Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using datacenter parameter")
    
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
                Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using cluster parameter")
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
            Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using esxi parameter")
            Write-Host "`tGathering host list..."
    
            foreach($invidualHost in $esxi) {
                $vHostList += $invidualHost.Trim() | Sort-Object -Property Name
            } #END foreach
        } #END if/else
    
        <#
            Validate export switches,
            folder path and dependencies
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate export switches and folder path")
    
        if ($ExportCSV -or $ExportExcel) {
            $currentLocation = (Get-Location).Path
    
            if ([string]::IsNullOrWhiteSpace($folderPath)) {
                Write-Verbose ((Get-Date -Format G) + "`t-folderPath parameter is Null or Empty")
                Write-Warning "`tFolder Path (-folderPath) was not specified for saving exported data. The current location: '$currentLocation' will be used"
                $outputFile = $currentLocation + "\" + $outputFile
            } else {
    
                if (Test-Path $folderPath) {
                    Write-Verbose ((Get-Date -Format G) + "`t'$folderPath' path found")
                    $lastCharsOfFolderPath = $folderPath.Substring($folderPath.Length - 1)
    
                    if ($lastCharsOfFolderPath -eq "\" -or $lastCharsOfFolderPath -eq "/") {
                        $outputFile = $folderPath + $outputFile
                    } else {
                        $outputFile = $folderPath + "\" + $outputFile
                    } #END if/else
                    Write-Verbose ((Get-Date -Format G) + "`t$outputFile")
                } else {
                    Write-Warning "`t'$folderPath' path not found. The current location: '$currentLocation' will be used instead"
                    $outputFile = $currentLocation + "\" + $outputFile
                } #END if/else
            } #END if/else
        } #END if
    
        if ($ExportExcel) {
    
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Write-Verbose ((Get-Date -Format G) + "`tImportExcel Module available")
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
         <#
            Validate that a cmdlet switch was used. Options are
            -PhysicalAdapters, -VMkernelAdapters, -VirtualSwitches.
            By default all are executed unless one is specified. 
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate cmdlet switches")
    
        if ($PhysicalAdapters -or $VMkernelAdapters -or $VirtualSwitches) {
            Write-Verbose ((Get-Date -Format G) + "`tA cmdlet switch was specified")
        } else {
            Write-Verbose ((Get-Date -Format G) + "`tA cmdlet switch was not specified")
            Write-Verbose ((Get-Date -Format G) + "`tWill execute all (-PhysicalAdapters, -VMkernelAdapters, -VirtualSwitches)")
            $PhysicalAdapters = $true
            $VMkernelAdapters = $true
            $VirtualSwitches = $true
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
            Write-Verbose ((Get-Date -Format G) + "`t$vmhost Connection State: " + $vmhost.ConnectionState)
    
            if ($vmhost.ConnectionState -eq "Connected" -or $vmhost.ConnectionState -eq "Maintenance") {
                <#
                    Do nothing - ESXi host is reachable
                #>
            } else {
                <#
                    Use a custom object and array to keep track of skipped
                    hosts and continue to the next foreach loop
                #>
                $skipCollection += [pscustomobject]@{
                    'Hostname' = $esxihost
                    'Connection State' = $esxihost.ConnectionState
                } #END [PSCustomObject]
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
                    Write-Verbose ((Get-Date -Format G) + "`tGet device details for: " + $nic.Name)
                    $pciList = $esxcli2.hardware.pci.list.Invoke() | Where-Object {$_.VMKernelName -eq $nic.Name}
                    $nicList = $esxcli2.network.nic.list.Invoke() | Where-Object {$_.Name -eq $nic.Name}
    
                    <#
                        Get uplink vSwitch, check standard
                        vSwitch first then Distributed.
                    #>
                    if ($vSwitch = $esxcli2.network.vswitch.standard.list.Invoke() | Where-Object {$_.uplinks -contains $nic.Name}) {
                        Write-Verbose ((Get-Date -Format G) + "`tUplinks to vswitch: " + $vSwitch.Name)
                    } else {
                        $vSwitch = $esxcli2.network.vswitch.dvs.vmware.list.Invoke() | Where-Object {$_.uplinks -contains $nic.Name}
                        Write-Verbose ((Get-Date -Format G) + "`tUplinks to vswitch: " + $vSwitch.Name)
                    } #END if/else
    
                    <#
                        Get Device Discovery Protocol CDP/LLDP
                    #>
                    Write-Verbose ((Get-Date -Format G) + "`tGet Device Discovery Protocol for: " + $nic.Name)
                    $esxiHostView = $vmhost | Get-View 
                    $networkSystem = $esxiHostView.Configmanager.Networksystem
                    $networkView = Get-View $networkSystem
                    $networkViewInfo = $networkView.QueryNetworkHint($nic.Name)
    
                    If ($networkViewInfo.connectedswitchport -ne $null) {
                        Write-Verbose ((Get-Date -Format G) + "`tDevice Discovery Protocol: CDP")
                        $ddp = "CDP"
                        $ddpExtended = $networkViewInfo.connectedswitchport
                        $ddpDevID = $ddpExtended.DevId
                        $ddpDevIP = $ddpExtended.Address
                        $ddpDevPortId = $ddpExtended.PortId
                    } else {
                        Write-Verbose ((Get-Date -Format G) + "`tCDP not found")
    
                        if ($networkViewInfo.lldpinfo -ne $null) {
                            Write-Verbose ((Get-Date -Format G) + "`tDevice Discovery Protocol: LLDP")
                            $ddp = "LLDP"
                            $ddpDevID = $networkViewInfo.lldpinfo.Parameter | Where-Object {$_.Key -eq "System Name"} | Select-Object -ExpandProperty Value  
                            $ddpDevIP = $networkViewInfo.lldpinfo.Parameter | Where-Object {$_.Key -eq "Management Address"} | Select-Object -ExpandProperty Value  
                            $ddpDevPortId = $networkViewInfo.lldpinfo.Portid
                        } else {
                            Write-Verbose ((Get-Date -Format G) + "`tLLDP not found")
                            $ddp = $null
                            $ddpDevID = $null
                            $ddpDevIP = $null
                            $ddpDevPortId = $null
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
                    Write-Verbose ((Get-Date -Format G) + "`tGet device details for: " + $nic.Name)
    
                    <#
                        Get VMkernel adapter enabled services
                    #>
                    $enabledServices = @()
                    Write-Verbose ((Get-Date -Format G) + "`tGet Enabled Services for: " + $nic.Name)
    
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
                        and VLAN ID.
                    #>
                    Write-Verbose ((Get-Date -Format G) + "`tGet Port Group details for: " + $nic.PortGroupName)
                    $interfaceList = $esxcli2.network.ip.interface.list.Invoke() | Where-Object {$_.Name -eq $nic.Name}
    
                    <#
                        Get PortGroup Teaming Policy and vSwitch MTU using
                        Active adapter associated with the VMKernel Port.
                        Test against both Standard and Distributed Switch.
                    #>
                    Write-Verbose ((Get-Date -Format G) + "`tGet vSwitch MTU for: " + $interfaceList.Portset)
    
                    if($portGroup = Get-VirtualPortGroup -VMhost $vmhost -Name $nic.PortGroupName -Standard -ErrorAction SilentlyContinue | Where-Object {$_.VirtualSwitchName -eq $interfaceList.Portset}) {
                        Write-Verbose ((Get-Date -Format G) + "`tStandard vSwitch")
                        $portGroupTeam = $portGroup | Get-NicTeamingPolicy
                        $portVLanId = $portGroup | Select-Object -ExpandProperty VLanId
                        $vSwitch = $esxcli2.network.vswitch.standard.list.Invoke(@{vswitchname = $interfaceList.Portset})
                        $activeAdapters = (@($PortGroupTeam.ActiveNic) -join ',')
                        $standbyAdapters = (@($PortGroupTeam.StandbyNic) -join ',')
                        $unusedAdapters = (@($PortGroupTeam.UnusedNic) -join ',')
                    } else {
                        Write-Verbose ((Get-Date -Format G) + "`tDistributed vSwitch")
                        $portGroup = Get-VDPortgroup -Name $nic.PortGroupName -VDSwitch $interfaceList.Portset -ErrorAction SilentlyContinue
                        $portGroupTeam = $portGroup | Get-VDUplinkTeamingPolicy
                        $portVLanId = $portGroup.VlanConfiguration.VlanId
                        $vSwitch = $esxcli2.network.vswitch.dvs.vmware.list.Invoke(@{vdsname = $interfaceList.Portset})
                        $activeAdapters = (@($PortGroupTeam.ActiveUplinkPort) -join ',')
                        $standbyAdapters = (@($PortGroupTeam.StandbyUplinkPort) -join ',')
                        $unusedAdapters = (@($PortGroupTeam.UnusedUplinkPort) -join ',')
                    } #END if/else
    
                    <#
                        Get TCP/IP Stack details
                    #>
                    Write-Verbose ((Get-Date -Format G) + "`tGet VMkernel TCP/IP configuration...")
                    $tcpipConfig = $vmhost | Get-VMHostNetwork
    
                    if ($tcpipConfig.VirtualNic.Name -contains $nic.Name) {
                        $vmkGateway = $tcpipConfig.VMKernelGateway
                        $dnsAddress = $tcpipConfig.DnsAddress
                    } else {
                        $vmkGateway = $null
                        $dnsAddress = $null
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
                        'Active adapters' = $activeAdapters
                        'Standby adapters' = $standbyAdapters
                        'Unused adapters' = $unusedAdapters
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
                    Write-Verbose ((Get-Date -Format G) + "`tGet standard switch details for: " + $vSwitch.Name)
    
                    <#
                        Get PortGroup details,
                        Security Policy, and Teaming Policy
                    #>
                    if ($portGroups = $vSwitch | Get-VirtualPortGroup) {
    
                        foreach ($port in $portGroups) {
                        
                            Write-Verbose ((Get-Date -Format G) + "`tGet Port Group details for: " + $port.Name)
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
                    Write-Verbose ((Get-Date -Format G) + "`tGet distributed switch details for: " + $vSwitch.Name)
                    
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
    
                            Write-Verbose ((Get-Date -Format G) + "`tGet Port Group details for: " + $port.Name)
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
            Write-Verbose ((Get-Date -Format G) + "`tInformation gathered")
        } else {
            Write-Verbose ((Get-Date -Format G) + "`tNo information gathered")
        } #END if/else
    
        <#
            Output to screen
            Export data to CSV, Excel
        #>
        if ($PhysicalAdapterCollection) {
            Write-Host "`n" "ESXi Physical Adapters:" -ForegroundColor Green
    
            if ($ExportCSV) {
                $PhysicalAdapterCollection | Export-Csv ($outputFile + "PhysicalAdapters.csv") -NoTypeInformation
                Write-Host "`tData exported to" ($outputFile + "PhysicalAdapters.csv") "file" -ForegroundColor Green
            } elseif ($ExportExcel) {
                $PhysicalAdapterCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Physical_Adapters -NoNumberConversion * -AutoSize -BoldTopRow
                Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
            } elseif ($PassThru) {
                $PhysicalAdapterCollection   
            } else {
                $PhysicalAdapterCollection | Format-List
            }#END if
        } #END if
    
        if ($VMkernelAdapterCollection) {
            Write-Host "`n" "ESXi VMkernel Adapters:" -ForegroundColor Green
    
            if ($ExportCSV) {
                $VMkernelAdapterCollection | Export-Csv ($outputFile + "VMkernelAdapters.csv") -NoTypeInformation
                Write-Host "`tData exported to" ($outputFile + "VMkernelAdapters.csv") "file" -ForegroundColor Green
            } elseif ($ExportExcel) {
                $VMkernelAdapterCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname VMkernel_Adapters -NoNumberConversion * -AutoSize -BoldTopRow
                Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
            } elseif ($PassThru){
                $VMkernelAdapterCollection
            } else{
                $VMkernelAdapterCollection | Format-List
            } #END if
        } #END if
    
        if ($VirtualSwitchesCollection) {
            Write-Host "`n" "ESXi Virtual Switches:" -ForegroundColor Green
    
            if ($ExportCSV) {
                $VirtualSwitchesCollection | Export-Csv ($outputFile + "VirtualSwitches.csv") -Force -NoTypeInformation
                Write-Host "`tData exported to" ($outputFile + "VirtualSwitches.csv") "file" -ForegroundColor Green
            } elseif ($ExportExcel) {
                $VirtualSwitchesCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Virtual_Switches -NoNumberConversion * -AutoSize -BoldTopRow
                Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
            } elseif ($PassThru){
                $VirtualSwitchesCollection 
            } else {
                $VirtualSwitchesCollection | Format-List
            } #END if
        } #END if
    } #END function
    
    function Get-ESXStorage {
    <#
    .SYNOPSIS
        Get ESXi Storage Details
    .DESCRIPTION
        Will get iSCSI Software and Fibre Channel Adapter (HBA) details including Datastores
        All this can be gathered for a vSphere Cluster, Datacenter or individual ESXi host
    .NOTES
        Author     : Edgar Sanchez - @edmsanchez13
        Contributor: Ariel Sanchez - @arielsanchezmor
    .Link
      https://github.com/arielsanchezmora/vDocumentation
    .INPUTS
       No inputs required
    .OUTPUTS
       CSV file
       Excel file
    .PARAMETER esxi
       The name(s) of the vSphere ESXi Host(s)
    .EXAMPLE
        Get-ESXStorage -esxi devvm001.lab.local
    .PARAMETER cluster
       The name(s) of the vSphere Cluster(s)
    .EXAMPLE
        Get-ESXStorage -cluster production-cluster
    .PARAMETER datacenter
       The name(s) of the vSphere Virtual DataCenter(s)
    .EXAMPLE
        Get-ESXStorage -datacenter vDC001
        Get-ESXInventory -datacenter "all vdc" will gather all hosts in vCenter(s). This is the default if no Parameter (-esxi, -cluster, or -datacenter) is specified. 
    .PARAMETER ExportCSV
        Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
    .EXAMPLE
       Get-ESXStorage -cluster production-cluster -ExportCSV
    .PARAMETER ExportExcel
        Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
        ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
        File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
    .EXAMPLE
        Get-ESXStorage -cluster production-cluster -ExportExcel
    .PARAMETER StorageAdapters
        Default switch to get iSCSI Software and Fibre Channel Adapter (HBA) details
        This is default option that will get processed if no switch parameter is provided.
    .EXAMPLE
        Get-ESXStorage -cluster production-cluster -StorageAdapters
    .PARAMETER Datastores
        Switch to get Datastores details
    .EXAMPLE
        Get-ESXStorage -cluster production-cluster -Datastores
    .PARAMETER folderPath
        Specificies an alternate folder path of where the exported file should be saved.
    .EXAMPLE
        Get-ESXStorage -cluster production-cluster -ExportExcel -folderPath C:\temp
    .PARAMETER PassThru
        Returns the object to the console
    .EXAMPLE
        Get-ESXStorage -esxi devvm001.lab.local -PassThru
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
            [switch]$StorageAdapters,
            [switch]$Datastores,
            [switch]$PassThru,
            $folderPath
        )
    
        $FibreChannelCollection = @()
        $iSCSICollection = @()
        $DatastoresCollection = @()
        $skipCollection = @()
        $vHostList = @()
        $vDCList = @()
        $date = Get-Date -format s
        $date = $date -replace ":","-"
        $outputFile = "Storage" + $date
    
    <#
    ----------------------------------------------------------[Execution]----------------------------------------------------------
    #>
    
        <#
            Validate if any connected to a VIServer
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate connection to to a vSphere server")
    
        if ($Global:DefaultViServers.Count -gt 0) {
            
            Write-Host "`tConnected to $Global:DefaultViServers" -ForegroundColor Green
        } else {
            Write-Error "You must be connected to a vSphere server before running this cmdlet."
            break
        } #END if/else
    
        <#
            Validate if a parameter was specified (-esxi, -cluster, or -datacenter)
            Although all 3 can be specified, only the first is used
            Example: -esxi "host001" -cluster "test-cluster". -esxi is the first parameter
            and what will be used.
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate parameters used")
    
        if ([string]::IsNullOrWhiteSpace($esxi) -and [string]::IsNullOrWhiteSpace($cluster) -and [string]::IsNullOrWhiteSpace($datacenter)) {
            Write-Verbose ((Get-Date -Format G) + "`tA parameter (-esxi, -cluster, -datacenter) was not specified. Will gather all hosts")
            $datacenter = "all vdc"
        } #END if
    
        <#
            Gather host list based on parameter used
        #>
        Write-Verbose ((Get-Date -Format G) + "`tGather host list")
    
        if ([string]::IsNullOrWhiteSpace($esxi)) {
            Write-Verbose ((Get-Date -Format G) + "`t-esxi parameter is Null or Empty")
    
            if ([string]::IsNullOrWhiteSpace($cluster)) {
                Write-Verbose ((Get-Date -Format G) + "`t-cluster parameter is Null or Empty")
    
                if ([string]::IsNullOrWhiteSpace($datacenter)) {
                    Write-Verbose ((Get-Date -Format G) + "`t-datacenter parameter is Null or Empty")
    
                } else {
                    Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using datacenter parameter")
                    
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
                Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using cluster parameter")
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
            Write-Verbose ((Get-Date -Format G) + "`tExecuting cmdlet using esxi parameter")
            Write-Host "`tGathering host list..."
            
            foreach($invidualHost in $esxi) {
                $vHostList += $invidualHost.Trim() | Sort-Object -Property Name
            } #END foreach
        } #END if/else
    
        <#
            Validate export switches,
            folder path and dependencies
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate export switches and folder path")
    
        if ($ExportCSV -or $ExportExcel) {
            $currentLocation = (Get-Location).Path
    
            if ([string]::IsNullOrWhiteSpace($folderPath)) {
                Write-Verbose ((Get-Date -Format G) + "`t-folderPath parameter is Null or Empty")
                Write-Warning "`tFolder Path (-folderPath) was not specified for saving exported data. The current location: '$currentLocation' will be used"
                $outputFile = $currentLocation + "\" + $outputFile
            } else {
    
                if (Test-Path $folderPath) {
                    Write-Verbose ((Get-Date -Format G) + "`t'$folderPath' path found")
                    $lastCharsOfFolderPath = $folderPath.Substring($folderPath.Length - 1)
    
                    if ($lastCharsOfFolderPath -eq "\" -or $lastCharsOfFolderPath -eq "/") {
                        $outputFile = $folderPath + $outputFile
                    } else {
                        $outputFile = $folderPath + "\" + $outputFile
                    } #END if/else
                    Write-Verbose ((Get-Date -Format G) + "`t$outputFile")
                } else {
                    Write-Warning "`t'$folderPath' path not found. The current location: '$currentLocation' will be used instead"
                    $outputFile = $currentLocation + "\" + $outputFile
                } #END if/else
            } #END if/else
        } #END if
    
        if ($ExportExcel) {
    
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Write-Verbose ((Get-Date -Format G) + "`tImportExcel Module available")
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
            -Hardware, -Configuration. By default all are executed
            unless one is specified. 
        #>
        Write-Verbose ((Get-Date -Format G) + "`tValidate cmdlet switches")
    
        if ($StorageAdapters -or $Datastores) {
            Write-Verbose ((Get-Date -Format G) + "`tA cmdlet switch was specified")
        } else {
            Write-Verbose ((Get-Date -Format G) + "`tA cmdlet switch was not specified")
            Write-Verbose ((Get-Date -Format G) + "`tWill execute all (-StorageAdapters -Datastores)")
            $StorageAdapters = $true
            $Datastores = $true
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
            Write-Verbose ((Get-Date -Format G) + "`t$vmhost Connection State: " + $vmhost.ConnectionState)
    
            if ($vmhost.ConnectionState -eq "Connected" -or $vmhost.ConnectionState -eq "Maintenance") {
                <#
                    Do nothing - ESXi host is reachable
                #>
            } else {
                <#
                    Use a custom object and array to keep track of skipped
                    hosts and continue to the next foreach loop
                #>
                $skipCollection += [pscustomobject]@{
                    'Hostname' = $esxihost
                    'Connection State' = $esxihost.ConnectionState
                } #END [PSCustomObject]
                continue
            } #END if/else
    
            $esxcli2 = Get-EsxCli -VMHost $esxihost -V2
    
            <#
                Get Storage adapters (HBA) details
            #>
            if ($StorageAdapters) {
                Write-Host "`tGathering storage adapter details from $vmhost ..."
                
                <#
                    Get iSCSI Software HBA 
                #>
                Write-Verbose ((Get-Date -Format G) + "`tGet iSCSI Software Adapter...")
                if ($hba = $vmhost | Get-VMHostHba -Type iScsi | Where-Object {$_.Model -eq "iSCSI Software Adapter"}) {
                    
                    <#
                        Get iSCSI HBA Details
                    #>
                    Write-Verbose ((Get-Date -Format G) + "`tGet iSCSI HBA details for: " + $hba.Device)
                    $hbaBinding = $esxcli2.iscsi.networkportal.list.Invoke(@{adapter = $hba.Device})
                    $hbaTarget = Get-IScsiHbaTarget -IScsiHba $hba
                    $sendList = $hbaTarget | Where-Object {$_.Type -eq "Send"} | Select-Object -ExpandProperty Address
                    $staticList = $hbaTarget | Where-Object {$_.Type -eq "Static"} | Select-Object -ExpandProperty Address
    
                    <#
                        Get active physical adapters
                        based on PortGroup.Test both
                        standard and distributed switch
                    #>
                    foreach($vmkNic in $hbaBinding) {
                        Write-Verbose ((Get-Date -Format G) + "`tGet active physical adapter for: " + $vmkNic.PortGroup)
                        $vmNic = $vmhost | Get-VMHostNetworkAdapter | Where-Object {$_.Mac -eq $vmkNic.MACAddress}
                        $nicList = $esxcli2.network.nic.list.Invoke() | Where-Object {$_.Name -eq $vmNic.Name}
                        $portGroupTeam = Get-VirtualPortGroup -VMhost $vmhost -Name $vmkNic.PortGroup | Where-Object {$_.VirtualSwitchName -eq $vmkNic.Vswitch} | Get-NicTeamingPolicy
                        
                        if($portGroup = Get-VirtualPortGroup -VMhost $vmhost -Name $vmkNic.PortGroup -Standard -ErrorAction SilentlyContinue | Where-Object {$_.VirtualSwitchName -eq $vmkNic.Vswitch}) {
                            Write-Verbose ((Get-Date -Format G) + "`tStandard vSwitch")
                            $portGroupTeam = $portGroup | Get-NicTeamingPolicy
                            $vSwitch = $esxcli2.network.vswitch.standard.list.Invoke(@{vswitchname = $vmkNic.Vswitch})
                            $activeAdapters = (@($PortGroupTeam.ActiveNic) -join ',')
                        } else {
                            Write-Verbose ((Get-Date -Format G) + "`tDistributed vSwitch")
                            $portGroup = Get-VDPortgroup -Name $vmkNic.PortGroup -VDSwitch $vmkNic.Vswitch -ErrorAction SilentlyContinue
                            $portGroupTeam = $portGroup | Get-VDUplinkTeamingPolicy
                            $vSwitch = $esxcli2.network.vswitch.dvs.vmware.list.Invoke(@{vdsname = $vmkNic.Vswitch})
                            $activeAdapters = (@($PortGroupTeam.ActiveUplinkPort) -join ',')
                        } #END if/else
    
                        <#
                            Use a custom object to store
                            collected data
                        #>
                        $iSCSICollection += [PSCustomObject]@{
                            'Hostname' = $vmhost
                            'Device' = $hba.Device
                            'iSCSI Name' = $hba.IScsiName
                            'Model' = $hba.Model
                            'Send Targets' = (@($sendList) -join ',')
                            'Static Tarets' = (@($staticList) -join ',')
                            'Port Group' = $vmkNic.PortGroup + " (" + $vmkNic.Vswitch + ")"
                            'VMkernel Adapter' = $vmkNic.Vmknic
                            'Port Binding' = $vmkNic.CompliantStatus
                            'Path Status' = $vmkNic.PathStatus
                            'Physical Network Adapter' = $vmNic.Name + " (" + $vmnic.BitRatePerSec/1000 + " Gbit/s, " + $nicList.Duplex + ")"
                        } #END [PSCustomObject]
    
                    } #END foreach
                } #END if
    
                <#
                    Get Fibre Channel HBA
                #>
                Write-Verbose ((Get-Date -Format G) + "`tGet Fibre Channel Adapter...")
                if ($hba = $vmhost | Get-VMHostHba -Type FibreChannel) {
                    
                    <#
                        Get Fibre Channel HBA Details
                    #>
                    foreach($hbaDevice in $hba) {
                        Write-Verbose ((Get-Date -Format G) + "`tGet Fibre Channel HBA details for: " + $hbaDevice.Device)
                        $nodeWWN = ([String]::Format("{0:X}", $HbaDevice.NodeWorldWideName) -split "(\w{2})" | Where-Object {$_ -ne ""}) -join ":"
                        $portWWN = ([String]::Format("{0:X}", $HbaDevice.PortWorldWideName) -split "(\w{2})" | Where-Object {$_ -ne ""}) -join ":"
    
                         <#
                            Use a custom object to store
                            collected data
                        #>
                        $FibreChannelCollection += [PSCustomObject]@{
                            'Hostname' = $vmhost
                            'Device' = $hbaDevice.Device
                            'Model' = $hbaDevice.Model
                            'Node WWN' = $nodeWWN
                            'Port WWN' = $portWWN
                            'Driver' = $hbaDevice.Driver
                            'Speed (Gb)' = $hbaDevice.Speed
                            'Status' = $hbaDevice.Status
                        } #END [PSCustomObject]
                    } # END foreach
                } #END if
            } #END if
    
            <#
                Get Datastores
            #>
            if($Datastores) {
                
                <#
                    Get Datastore details
                #>
                Write-Host "`tGathering Datastore details from $vmhost ..."
                $hostDSList = $vmhost | Get-Datastore | Sort-Object -Property Name
    
                foreach($oneDS in $hostDSList) {
                    Write-Verbose ((Get-Date -Format G) + "`tGet Datastore details for: " + $oneDS.Name)
                    $dsCName = $oneDS.ExtensionData.Info.Vmfs.Extent | Select-Object -ExpandProperty DiskName
                    $dsDisk= $esxcli2.storage.nmp.device.list.Invoke(@{device = $dsCName})
    
                    <#
                        Validate against exising collection
                        to speed up foreach
                    #>
                    $itemInCollection = $DatastoresCollection | Where-Object {$_.'Canonical Name' -like $dsCName} | Select-Object -First 1
    
                    if($itemInCollection) {
                        Write-Verbose ((Get-Date -Format G) + "`t$dsCName already in collection; will reuse associated properties")
                        $dsName = $itemInCollection.'Datastore Name'
                        $dsDisplayName = $itemInCollection.'Device Name'
                        $LUN = $itemInCollection.LUN
                        $dsType = $itemInCollection.Type
                        $dsCluster = $itemInCollection.'Datastore Cluster'
                        $dsCapacityGB = $itemInCollection.'Capacity (GB)'
                        $ProvisionedGB = $itemInCollection.'Provisioned Space (GB)'
                        $dsFreeGB = $itemInCollection.'Free Space (GB)'
                        $dsTransportType = $itemInCollection.Transport
                        $dsMountPoint = $itemInCollection.'Mount Point'
                        $dsFileSystem = $itemInCollection.'File System Version'
    
                    } else {
                        Write-Verbose ((Get-Date -Format G) + "`t$dsCName not yet in collection; will query associated properties")
                        $dsView = $oneDS | Get-View
                        $dsSummary = $dsView | Select-Object -ExpandProperty Summary
                        $provisionedGB = [math]::round(($dsSummary.Capacity - $dsSummary.FreeSpace + $dsSummary.Uncommitted) / 1GB,2)
                        $dspath = $esxcli2.storage.core.path.list.Invoke(@{device = $dsCName}) | Select-Object -First 1
                        $dsDisplayName = (($dspath.DeviceDisplayName -Split " [(]")[0])
                        $LUN = $dspath.LUN
                        $vmHBA = $dspath.Adapter
                        $dsTransport = $vmhost | Get-VMHostHba | Select-Object Device, Type | Where-Object {$_.Device -eq $vmHBA}
    
                        if($oneDS.ParentFolder -eq $null -and $oneDS.ParentFolderId -match "StoragePod") {
                            $dsCluster = Get-DatastoreCluster -Id $oneDS.ParentFolderId | Select-Object -ExpandProperty Name
                            Write-Verbose ((Get-Date -Format G) + "`tDatastore is part of Datastore Cluster: " + $dsCluster)
                        } else {
                            $dsCluster = $null
                            Write-Verbose ((Get-Date -Format G) + "`tDatastore not part of a Datastore Cluster")
                        } #END if/else
    
                        $dsName = $oneDS.Name
                        $dsType = $oneDS.Type
                        $dsCapacityGB = $oneDS.CapacityGB
                        $dsFreeGB = [math]::round($oneDS.FreeSpaceGB,2)
                        $dsTransportType = $dsTransport.Type
                        $dsMountPoint = (($dsSummary.Url -split "ds://")[1])
                        $dsFileSystem = $oneDS.FileSystemVersion
                    } #END if/else
    
                    <#
                        Use a custom object to store
                        collected data
                    #>
                    $DatastoresCollection += [PSCustomObject]@{
                        'Hostname' = $vmhost
                        'Datastore Name' = $dsName
                        'Device Name' = $dsDisplayName
                        'Canonical Name' = $dsCName
                        'LUN' = $LUN
                        'Type' = $dsType
                        'Datastore Cluster' = $dsCluster
                        'Capacity (GB)' = $dsCapacityGB
                        'Provisioned Space (GB)' = $ProvisionedGB
                        'Free Space (GB)' = $dsFreeGB
                        'Transport' = $dsTransportType
                        'Mount Point' = $dsMountPoint
                        'Multipath Policy' = $dsDisk.PathSelectionPolicy
                        'File System Version' = $dsFileSystem
                    } #END [PSCustomObject]
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
        if ($iSCSICollection -or $FibreChannelCollection -or $DatastoresCollection) {
            Write-Verbose ((Get-Date -Format G) + "`tInformation gathered")
        } else {
            Write-Verbose ((Get-Date -Format G) + "`tNo information gathered")
        } #END if/else
    
        <#
            Output to screen
            Export data to CSV, Excel
        #>
        if ($iSCSICollection) {
            Write-Host "`n" "ESXi Storage iSCSI HBA:" -ForegroundColor Green
    
            if ($ExportCSV) {
                $iSCSICollection | Export-Csv ($outputFile + "iSCSI_HBA.csv") -NoTypeInformation
                Write-Host "`tData exported to" ($outputFile + "iSCSI_HBA.csv") "file" -ForegroundColor Green
            } elseif ($ExportExcel) {
                $iSCSICollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname iSCSI_HBA -NoNumberConversion * -AutoSize -BoldTopRow
                Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
            } elseif ($PassThru){
                $iSCSICollection
            } else{
                $iSCSICollection | Format-List
            } #END if
        } #END if
    
        if ($FibreChannelCollection) {
            Write-Host "`n" "ESXi FibreChannel HBA:" -ForegroundColor Green
    
            if ($ExportCSV) {
                $FibreChannelCollection | Export-Csv ($outputFile + "FibreChannelHBA.csv") -NoTypeInformation
                Write-Host "`tData exported to" ($outputFile + "FibreChannelHBA.csv") "file" -ForegroundColor Green
            } elseif ($ExportExcel) {
                $FibreChannelCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname FibreChannel_HBA -NoNumberConversion * -AutoSize -BoldTopRow
                Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
            } elseif ($PassThru) {
                $FibreChannelCollection
            } else {
                $FibreChannelCollection | Format-List
            } #END if
        } #END if
    
        if ($DatastoresCollection) {
            Write-Host "`n" "ESXi FibreChannel HBA:" -ForegroundColor Green
    
            if ($ExportCSV) {
                $DatastoresCollection | Export-Csv ($outputFile + "Datastores.csv") -NoTypeInformation
                Write-Host "`tData exported to" ($outputFile + "Datastores.csv") "file" -ForegroundColor Green
            } elseif ($ExportExcel) {
                $DatastoresCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Datastores -NoNumberConversion * -AutoSize -BoldTopRow
                Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
            } elseif ($PassThru) {
                $DatastoresCollection
            } else {
                $DatastoresCollection | Format-List
            }#END if
        } #END if
    } #END function
    