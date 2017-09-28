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
    $date = $date -replace ":", "-"
    $outputFile = "Networking" + $date
    
    <#
     ----------------------------------------------------------[Execution]----------------------------------------------------------
    #>
    
    <#
      Query PowerCLI and vDocumentation versions if
      running Verbose
    #>
    if ($VerbosePreference -eq "continue") {
        Write-Verbose -Message ((Get-Date -Format G) + "`tPowercli Version:")
        Get-Module -Name VMware.* | Select-Object -Property Name, Version | Format-Table -AutoSize
        Write-Verbose -Message ((Get-Date -Format G) + "`tvDocumentation Version:")
        Get-Module -Name vDocumentation | Select-Object -Property Name, Version | Format-Table -AutoSize
    } #END if

    <#
      Check for an active connection to a VIServer
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate connection to a vSphere server")
    if ($Global:DefaultViServers.Count -gt 0) {
        Write-Host "`tConnected to $Global:DefaultViServers" -ForegroundColor Green
    }
    else {
        Write-Error -Message "You must be connected to a vSphere server before running this Cmdlet."
        break
    } #END if/else
    
    <#
      Validate if a parameter was specified (-esxi, -cluster, or -datacenter)
      Although all 3 can be specified, only the first is used
      Example: -esxi "host001" -cluster "test-cluster". -esxi is the first parameter
      and what will be used.
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate parameters used")
    if ([string]::IsNullOrWhiteSpace($esxi) -and [string]::IsNullOrWhiteSpace($cluster) -and [string]::IsNullOrWhiteSpace($datacenter)) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tA parameter (-esxi, -cluster, -datacenter) was not specified. Will gather all hosts")
        $datacenter = "all vdc"
    } #END if
    
    <#
      Gather host list based on parameter used
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tGather host list")
    if ([string]::IsNullOrWhiteSpace($esxi)) {
        Write-Verbose -Message ((Get-Date -Format G) + "`t-esxi parameter is Null or Empty")
        if ([string]::IsNullOrWhiteSpace($cluster)) {
            Write-Verbose -Message ((Get-Date -Format G) + "`t-cluster parameter is Null or Empty")
            if ([string]::IsNullOrWhiteSpace($datacenter)) {
                Write-Verbose -Message ((Get-Date -Format G) + "`t-datacenter parameter is Null or Empty")
            }
            else {
                Write-Verbose -Message ((Get-Date -Format G) + "`tExecuting Cmdlet using datacenter parameter")
                if ($datacenter -eq "all vdc") {
                    Write-Host "`tGathering all hosts from the following vCenter(s): " $Global:DefaultViServers
                    $vHostList = Get-VMHost | Sort-Object -Property Name
                }
                else {
                    Write-Host "`tGathering host list from the following DataCenter(s): " (@($datacenter) -join ',')
                    foreach ($vDCname in $datacenter) {
                        $tempList = Get-Datacenter -Name $vDCname.Trim() -ErrorAction SilentlyContinue | Get-VMHost 
                        if ([string]::IsNullOrWhiteSpace($tempList)) {
                            Write-Warning -Message "`tDatacenter with name $vDCname was not found in $Global:DefaultViServers"
                        }
                        else {
                            $vHostList += $tempList | Sort-Object -Property Name
                        } #END if/else
                    } #END foreach
                } #END if/else
            } #END if/else
        }
        else {
            Write-Verbose -Message ((Get-Date -Format G) + "`tExecuting Cmdlet using cluster parameter")
            Write-Host "`tGathering host list from the following Cluster(s): " (@($cluster) -join ',')
            foreach ($vClusterName in $cluster) {
                $tempList = Get-Cluster -Name $vClusterName.Trim() -ErrorAction SilentlyContinue | Get-VMHost 
                if ([string]::IsNullOrWhiteSpace($tempList)) {
                    Write-Warning -Message "`tCluster with name $vClusterName was not found in $Global:DefaultViServers"
                }
                else {
                    $vHostList += $tempList | Sort-Object -Property Name
                } #END if/else
            } #END foreach
        } #END if/else
    }
    else { 
        Write-Verbose -Message ((Get-Date -Format G) + "`tExecuting Cmdlet using esxi parameter")
        Write-Host "`tGathering host list..."
        foreach ($invidualHost in $esxi) {
            $vHostList += $invidualHost.Trim() | Sort-Object -Property Name
        } #END foreach
    } #END if/else
    
    <#
      Validate export switches,
      folder path and dependencies
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate export switches and folder path")
    if ($ExportCSV -or $ExportExcel) {
        $currentLocation = (Get-Location).Path
        if ([string]::IsNullOrWhiteSpace($folderPath)) {
            Write-Verbose -Message ((Get-Date -Format G) + "`t-folderPath parameter is Null or Empty")
            Write-Warning -Message "`tFolder Path (-folderPath) was not specified for saving exported data. The current location: '$currentLocation' will be used"
            $outputFile = $currentLocation + "\" + $outputFile
        }
        else {
            if (Test-Path $folderPath) {
                Write-Verbose -Message ((Get-Date -Format G) + "`t'$folderPath' path found")
                $lastCharsOfFolderPath = $folderPath.Substring($folderPath.Length - 1)
                if ($lastCharsOfFolderPath -eq "\" -or $lastCharsOfFolderPath -eq "/") {
                    $outputFile = $folderPath + $outputFile
                }
                else {
                    $outputFile = $folderPath + "\" + $outputFile
                } #END if/else
                Write-Verbose -Message ((Get-Date -Format G) + "`t$outputFile")
            }
            else {
                Write-Warning -Message "`t'$folderPath' path not found. The current location: '$currentLocation' will be used instead"
                $outputFile = $currentLocation + "\" + $outputFile
            } #END if/else
        } #END if/else
    } #END if
    
    if ($ExportExcel) {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Write-Verbose -Message ((Get-Date -Format G) + "`tImportExcel Module available")
        }
        else {
            Write-Warning -Message "`tImportExcel Module missing. Will export data to CSV file instead"
            Write-Warning -Message "`tImportExcel Module can be installed directly from the PowerShell Gallery"
            Write-Warning -Message "`tSee https://github.com/dfinke/ImportExcel for more information"
            $ExportExcel = $false
            $ExportCSV = $true
        } #END if/else
    } #END if
    
    <#
      Validate that a Cmdlet switch was used. Options are
      -PhysicalAdapters, -VMkernelAdapters, -VirtualSwitches.
      By default all are executed unless one is specified. 
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate Cmdlet switches")
    if ($PhysicalAdapters -or $VMkernelAdapters -or $VirtualSwitches) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tA Cmdlet switch was specified")
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tA Cmdlet switch was not specified")
        Write-Verbose -Message ((Get-Date -Format G) + "`tWill execute all (-PhysicalAdapters, -VMkernelAdapters, -VirtualSwitches)")
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
        Write-Verbose -Message ((Get-Date -Format G) + "`t$vmhost Connection State: " + $vmhost.ConnectionState)
        if ($vmhost.ConnectionState -eq "Connected" -or $vmhost.ConnectionState -eq "Maintenance") {
            <#
              Do nothing - ESXi host is reachable
            #>
        }
        else {
            <#
              Use a custom object to keep track of skipped
              hosts and continue to the next foreach loop
            #>
            $skipCollection += [pscustomobject]@{
                'Hostname'         = $esxihost
                'Connection State' = $esxihost.ConnectionState
            } #END [PSCustomObject]
            continue
        } #END if/else
        $esxcli2 = Get-EsxCli -VMHost $esxihost -V2
    
        <#
          Get physical adapter details
        #>
        if ($PhysicalAdapters) {
            Write-Host "`tGathering physical adapter details from $vmhost ..."
            $vmnics = $vmhost | Get-VMHostNetworkAdapter -Physical | Select-Object Name, Mac, Mtu
            foreach ($nic in $vmnics) {
                Write-Verbose -Message ((Get-Date -Format G) + "`tGet device details for: " + $nic.Name)
                $pciList = $esxcli2.hardware.pci.list.Invoke() | Where-Object {$_.VMKernelName -eq $nic.Name}
                $nicList = $esxcli2.network.nic.list.Invoke() | Where-Object {$_.Name -eq $nic.Name}
    
                <#
                  Get uplink vSwitch, check standard
                  vSwitch first then Distributed.
                #>
                if ($vSwitch = $esxcli2.network.vswitch.standard.list.Invoke() | Where-Object {$_.uplinks -contains $nic.Name}) {
                    Write-Verbose -Message ((Get-Date -Format G) + "`tUplinks to vswitch: " + $vSwitch.Name)
                }
                else {
                    $vSwitch = $esxcli2.network.vswitch.dvs.vmware.list.Invoke() | Where-Object {$_.uplinks -contains $nic.Name}
                    Write-Verbose -Message ((Get-Date -Format G) + "`tUplinks to vswitch: " + $vSwitch.Name)
                } #END if/else
    
                <#
                  Get Device Discovery Protocol CDP/LLDP
                #>
                Write-Verbose -Message ((Get-Date -Format G) + "`tGet Device Discovery Protocol for: " + $nic.Name)
                $esxiHostView = $vmhost | Get-View 
                $networkSystem = $esxiHostView.Configmanager.Networksystem
                $networkView = Get-View $networkSystem
                $networkViewInfo = $networkView.QueryNetworkHint($nic.Name)
                if ($networkViewInfo.connectedswitchport -ne $null) {
                    Write-Verbose -Message ((Get-Date -Format G) + "`tDevice Discovery Protocol: CDP")
                    $ddp = "CDP"
                    $ddpExtended = $networkViewInfo.connectedswitchport
                    $ddpDevID = $ddpExtended.DevId
                    $ddpDevIP = $ddpExtended.Address
                    $ddpDevPortId = $ddpExtended.PortId
                }
                else {
                    Write-Verbose -Message ((Get-Date -Format G) + "`tCDP not found")
                    if ($networkViewInfo.lldpinfo -ne $null) {
                        Write-Verbose -Message ((Get-Date -Format G) + "`tDevice Discovery Protocol: LLDP")
                        $ddp = "LLDP"
                        $ddpDevID = $networkViewInfo.lldpinfo.Parameter | Where-Object {$_.Key -eq "System Name"} | Select-Object -ExpandProperty Value  
                        $ddpDevIP = $networkViewInfo.lldpinfo.Parameter | Where-Object {$_.Key -eq "Management Address"} | Select-Object -ExpandProperty Value  
                        $ddpDevPortId = $networkViewInfo.lldpinfo.Portid
                    }
                    else {
                        Write-Verbose -Message ((Get-Date -Format G) + "`tLLDP not found")
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
                    'Hostname'           = $vmhost
                    'Name'               = $nic.Name
                    'Slot Description'   = $pciList.SlotDescription
                    'Device'             = $nicList.Description
                    'Duplex'             = $nicList.Duplex
                    'Link'               = $nicList.Link
                    'MAC'                = $nic.Mac
                    'MTU'                = $nicList.MTU
                    'Speed'              = $nicList.Speed
                    'vSwitch'            = $vSwitch.Name
                    'vSwitch MTU'        = $vSwitch.MTU
                    'Discovery Protocol' = $ddp
                    'Device ID'          = $ddpDevID
                    'Device IP'          = $ddpDevIP
                    'Port'               = $ddpDevPortId
                } #END [PSCustomObject]
            } #END foreach
        } #END if
    
        <#
          Get VMkernel adapter details
        #>
        if ($VMkernelAdapters) {
            Write-Host "`tGathering VMkernel adapter details from $vmhost ..."
            $vmnics = $vmhost | Get-VMHostNetworkAdapter -VMKernel
            foreach ($nic in $vmnics) {
                Write-Verbose -Message ((Get-Date -Format G) + "`tGathering details for: " + $nic.Name)
    
                <#
                  Get VMkernel adapter enabled services
                #>
                $enabledServices = @()
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
                  Get VMkernel adapter associated vSwitch, PortGroup Teaming Policy
                  and vSwitch MTU using Active adapter associated with the VMKernel Port.
                  Test against both Standard and Distributed Switch.
                #>
                $interfaceList = $esxcli2.network.ip.interface.list.Invoke() | Where-Object {$_.Name -eq $nic.Name}
                if ($interfaceList.VDSName -eq "N/A") {
                    Write-Verbose -Message ((Get-Date -Format G) + "`tStandard vSwitch: " + $interfaceList.Portset)
                    Write-Verbose -Message ((Get-Date -Format G) + "`tGet PortGroup details for: " + $nic.PortGroupName)
                    $portGroup = Get-VirtualPortGroup -VMhost $vmhost -Name $nic.PortGroupName -Standard -ErrorAction SilentlyContinue
                    $portGroupTeam = $portGroup | Get-NicTeamingPolicy
                    $portVLanId = $portGroup | Select-Object -ExpandProperty VLanId
                    $vSwitch = $esxcli2.network.vswitch.standard.list.Invoke(@{vswitchname = $interfaceList.Portset})
                    $vSwitchName = $interfaceList.Portset
                    $activeAdapters = (@($PortGroupTeam.ActiveNic) -join ',')
                    $standbyAdapters = (@($PortGroupTeam.StandbyNic) -join ',')
                    $unusedAdapters = (@($PortGroupTeam.UnusedNic) -join ',')
                }
                else {
                    if ($interfaceList.VDSUUID) {
                        Write-Verbose -Message ((Get-Date -Format G) + "`tDistributed vSwitch: " + $interfaceList.VDSName)
                        Write-Verbose -Message ((Get-Date -Format G) + "`tGet PortGroup details for: " + $nic.PortGroupName)
                        $portGroup = Get-VDPortgroup -Name $nic.PortGroupName -VDSwitch $interfaceList.Portset -ErrorAction SilentlyContinue
                        $portGroupTeam = $portGroup | Get-VDUplinkTeamingPolicy
                        $portVLanId = $portGroup.VlanConfiguration.VlanId
                        $vSwitch = $esxcli2.network.vswitch.dvs.vmware.list.Invoke(@{vdsname = $interfaceList.VDSName})
                        $vSwitchName = $interfaceList.VDSName
                        $activeAdapters = (@($PortGroupTeam.ActiveUplinkPort) -join ',')
                        $standbyAdapters = (@($PortGroupTeam.StandbyUplinkPort) -join ',')
                        $unusedAdapters = (@($PortGroupTeam.UnusedUplinkPort) -join ',')
                    }
                    else {
                        $vSwitch = $esxcli2.network.vswitch.dvs.vmware.list.Invoke() | Where-Object {$_.VDSID -eq $interfaceList.VDSName}
                        Write-Verbose -Message ((Get-Date -Format G) + "`t3rd Party Distributed vSwitch: " + $vSwitch.Name)
                        $portVLanId = $null
                        $vSwitchName = $vSwitch.Name
                        $activeAdapters = $null
                        $standbyAdapters = $null
                        $unusedAdapters = $null
                    } #END if/else
                } #END if/else
    
                <#
                  Get TCP/IP Stack details
                #>
                Write-Verbose -Message ((Get-Date -Format G) + "`tGet VMkernel TCP/IP configuration...")
                $tcpipConfig = $vmhost | Get-VMHostNetwork
                if ($tcpipConfig.VirtualNic.Name -contains $nic.Name) {
                    $vmkGateway = $tcpipConfig.VMKernelGateway
                    $dnsAddress = $tcpipConfig.DnsAddress
                }
                else {
                    $vmkGateway = $null
                    $dnsAddress = $null
                } #END if/else
                                        
                <#
                  Use a custom object to store
                  collected data
                #>
                $VMkernelAdapterCollection += [PSCustomObject]@{
                    'Hostname'         = $vmhost
                    'Name'             = $nic.Name
                    'MAC'              = $nic.Mac
                    'MTU'              = $nic.MTU
                    'IP'               = $nic.IP
                    'Subnet Mask'      = $nic.SubnetMask
                    'TCP/IP Stack'     = $interfaceList.NetstackInstance
                    'Default Gateway'  = $vmkGateway
                    'DNS'              = (@($dnsAddress) -join ',')
                    'PortGroup Name'   = $nic.PortGroupName
                    'VLAN ID'          = $portVLanId
                    'Enabled Services' = (@($enabledServices) -join ',')
                    'vSwitch'          = $vSwitchName
                    'vSwitch MTU'      = $vSwitch.MTU
                    'Active adapters'  = $activeAdapters
                    'Standby adapters' = $standbyAdapters
                    'Unused adapters'  = $unusedAdapters
                } #END [PSCustomObject]
            } #END foreach
        } #END if
    
        <#
          Get virtual vSwitches details
        #>
        if ($VirtualSwitches) {
            Write-Host "`tGathering virtual vSwitches details from $vmhost ..."
    
            <#
              Get standard switch details
            #>
            $StdvSwitch = Get-VirtualSwitch -VMHost $vmhost -Standard
            foreach ($vSwitch in $StdvSwitch) {
                Write-Verbose -Message ((Get-Date -Format G) + "`tStandard vSwitch: " + $vSwitch.Name)
    
                <#
                  Get PortGroup details,
                  Security Policy, and Teaming Policy
                #>
                if ($portGroups = $vSwitch | Get-VirtualPortGroup) { 
                    foreach ($port in $portGroups) {
                        Write-Verbose -Message ((Get-Date -Format G) + "`tGet Port Group details for: " + $port.Name)
                        $portGroupTeam = $port | Get-NicTeamingPolicy
                        $portGroupSecurity = $port | Get-SecurityPolicy
                        $portGroupPolicy = @()
                        if ($portGroupSecurity.AllowPromiscuous) {
                            $portGroupPolicy += "Accept"
                        }
                        else {
                            $portGroupPolicy += "Reject"
                        } #END if/else
                        if ($portGroupSecurity.MacChanges) {
                            $portGroupPolicy += "Accept"
                        }
                        else {
                            $portGroupPolicy += "Reject"
                        } #END if/else
                        if ($portGroupSecurity.ForgedTransmits) {
                            $portGroupPolicy += "Accept"
                        }
                        else {
                            $portGroupPolicy += "Reject"
                        } #END if/else
    
                        <#
                          Use a custom object to store
                          collected data
                        #>
                        $VirtualSwitchesCollection += [PSCustomObject]@{
                            'Hostname'                                        = $vmhost
                            'Type'                                            = "Standard"
                            'Version'                                         = $null
                            'Name'                                            = $vSwitch.Name
                            'Uplink/ConnectedAdapters'                        = (@($vSwitch.Nic) -join ',')
                            'PortGroup'                                       = $port.Name
                            'VLAN ID'                                         = $port.VLanId
                            'Active adapters'                                 = (@($PortGroupTeam.ActiveNic) -join ',')
                            'Standby adapters'                                = (@($PortGroupTeam.StandbyNic) -join ',')
                            'Unused adapters'                                 = (@($PortGroupTeam.UnusedNic) -join ',')
                            'Security Promiscuous/MacChanges/ForgedTransmits' = (@($portGroupPolicy) -join '/')                        
                        } #END [PSCustomObject]
                    } #END foreach
                }
                else {
                    <#
                      Use a custom object to store
                      collected data
                    #>
                    $VirtualSwitchesCollection += [PSCustomObject]@{
                        'Hostname'                                        = $vmhost
                        'Type'                                            = "Standard"
                        'Version'                                         = $null
                        'Name'                                            = $vSwitch.Name
                        'Uplink/ConnectedAdapters'                        = (@($vSwitch.Nic) -join ',')
                        'PortGroup'                                       = $null
                        'VLAN ID'                                         = $null
                        'Active adapters'                                 = $null
                        'Standby adapters'                                = $null
                        'Unused adapters'                                 = $null
                        'Security Promiscuous/MacChanges/ForgedTransmits' = $null
                    } #END [PSCustomObject]
                } #END if/else
            } #END foreach
    
            <#
              Get distributed vSwitch details
            #>
            $dVSwitch = Get-VDSwitch -VMHost $vmhost
            foreach ($vSwitch in $dVSwitch) {
                if ($vSwitch.Vendor -match "VMware") {
                    Write-Verbose -Message ((Get-Date -Format G) + "`tDistributed vSwitch: " + $vSwitch.Name)
                    
                    <#
                      Get PortGroup details,
                      Security Policy, and Teaming Policy
                    #>
                    $portGroups = $vSwitch | Get-VDPortgroup
                    $vSwitchUplink = @()
                    if ($portGroups) {
                        <#
                          Get distributed switch Uplinks
                        #>
                        $vSwitchUplink = $vSwitch | Get-VDPort -Uplink
                        $uplinkConnected = @()
                        foreach ($uplink in $vSwitchUplink) {
                            $uplinkConnected += $uplink.Name + "," + $uplink.ConnectedEntity
                        } #END foreach
                        foreach ($port in $portGroups | Where-Object {!$_.IsUplink}) {
                            Write-Verbose -Message ((Get-Date -Format G) + "`tGet Port Group details for: " + $port.Name)
                            $portGroupTeam = $port | Get-VDUplinkTeamingPolicy
                            $portGroupSecurity = $port | Get-VDSecurityPolicy
                            $portGroupPolicy = @()
                            if ($portGroupSecurity.AllowPromiscuous) {
                                $portGroupPolicy += "Accept"
                            }
                            else {
                                $portGroupPolicy += "Reject"
                            } #END if/else
                            if ($portGroupSecurity.MacChanges) {
                                $portGroupPolicy += "Accept"
                            }
                            else {
                                $portGroupPolicy += "Reject"
                            } #END if/else
                            if ($portGroupSecurity.ForgedTransmits) {
                                $portGroupPolicy += "Accept"
                            }
                            else {
                                $portGroupPolicy += "Reject"
                            } #END if/else
    
                            <#
                              Use a custom object to store
                              collected data
                            #>
                            $VirtualSwitchesCollection += [PSCustomObject]@{
                                'Hostname'                                        = $vmhost
                                'Type'                                            = "Distributed"
                                'Version'                                         = $vSwitch.Version
                                'Name'                                            = $vSwitch.Name
                                'Uplink/ConnectedAdapters'                        = (@($uplinkConnected) -join '/')
                                'PortGroup'                                       = $port.Name
                                'VLAN ID'                                         = $port.VlanConfiguration.VlanId
                                'Active adapters'                                 = (@($PortGroupTeam.ActiveUplinkPort) -join ',')
                                'Standby adapters'                                = (@($PortGroupTeam.StandbyUplinkPort) -join ',')
                                'Unused adapters'                                 = (@($PortGroupTeam.UnusedUplinkPort) -join ',')
                                'Security Promiscuous/MacChanges/ForgedTransmits' = (@($portGroupPolicy) -join '/')
                            } #END [PSCustomObject]
                        } #END foreach
                    }
                    else {
                        <#
                          Use a custom object to store
                          collected data
                        #>
                        $VirtualSwitchesCollection += [PSCustomObject]@{
                            'Hostname'                                        = $vmhost
                            'Type'                                            = "Distributed"
                            'Version'                                         = $vSwitch.Version
                            'Name'                                            = $vSwitch.Name
                            'Uplink/ConnectedAdapters'                        = (@($uplinkConnected) -join '/')
                            'PortGroup'                                       = $port.Name
                            'VLAN ID'                                         = $port.VlanConfiguration.VlanId
                            'Active adapters'                                 = (@($PortGroupTeam.ActiveUplinkPort) -join ',')
                            'Standby adapters'                                = (@($PortGroupTeam.StandbyUplinkPort) -join ',')
                            'Unused adapters'                                 = (@($PortGroupTeam.UnusedUplinkPort) -join ',')
                            'Security Promiscuous/MacChanges/ForgedTransmits' = (@($portGroupPolicy) -join '/')
                        } #END [PSCustomObject]
                    } #END if/else
                }
                else {
                    Write-Verbose -Message ((Get-Date -Format G) + "`t3rd Party distributed vSwitch: " + $vSwitch.Name)
                    $ThirdPartyVirtualSwitchesCollection = @()
                    $vSwitchView = $vSwitch | Get-View
                    $vSwitchUplink = $esxcli2.network.vswitch.dvs.vmware.list.Invoke() | Where-Object {$_.Name -eq $vSwitch.Name}
                    $portGroups = $vSwitch | Get-VDPortgroup | Where-Object {$_.IsUplink -eq $false} | Select-Object -Property Name

                    <#
                      Use a custom object to store
                      collected data
                    #>
                    $ThirdPartyVirtualSwitchesCollection += [PSCustomObject]@{
                        'Hostname' = $vmhost
                        'Type'     = "Distributed"
                        'Version'  = $vSwitch.Version
                        'Build'    = $vSwitchView.Summary.ProductInfo.Build
                        'Name'     = $vSwitch.Name
                        'Vendor'   = $vSwitchView.Summary.ProductInfo.Vendor
                        'Model'    = $vSwitchView.Summary.ProductInfo.Name
                        'Bundle ID' = $vSwitchView.Summary.ProductInfo.BundleId
                        'Bundle URL' = $vSwitchView.Summary.ProductInfo.BundleUrl
                        'MTU'        = $vSwitchUplink.MTU
                        'Uplinks'    = (@($vSwitchUplink.uplinks) -join ',')
                        'PortGroups'  = (@($portGroups.Name) -join ',')
                    } #END [PSCustomObject]    
                    } #END if/else
            } #END foreach
        } #END if
    } #END foreach       
    
    <#
      Display skipped hosts and their connection status
    #>
    If ($skipCollection) {
        Write-Warning -Message "`tCheck Connection State or Host name "
        Write-Warning -Message "`tSkipped hosts: "
        $skipCollection | Format-Table -AutoSize
    } #END if
    
    <#
      Validate output arrays
    #>
    if ($PhysicalAdapterCollection -or $VMkernelAdapterCollection -or $VirtualSwitchesCollection -or $ThirdPartyVirtualSwitchesCollection) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tInformation gathered")
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tNo information gathered")
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
        }
        elseif ($ExportExcel) {
            $PhysicalAdapterCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Physical_Adapters -NoNumberConversion * -AutoSize -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $PhysicalAdapterCollection   
        }
        else {
            $PhysicalAdapterCollection | Format-List
        }#END if/else
    } #END if
    if ($VMkernelAdapterCollection) {
        Write-Host "`n" "ESXi VMkernel Adapters:" -ForegroundColor Green
        if ($ExportCSV) {
            $VMkernelAdapterCollection | Export-Csv ($outputFile + "VMkernelAdapters.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "VMkernelAdapters.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $VMkernelAdapterCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname VMkernel_Adapters -NoNumberConversion * -AutoSize -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $VMkernelAdapterCollection
        }
        else {
            $VMkernelAdapterCollection | Format-List
        } #END if/else
    } #END if
    if ($VirtualSwitchesCollection) {
        Write-Host "`n" "ESXi Virtual Switches:" -ForegroundColor Green
        if ($ExportCSV) {
            $VirtualSwitchesCollection | Export-Csv ($outputFile + "VirtualSwitches.csv") -Force -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "VirtualSwitches.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $VirtualSwitchesCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Virtual_Switches -NoNumberConversion * -AutoSize -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $VirtualSwitchesCollection 
        }
        else {
            $VirtualSwitchesCollection | Format-List
        } #END if/else
    } #END if
    if ($ThirdPartyVirtualSwitchesCollection) {
        Write-Host "`n" "ESXi 3rd party Virtual Switches:" -ForegroundColor Green
        if ($ExportCSV) {
            $ThirdPartyVirtualSwitchesCollection | Export-Csv ($outputFile + "3rdPartyvSwitches.csv") -Force -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "3rdPartyvSwitches.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $ThirdPartyVirtualSwitchesCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname 3rdParty_vSwitches -NoNumberConversion * -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $ThirdPartyVirtualSwitchesCollectionn 
        }
        else {
            $ThirdPartyVirtualSwitchesCollection | Format-List
        } #END if/else
    } #END if
} #END function