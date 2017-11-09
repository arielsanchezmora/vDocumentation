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
  
    <#
      Supressing "AvoidUsingConvertToSecureStringWithPlainText"
      in PSScriptAnalyzer code analysis. It's a one time token
      that is converted to secure string, not clear text string
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseOutputTypeCorrectly")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars")] # for $global:DefaultVIServers
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText")]
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
    $skipCollection = @()
    $vHostList = @()
    $ReturnCollection = @()
    $date = Get-Date -format s
    $date = $date -replace ":", "-"
    $outputFile = "Inventory" + $date
    
    <#
     ----------------------------------------------------------[Execution]----------------------------------------------------------
    #>
  
    <#
      Query PowerCLI and vDocumentation versions if
      running Verbose
    #>
    if ($VerbosePreference -eq "continue") {
        Write-Verbose -Message ((Get-Date -Format G) + "`tPowerCLI Version:")
        Get-Module -Name VMware.* | Select-Object -Property Name, Version | Format-Table -AutoSize
        Write-Verbose -Message ((Get-Date -Format G) + "`tvDocumentation Version:")
        Get-Module -Name vDocumentation | Select-Object -Property Name, Version | Format-Table -AutoSize
    } #END if

    <#
      Check for an active connection to a VIServer
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate connection to a vSphere server")
    if ($Global:DefaultViServers.Count -gt 0) {
        Write-Output -InputObject "`tConnected to $Global:DefaultViServers" -ForegroundColor Green
    }
    else {
        Write-Error -Message "You must be connected to a vSphere server before running this Cmdlet."
        break
    } #END if/else
    
    <#
      Validate if a parameter was specified (-esxi, -cluster, or -datacenter)
      Although all 3 can be specified, only the first one is used
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
                    Write-Output -InputObject "`tGathering all hosts from the following vCenter(s): " $Global:DefaultViServers
                    $vHostList = Get-VMHost | Sort-Object -Property Name                    
                }
                else {
                    Write-Output -InputObject "`tGathering host list from the following DataCenter(s): " (@($datacenter) -join ',')
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
            Write-Output -InputObject "`tGathering host list from the following Cluster(s): " (@($cluster) -join ',')
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
        Write-Output -InputObject "`tGathering host list..."
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
      -Hardware, -Configuration. By default all are executed
      unless one is specified. 
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate Cmdlet switches")
    if ($Hardware -or $Configuration) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tA Cmdlet switch was specified")
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tA Cmdlet switch was not specified")
        Write-Verbose -Message ((Get-Date -Format G) + "`tWill execute all (-Hardware -Configuration)")
        $Hardware = $true
        $Configuration = $true
    } #END if/else
    
    <#
      Initialize variables used for -Configuration switch
    #>
    if ($Configuration) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tInitializing -Configuration Cmdlet switch variables...")
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
            Write-Output -InputObject "`tGathering Hardware inventory from $vmhost ..."
            $mgmtIP = $vmhost | Get-VMHostNetworkAdapter | Where-Object {$_.ManagementTrafficEnabled -eq 'True'} | Select-Object -ExpandProperty IP
            $hardwarePlatfrom = $esxcli2.hardware.platform.get.Invoke()
    
            <#
              Get RAC IP, and Firmware
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering RAC IP...")
            $cimServicesTicket = $vmhostView.AcquireCimServicesTicket()
            $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $cimServicesTicket.SessionId, (ConvertTo-SecureString $cimServicesTicket.SessionId -AsPlainText -Force)
            $cimOpt = New-CimSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -Encoding Utf8 –UseSsl
            $session = New-CimSession -Authentication Basic -Credential $credential -ComputerName $vmhost -port 443 -SessionOption $cimOpt -ErrorAction SilentlyContinue
            $rac = $session | Get-CimInstance CIM_IPProtocolEndpoint -ErrorAction SilentlyContinue | Where-Object {$_.Name -match "Management Controller IP"}
            if ($rac.Name) {
                $racIP = $rac.IPv4Address
            }
            else { 
                $racIP = $null
            } #END if/ese
            if ($bmc = $vmhost.ExtensionData.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo | Where-Object {$_.Name -match "BMC Firmware"}) {
                $bmcFirmware = (($bmc.Name -split "firmware")[1]) -split " " | Select-Object -Last 1
            }
            else {
                $bmcFirmware = $null
            } #END if/else
    
            <#
              Use a custom object to store
              collected data
            #>
            $hardwareCollection += [PSCustomObject]@{
                'Hostname'           = $vmhost
                'Management IP'      = $mgmtIP
                'RAC IP'             = $racIP
                'RAC Firmware'       = $bmcFirmware
                'Product'            = $vmhostView.Config.Product.Name
                'Version'            = $vmhostView.Config.Product.Version
                'Build'              = $vmhost.Build
                'Update'             = $esxiVersion.Update
                'Patch'              = $esxiVersion.Patch
                'Make'               = $hostHardware.Manufacturer
                'Model'              = $hostHardware.Model
                'S/N'                = $hardwarePlatfrom.serialNumber
                'BIOS'               = $hostHardware.BiosVersion
                'BIOS Release Date'  = (($vmhost.ExtensionData.Hardware.BiosInfo.ReleaseDate -split " ")[0])
                'CPU Model'          = $hostHardware.CpuModel
                'CPU Count'          = $hostHardware.CpuCount
                'CPU Core Total'     = $hostHardware.CpuCoreCountTotal
                'Speed (MHz)'        = $hostHardware.MhzPerCpu
                'Memory (GB)'        = $vmhost.MemoryTotalGB -as [int]
                'Memory Slots Count' = $hostHardware.MemorySlotCount
                'Memory Slots Used'  = $hostHardware.MemoryModules.Count
                'Power Supplies'     = $hostHardware.PowerSupplies.Count
                'NIC Count'          = $hostHardware.NicCount
            } #END [PSCustomObject]
        } #END if
    
        <#
          Get ESXi configuration details
        #>
        if ($Configuration) {
            Write-Output -InputObject "`tGathering configuration details from $vmhost ..."

            <#
              Get ESXi licensing
              and software configuration
            #>
            $vmhostID = $vmhostView.Config.Host.Value
            $vmhostLM = $licenseManagerAssign.QueryAssignedLicenses($vmhostID)
            $vmhostPatch = $esxcli2.software.vib.list.Invoke() | Where-Object {$_.ID -match $vmhost.Build} | Select-Object -First 1
            $vmhostvDC = $vmhost | Get-Datacenter | Select-Object -ExpandProperty Name
            $vmhostCluster = $vmhost | Get-Cluster | Select-Object -ExpandProperty Name
            $imageProfile = $esxcli2.software.profile.get.Invoke()
                   
            <#
              Get services configuraiton
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering services configuration...")
            $vmServices = $vmhost | Get-VMHostService
            $vmhostFireWall = $vmhost | Get-VMHostFirewallException
            $ntpServerList = $vmhost | Get-VMHostNtpServer
            $ntpService = $vmhost | Get-VMHostService | Where-Object {$_.key -eq "ntpd"}
            $vmhostFireWall = $vmhost | Get-VMHostFirewallException
            $ntpFWException = $vmhostFireWall | Select-Object Name, Enabled | Where-Object {$_.Name -eq "NTP Client"}
            $sshService = $vmServices | Where-Object {$_.key -eq "TSM-SSH"}
            $sshServerFWException = $vmhostFireWall | Select-Object Name, Enabled | Where-Object {$_.Name -eq "SSH Server"}
            $esxiShellService = $vmServices | Where-Object {$_.key -eq "TSM"}
            $ShellTimeOut = (Get-AdvancedSetting -Entity $vmhost -Name "UserVars.ESXiShellTimeOut" -ErrorAction SilentlyContinue).value
            $interactiveShellTimeOut = (Get-AdvancedSetting -Entity $vmhost -Name "UserVars.ESXiShellInteractiveTimeOut" -ErrorAction SilentlyContinue).value
    
            <#
              Get syslog configuration
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering Syslog Configuration...")
            $syslogList = @()
            $syslogFWException = $vmhostFireWall | Select-Object Name, Enabled | Where-Object {$_.Name -eq "syslog"}
            foreach ($syslog in  $vmhost | Get-VMHostSysLogServer) {
                $syslogList += $syslog.Host + ":" + $syslog.Port
            } #END foreach
    
            <#
              Get UpTime and install date
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering UpTime Configuration...")
            $bootTimeUTC = $vmhost.ExtensionData.Runtime.BootTime
            $localTimeZone = Get-TimeZone
            $BootTime = [System.TimeZoneInfo]::ConvertTime($bootTimeUTC, $localTimeZone)
            $upTime = New-TimeSpan -Seconds $vmhost.ExtensionData.Summary.QuickStats.Uptime
            $upTimeDays = $upTime.Days
            $upTimeHours = $upTime.Hours
            $upTimeMinutes = $upTime.Minutes
            $vmUUID = $esxcli2.system.uuid.get.Invoke()
            $decimalDate = [Convert]::ToInt32($vmUUID.Split("-")[0], 16)
            $installDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime(([DateTime]'1/1/1970').AddSeconds($decimalDate))

            <#
              Get ESXi installation type
              and Boot config
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering ESXi installation type...")
            $bootDevice = $esxcli2.system.boot.device.get.Invoke()
            if ($bootDevice.BootFilesystemUUID) {
                if ($bootDevice.BootFilesystemUUID[6] -eq 'e') {
                    $installType = "Embedded"
                    $bootSource = $null
                }
                else {
                    $installType = "Installable"
                    $bootSource = $esxcli2.storage.filesystem.list.Invoke() | Where-Object {$_.UUID -eq $bootDevice.BootFilesystemUUID} | Select-Object -ExpandProperty MountPoint
                } #END if/else
            }
            else {
                if ($bootDevice.StatelessBootNIC) {
                    $installType = "PXE Stateless"
                    $bootSource = $bootDevice.StatelessBootNIC
                }
                else {
                    $installType = "PXE"
                    $bootSource = $bootDevice.BootNIC
                } #END if/else
            } #END if/else

            <#
              Use a custom object to store
              collected data
            #>
            $configurationCollection += [PSCustomObject]@{
                'Hostname'              = $vmhost
                'Make'                  = $hostHardware.Manufacturer
                'Model'                 = $hostHardware.Model
                'CPU Model'             = $hostHardware.CpuModel
                'Hyper-Threading'       = $vmhost.HyperthreadingActive
                'Max EVC Mode'          = $vmhost.MaxEVCMode
                'Product'               = $vmhostView.Config.Product.Name
                'Version'               = $vmhostView.Config.Product.Version
                'Build'                 = $vmhost.Build
                'Update'                = $esxiVersion.Update
                'Patch'                 = $esxiVersion.Patch
                'Install Type'          = $installType
                'Boot From'             = $bootSource
                'Image Profile'         = $imageProfile.Name
                'Acceptance Level'      = $imageProfile.AcceptanceLevel 
                'Boot Time'             = $BootTime
                'Uptime'                = "$upTimeDays Day(s), $upTimeHours Hour(s), $upTimeMinutes Minute(s)"
                'Install Date'          = $installDate
                'Last Patched'          = $vmhostPatch.InstallDate
                'License Version'       = $vmhostLM.AssignedLicense.Name | Select-Object -Unique
                'License Key'           = $vmhostLM.AssignedLicense.LicenseKey | Select-Object -Unique
                'Connection State'      = $vmhost.ConnectionState
                'Standalone'            = $vmhost.IsStandalone
                'Cluster'               = $vmhostCluster
                'Virtual Datacenter'    = $vmhostvDC
                'vCenter'               = $vmhost.ExtensionData.CLient.ServiceUrl.Split('/')[2]
                'NTP'               = $ntpService.Label
                'NTP Running'       = $ntpService.Running
                'NTP Startup Policy'        = $ntpService.Policy
                'NTP Client Enabled'    = $ntpFWException.Enabled
                'NTP Server'            = (@($ntpServerList) -join ',')
                'SSH'                       = $sshService.Label
                'SSH Running'               = $sshService.Running
                'SSH Startup Policy'        = $sshService.Policy
                'SSH TimeOut'               = $ShellTimeOut
                'SSH Server Enabled'        = $sshServerFWException.Enabled
                'ESXi Shell'                = $esxiShellService.Label
                'ESXi Shell Running'        = $esxiShellService.Running
                'ESXi Shell Startup Policy' = $esxiShellService.Policy
                'ESXi Shell TimeOut'        = $interactiveShellTimeOut
                'Syslog Server'         = (@($syslogList) -join ',')
                'Syslog Client Enabled' = $syslogFWException.Enabled
            } #END [PSCustomObject]
        } #END if
    } #END foreach
    
    <#
      Display skipped hosts and their connection status
    #>
    If ($skipCollection) {
        Write-Warning -Message "`tCheck Connection State or Host name"
        Write-Warning -Message "`tSkipped hosts:"
        $skipCollection | Format-Table -AutoSize
    } #END if
    
    <#
      Validate output arrays
    #>
    if ($hardwareCollection -or $configurationCollection) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tInformation gathered")
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tNo information gathered")
    } #END if/else
    
    <#
      Output to screen
      Export data to CSV, Excel
    #>
    if ($hardwareCollection) {
        Write-Output -InputObject "`n" "ESXi Hardware Inventory:" -ForegroundColor Green
        if ($ExportCSV) {
            $hardwareCollection | Export-Csv ($outputFile + "Hardware.csv") -NoTypeInformation
            Write-Output -InputObject "`tData exported to" ($outputFile + "Hardware.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $hardwareCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Hardware_Inventory -NoNumberConversion * -AutoSize -BoldTopRow
            Write-Output -InputObject "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $ReturnCollection += $hardwareCollection 
            $ReturnCollection 
        }
        else {
            $hardwareCollection | Format-List
        }#END if/else
    } #END if
    
    if ($configurationCollection) {
        Write-Output -InputObject "`n" "ESXi Host Configuration:" -ForegroundColor Green
        if ($ExportCSV) {
            $configurationCollection | Export-Csv ($outputFile + "Configuration.csv") -NoTypeInformation
            Write-Output -InputObject "`tData exported to" ($outputFile + "Configuration.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $configurationCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Host_Configuration -NoNumberConversion * -BoldTopRow
            Write-Output -InputObject "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $ReturnCollection += $configurationCollection
            $ReturnCollection  
        }
        else {
            $configurationCollection | Format-List
        }#END if/else
    } #END if
} #END function