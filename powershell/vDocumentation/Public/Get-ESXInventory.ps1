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
       File Name    : Get-ESXInventory.ps1
       Author       : Edgar Sanchez - @edmsanchez13
       Contributor  : Ariel Sanchez - @arielsanchezmor
       Version      : 2.4.7
     .Link
       https://github.com/arielsanchezmora/vDocumentation
     .INPUTS
       No inputs required
     .OUTPUTS
       CSV file
       Excel file
     .PARAMETER VMhost
       The name(s) of the vSphere ESXi Host(s)
     .EXAMPLE
       Get-ESXInventory -VMhost devvm001.lab.local
     .PARAMETER Cluster
       The name(s) of the vSphere Cluster(s)
     .EXAMPLE
       Get-ESXInventory -Cluster production
     .PARAMETER Datacenter
       The name(s) of the vSphere Virtual Datacenter(s).
     .EXAMPLE
       Get-ESXInventory -Datacenter vDC001
     .PARAMETER ExportCSV
       Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-ESXInventory -Cluster production -ExportCSV
     .PARAMETER ExportExcel
       Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
       ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
       File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-ESXInventory -Cluster production -ExportExcel
     .PARAMETER Hardware
       Switch to get Hardware inventory
     .EXAMPLE
       Get-ESXInventory -Cluster production -Hardware
     .PARAMETER Configuration
       Switch to get system configuration details
     .EXAMPLE
       Get-ESXInventory -Cluster production -Configuration
     .PARAMETER folderPath
       Specify an alternate folder path where the exported data should be saved.
     .EXAMPLE
       Get-ESXInventory -Cluster production -ExportExcel -folderPath C:\temp
     .PARAMETER PassThru
       Switch to return object to command line
     .EXAMPLE
       Get-ESXInventory -VMhost 192.168.1.100 -Hardware -PassThru
    #> 
    
    <#
     ----------------------------------------------------------[Declarations]----------------------------------------------------------
    #>
  
    <#
      Supressing "AvoidUsingConvertToSecureStringWithPlainText"
      in PSScriptAnalyzer code analysis. It's a one time token
      that is converted to secure string, not clear text string
    #>
    [CmdletBinding(DefaultParameterSetName = 'VMhost')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("AvoidUsingConvertToSecureStringWithPlainText", "")]
    param (
        [Parameter(Mandatory = $false,
            ParameterSetName = "VMhost")]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String[]]$VMhost = "*",
        [Parameter(Mandatory = $false,
            ParameterSetName = "Cluster")]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String[]]$Cluster,
        [Parameter(Mandatory = $false,
            ParameterSetName = "DataCenter")]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String[]]$DataCenter,        
        [switch]$ExportCSV,
        [switch]$ExportExcel,
        [switch]$Hardware,
        [switch]$Configuration,
        [switch]$PassThru,
        $folderPath
    )
    
    $hardwareCollection = [System.Collections.ArrayList]@()
    $configurationCollection = [System.Collections.ArrayList]@()
    $skipCollection = @()
    $vHostList = @()
    $returnCollection = @()
    $date = Get-Date -format s
    $date = $date -replace ":", "-"
    $outputFile = "Inventory" + $date
    
    <#
     ----------------------------------------------------------[Execution]----------------------------------------------------------
    #>
  
    $stopWatch = [system.diagnostics.stopwatch]::startNew()
    if ($PSBoundParameters.ContainsKey('Cluster') -or $PSBoundParameters.ContainsKey('DataCenter')) {
        [String[]]$VMhost = $null
    } #END if

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
        Write-Host "`tConnected to $Global:DefaultViServers" -ForegroundColor Green
    }
    else {
        Write-Error -Message "You must be connected to a vSphere server before running this Cmdlet."
        break
    } #END if/else
        
    <#
      Gather host list based on Parameter set used
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tGather host list")
    if ($VMhost) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tExecuting Cmdlet using VMhost parameter set")
        Write-Output "`tGathering host list..."
        foreach ($invidualHost in $VMhost) {
            $tempList = Get-VMHost -Name $invidualHost.Trim() -ErrorAction SilentlyContinue
            if ($tempList) {
                $vHostList += $tempList
            }
            else {
                Write-Warning -Message "`tESXi host $invidualHost was not found in $Global:DefaultViServers"
            } #END if/else
        } #END foreach    
    } #END if
    if ($Cluster) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tExecuting Cmdlet using Cluster parameter set")
        Write-Output ("`tGathering host list from the following Cluster(s): " + (@($Cluster) -join ','))
        foreach ($vClusterName in $Cluster) {
            $tempList = Get-Cluster -Name $vClusterName.Trim() -ErrorAction SilentlyContinue | Get-VMHost 
            if ($tempList) {
                $vHostList += $tempList
            }
            else {
                Write-Warning -Message "`tCluster with name $vClusterName was not found in $Global:DefaultViServers"
            } #END if/else
        } #END foreach
    } #END if
    if ($DataCenter) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tExecuting Cmdlet using Datacenter parameter set")
        Write-Output ("`tGathering host list from the following DataCenter(s): " + (@($DataCenter) -join ','))
        foreach ($vDCname in $DataCenter) {
            $tempList = Get-Datacenter -Name $vDCname.Trim() -ErrorAction SilentlyContinue | Get-VMHost 
            if ($tempList) {
                $vHostList += $tempList
            }
            else {
                Write-Warning -Message "`tDatacenter with name $vDCname was not found in $Global:DefaultViServers"
            } #END if/else
        } #END foreach
    } #END if
    $tempList = $null
    
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
    } #END if
    
    <#
      Main code execution
    #>
    $vHostList = $vHostList | Sort-Object -Property Name
    foreach ($esxiHost in $vHostList) {
    
        <#
          Skip if ESXi host is not in a Connected
          or Maintenance ConnectionState
        #>
        Write-Verbose -Message ((Get-Date -Format G) + "`t$esxiHost Connection State: " + $esxiHost.ConnectionState)
        if ($esxiHost.ConnectionState -eq "Connected" -or $esxiHost.ConnectionState -eq "Maintenance") {
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
                'Hostname'         = $esxiHost.Name
                'Connection State' = $esxiHost.ConnectionState
            } #END [PSCustomObject]
            continue
        } #END if/else
        $esxcli = Get-EsxCli -VMHost $esxiHost -V2
        $hostHardware = $esxiHost | Get-VMHostHardware -WaitForAllData -SkipAllSslCertificateChecks -ErrorAction SilentlyContinue
        $vmhostView = $esxiHost | Get-View
        $esxiUpdateLevel = (Get-AdvancedSetting -Name "Misc.HostAgentUpdateLevel" -Entity $esxiHost -ErrorAction SilentlyContinue -ErrorVariable err).Value
        if ($esxiUpdateLevel) {
            $esxiVersion = ($esxiHost.Version) + " U" + $esxiUpdateLevel
        }
        else {
            $esxiVersion = $esxiHost.Version
            Write-Verbose -Message ((Get-Date -Format G) + "`tFailed to get ESXi Update Level, Error : " + $err)
        } #END if/else
                    
        <#
          Get Hardware invetory details
        #>
        if ($Hardware) {
            Write-Output "`tGathering Hardware inventory from $esxiHost ..."
            $mgmtIP = Get-VMHostNetworkAdapter -VMHost $esxiHost -VMKernel | Where-Object {$_.ManagementTrafficEnabled -eq 'True'} | Select-Object -ExpandProperty IP
            $hardwarePlatfrom = $esxcli.hardware.platform.get.Invoke()
    
            <#
              Get RAC IP, and Firmware
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering RAC IP...")
            $cimServicesTicket = $vmhostView.AcquireCimServicesTicket()
            $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $cimServicesTicket.SessionId, (ConvertTo-SecureString $cimServicesTicket.SessionId -AsPlainText -Force)
            try {
                $racIP = $null
                $racMAC = $null
                $cimOpt = New-CimSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -Encoding Utf8 –UseSsl
                $session = New-CimSession -Authentication Basic -Credential $credential -ComputerName $esxiHost -port 443 -SessionOption $cimOpt -ErrorAction SilentlyContinue -ErrorVariable err
                if ($err) {
                    Write-Verbose -Message ((Get-Date -Format G) + "`t$err")
                }
                $rac = $session | Get-CimInstance CIM_IPProtocolEndpoint -ErrorAction SilentlyContinue  -ErrorVariable err | Where-Object {$_.Name -match "Management Controller IP"}
                if ($rac.Name) {
                    $racIP = $rac.IPv4Address
                    $racMAC = $rac.MACAddress
                } #END if
            }
            catch {
                Write-Verbose -Message ((Get-Date -Format G) + "`tCIM session failed, error:")
                Write-Verbose -Message ((Get-Date -Format G) + "`t$err")
            } #END try/catch
            if ($bmc = $esxiHost.ExtensionData.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo | Where-Object {$_.Name -match "BMC Firmware"}) {
                $bmcFirmware = (($bmc.Name -split "firmware")[1]) -split " " | Select-Object -Last 1
            }
            else {
                Write-Verbose -Message ((Get-Date -Format G) + "`tFailed to get BMC firmware via CIM, testing using WSMan ...")
                try {
                    $bmcFirmware = $null
                    $cimOpt = New-WSManSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck -ErrorAction SilentlyContinue -ErrorVariable err 
                    $uri = "https`://" + $esxiHost.Name + "/wsman"
                    $resourceURI = "http://schema.omc-project.org/wbem/wscim/1/cim-schema/2/OMC_MCFirmwareIdentity"
                    $rac = Get-WSManInstance -Authentication basic -ConnectionURI $uri -Credential $credential -Enumerate -Port 443 -UseSSL -SessionOption $cimOpt -ResourceURI $resourceURI -ErrorAction SilentlyContinue -ErrorVariable err 
                    if ($rac.VersionString) {
                        $bmcFirmware = $rac.VersionString
                    } #END if
                }
                catch {
                    Write-Verbose -Message ((Get-Date -Format G) + "`tWSMan session failed, error:")
                    Write-Verbose -Message ((Get-Date -Format G) + "`t$err")
                } #END try/catch
            } #END if/else
    
            <#
              Use a custom object to store
              collected data
            #>
            $output = [PSCustomObject]@{
                'Hostname'           = $esxiHost.Name
                'Management IP'      = $mgmtIP
                'RAC IP'             = $racIP
                'RAC MAC'            = $racMAC
                'RAC Firmware'       = $bmcFirmware
                'Product'            = $vmhostView.Config.Product.Name
                'Version'            = $esxiVersion
                'Build'              = $esxiHost.Build
                'Make'               = $hostHardware.Manufacturer
                'Model'              = $hostHardware.Model
                'S/N'                = $hardwarePlatfrom.serialNumber
                'BIOS'               = $hostHardware.BiosVersion
                'BIOS Release Date'  = (($vmhostView.Hardware.BiosInfo.ReleaseDate -split " ")[0])
                'CPU Model'          = $hostHardware.CpuModel -replace '\s+', ' '
                'CPU Count'          = $hostHardware.CpuCount
                'CPU Core Total'     = $hostHardware.CpuCoreCountTotal
                'Speed (MHz)'        = $hostHardware.MhzPerCpu
                'Memory (GB)'        = $esxiHost.MemoryTotalGB -as [int]
                'Memory Slots Count' = $hostHardware.MemorySlotCount
                'Memory Slots Used'  = $hostHardware.MemoryModules.Count
                'Power Supplies'     = $hostHardware.PowerSupplies.Count
                'NIC Count'          = $hostHardware.NicCount
            } #END [PSCustomObject]
            [void]$hardwareCollection.Add($output)
        } #END if
    
        <#
          Get ESXi configuration details
        #>
        if ($Configuration) {
            Write-Output "`tGathering configuration details from $esxiHost ..."

            <#
              Get ESXi licensing
              and software configuration
            #>
            $vmhostID = $vmhostView.Config.Host.Value
            $vmhostLM = $licenseManagerAssign.QueryAssignedLicenses($vmhostID)
            $vmhostPatch = $esxcli.software.vib.list.Invoke() | Where-Object {$_.ID -match $esxiHost.Build} | Select-Object -First 1
            $vmhostvDC = $esxiHost | Get-Datacenter | Select-Object -ExpandProperty Name
            $vmhostCluster = $esxiHost | Get-Cluster | Select-Object -ExpandProperty Name
            $imageProfile = $esxcli.software.profile.get.Invoke()
                   
            <#
              Get services configuration
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering services configuration...")
            $vmServices = $esxiHost | Get-VMHostService
            $vmhostFireWall = $esxiHost | Get-VMHostFirewallException
            $ntpServerList = $esxiHost | Get-VMHostNtpServer
            $ntpService = $vmServices | Where-Object {$_.key -eq "ntpd"}
            $ntpFWException = $vmhostFireWall | Select-Object -Property Name, Enabled | Where-Object {$_.Name -eq "NTP Client"}
            $sshService = $vmServices | Where-Object {$_.key -eq "TSM-SSH"}
            $sshServerFWException = $vmhostFireWall | Select-Object -Property Name, Enabled | Where-Object {$_.Name -eq "SSH Server"}
            $esxiShellService = $vmServices | Where-Object {$_.key -eq "TSM"}
            $ShellTimeOut = (Get-AdvancedSetting -Entity $esxiHost -Name "UserVars.ESXiShellTimeOut" -ErrorAction SilentlyContinue).value
            $interactiveShellTimeOut = (Get-AdvancedSetting -Entity $esxiHost -Name "UserVars.ESXiShellInteractiveTimeOut" -ErrorAction SilentlyContinue).value
    
            <#
              Get syslog configuration
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering Syslog Configuration...")
            $syslogList = @()
            $syslogFWException = $vmhostFireWall | Select-Object -Property Name, Enabled | Where-Object {$_.Name -eq "syslog"}
            foreach ($syslog in  $esxiHost | Get-VMHostSysLogServer) {
                $syslogList += $syslog.Host + ":" + $syslog.Port
            } #END foreach
    
            <#
              Get UpTime and install date
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering UpTime Configuration...")
            $bootTimeUTC = $vmhostView.Runtime.BootTime
            $BootTime = $bootTimeUTC.ToLocalTime()
            $upTime = New-TimeSpan -Seconds $vmhostView.Summary.QuickStats.Uptime
            $upTimeDays = $upTime.Days
            $upTimeHours = $upTime.Hours
            $upTimeMinutes = $upTime.Minutes
            $vmUUID = $esxcli.system.uuid.get.Invoke()
            $decimalDate = [Convert]::ToInt32($vmUUID.Split("-")[0], 16)
            $installDate = ([DateTime]'1/1/1970').AddSeconds($decimalDate).ToLocalTime()

            <#
              Get ESXi installation type
              and Boot config
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering ESXi installation type...")
            $bootDevice = $esxcli.system.boot.device.get.Invoke()
            if ($bootDevice.BootFilesystemUUID) {
                if ($bootDevice.BootFilesystemUUID[6] -eq 'e') {
                    $installType = "Embedded"
                }
                else {
                    $installType = "Installable"
                    $bootSource = $esxcli.storage.filesystem.list.Invoke() | Where-Object {$_.UUID -eq $bootDevice.BootFilesystemUUID} | Select-Object -ExpandProperty MountPoint
                } #END if/else
                $storageDevice = $esxcli.storage.core.device.list.Invoke() | Where-Object {$_.IsBootDevice -eq $true}
                $bootVendor = $storageDevice.Vendor + " " + $storageDevice.Model
                $bootDisplayName = $storageDevice.DisplayName
                $bootPath = $storageDevice.DevfsPath
                $storagePath = $esxcli.storage.core.path.list.Invoke() | Where-Object {$_.Device -eq $storageDevice.Device}
                $bootRuntime = $storagePath.RuntimeName
                if ($installType -eq "Embedded") {
                    $bootSource = $storageDevice.DisplayName.Split('(')[0]
                }
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
                $bootVendor = $null
                $bootDisplayName = $null
                $bootPath = $null
                $bootRuntime = $null
            } #END if/else

            <#
              Use a custom object to store
              collected data
            #>
            $output = [PSCustomObject]@{
                'Hostname'                  = $esxiHost.Name
                'Make'                      = $hostHardware.Manufacturer
                'Model'                     = $hostHardware.Model
                'CPU Model'                 = $hostHardware.CpuModel -replace '\s+', ' '
                'Hyper-Threading'           = $esxiHost.HyperthreadingActive
                'Max EVC Mode'              = $esxiHost.MaxEVCMode
                'Product'                   = $vmhostView.Config.Product.Name
                'Version'                   = $esxiVersion
                'Build'                     = $esxiHost.Build
                'Install Type'              = $installType
                'Boot From'                 = $bootSource
                'Device Model'              = $bootVendor
                'Boot Device'               = $bootDisplayName
                'Runtime Name'              = $bootRuntime
                'Device Path'               = $bootPath
                'Image Profile'             = $imageProfile.Name
                'Acceptance Level'          = $imageProfile.AcceptanceLevel 
                'Boot Time'                 = $BootTime
                'Uptime'                    = "$upTimeDays Day(s), $upTimeHours Hour(s), $upTimeMinutes Minute(s)"
                'Install Date'              = $installDate
                'Last Patched'              = $vmhostPatch.InstallDate
                'License Version'           = $vmhostLM.AssignedLicense.Name | Select-Object -Unique
                'License Key'               = $vmhostLM.AssignedLicense.LicenseKey | Select-Object -Unique
                'Connection State'          = $esxiHost.ConnectionState
                'Standalone'                = $esxiHost.IsStandalone
                'Cluster'                   = $vmhostCluster
                'Virtual Datacenter'        = $vmhostvDC
                'vCenter'                   = $vmhostView.CLient.ServiceUrl.Split('/')[2]
                'NTP'                       = $ntpService.Label
                'NTP Running'               = $ntpService.Running
                'NTP Startup Policy'        = $ntpService.Policy
                'NTP Client Enabled'        = $ntpFWException.Enabled
                'NTP Server'                = (@($ntpServerList) -join ',')
                'SSH'                       = $sshService.Label
                'SSH Running'               = $sshService.Running
                'SSH Startup Policy'        = $sshService.Policy
                'SSH TimeOut'               = $ShellTimeOut
                'SSH Server Enabled'        = $sshServerFWException.Enabled
                'ESXi Shell'                = $esxiShellService.Label
                'ESXi Shell Running'        = $esxiShellService.Running
                'ESXi Shell Startup Policy' = $esxiShellService.Policy
                'ESXi Shell TimeOut'        = $interactiveShellTimeOut
                'Syslog Server'             = (@($syslogList) -join ',')
                'Syslog Client Enabled'     = $syslogFWException.Enabled
            } #END [PSCustomObject]
            [void]$configurationCollection.Add($output)
        } #END if
    } #END foreach
    $stopWatch.Stop()
    Write-Verbose -Message ((Get-Date -Format G) + "`tMain code execution completed")
    Write-Verbose -Message  ((Get-Date -Format G) + "`tScript Duration: " + $stopWatch.Elapsed.Duration())
    
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
        Write-Host "`n" "ESXi Hardware Inventory:" -ForegroundColor Green
        if ($ExportCSV) {
            $hardwareCollection | Export-Csv ($outputFile + "Hardware.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "Hardware.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $hardwareCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Hardware_Inventory -NoNumberConversion * -AutoSize -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $returnCollection += $hardwareCollection 
            $returnCollection 
        }
        else {
            $hardwareCollection | Format-List
        }#END if/else
    } #END if
    
    if ($configurationCollection) {
        Write-Host "`n" "ESXi Host Configuration:" -ForegroundColor Green
        if ($ExportCSV) {
            $configurationCollection | Export-Csv ($outputFile + "Configuration.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "Configuration.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $configurationCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Host_Configuration -NoNumberConversion * -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $returnCollection += $configurationCollection
            $returnCollection  
        }
        else {
            $configurationCollection | Format-List
        }#END if/else
    } #END if
} #END function