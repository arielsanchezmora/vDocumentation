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
       File Name    : Get-ESXIODevice.ps1
       Author       : Edgar Sanchez - @edmsanchez13
       Contributor  : Ariel Sanchez - @arielsanchezmor
       Version      : 2.4.3     
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
    $date = $date -replace ":", "-"
    $outputFile = "IODevice" + $date
    
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
                        if ($tempList) {
                            $vHostList += $tempList | Sort-Object -Property Name
                        }
                        else {
                            Write-Warning -Message "`tDatacenter with name $vDCname was not found in $Global:DefaultViServers"
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
                if ($tempList) {
                    $vHostList += $tempList | Sort-Object -Property Name
                }
                else {
                    Write-Warning -Message "`tCluster with name $vClusterName was not found in $Global:DefaultViServers"
                } #END if/else
            } #END foreach
        } #END if/else
    }
    else { 
        Write-Verbose -Message ((Get-Date -Format G) + "`tExecuting Cmdlet using esxi parameter")
        Write-Host "`tGathering host list..."
        foreach ($invidualHost in $esxi) {
            $tempList = Get-VMHost -Name $invidualHost.Trim() -ErrorAction SilentlyContinue
            if ($tempList) {
                $vHostList += $tempList | Sort-Object -Property Name
            }
            else {
                Write-Warning -Message "`tESXi host $invidualHost was not found in $Global:DefaultViServers"
            } #END if/else
        } #END foreach
    } #END if/else
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
      Main code execution
    #>
    foreach ($vmhost in $vHostList) {

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
                'Hostname'         = $vmhost.Name
                'Connection State' = $vmhost.ConnectionState
            } #END [PSCustomObject]
            continue
        } #END if/else
        $esxcli = Get-EsxCli -VMHost $vmhost -V2
    
        <#
          Get IO Device details
        #>
        Write-Host "`tGathering information from $vmhost ..."
        $pciDevices = $esxcli.hardware.pci.list.Invoke() | Where-Object {$_.VMKernelName -like "vmhba*" -or $_.VMKernelName -like "vmnic*" -or $_.VMKernelName -like "vmgfx*"} | Sort-Object -Property VMKernelName 
        foreach ($pciDevice in $pciDevices) {
            $device = $vmhost | Get-VMHostPciDevice | Where-Object {$pciDevice.Address -match $_.Id}
            Write-Verbose -Message ((Get-Date -Format G) + "`tGet driver version for: " + $pciDevice.ModuleName)
            $driverVersion = $esxcli.system.module.get.Invoke(@{module = $pciDevice.ModuleName}) | Select-Object -ExpandProperty Version
    
            <#
              Get NIC Firmware version
            #>
            if ($pciDevice.VMKernelName -like 'vmnic*') {
                Write-Verbose -Message ((Get-Date -Format G) + "`tGet Firmware version for: " + $pciDevice.VMKernelName)
                $vmnicDetail = $esxcli.network.nic.get.Invoke(@{nicname = $pciDevice.VMKernelName})
                $firmwareVersion = $vmnicDetail.DriverInfo.FirmwareVersion
                
                <#
                  Get NIC driver VIB package version
                #>
                Write-Verbose -Message ((Get-Date -Format G) + "`tGet VIB details for: " + $pciDevice.ModuleName)
                $driverVib = $esxcli.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object {$_.Name -eq $vmnicDetail.DriverInfo.Driver -or $_.Name -eq "net-" + $vmnicDetail.DriverInfo.Driver -or $_.Name -eq "net55-" + $vmnicDetail.DriverInfo.Driver}
                $vibName = $driverVib.Name
                $vibVersion = $driverVib.Version
    
                <#
                  If HP Smart Array vmhba* (scsi-hpsa driver) then get Firmware version
                  elese skip if VMkernnel is vmhba*. Can't get HBA Firmware from 
                  Powercli at the moment only through SSH or using Putty Plink+PowerCli.
                #>
            }
            elseif ($pciDevice.VMKernelName -like 'vmhba*') {
                if ($pciDevice.DeviceName -match "smart array") {
                    Write-Verbose -Message ((Get-Date -Format G) + "`tGet Firmware version for: " + $pciDevice.VMKernelName)
                    $hpsa = $vmhost.ExtensionData.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo | Where-Object {$_.Name -match "HP Smart Array"}
                    if ($hpsa) {
                        $firmwareVersion = (($hpsa.Name -split "firmware")[1]).Trim()
                    }
                    else {
                        Write-Verbose -Message ((Get-Date -Format G) + "`tGet Extension data failed. Skip Firmware version check for: " + $pciDevice.DeviceName)
                        $firmwareVersion = $null    
                    } #END if/else
                }
                else {
                    Write-Verbose -Message ((Get-Date -Format G) + "`tSkip Firmware version check for: " + $pciDevice.DeviceName)
                    $firmwareVersion = $null    
                } #END if/else
                        
                <#
                  Get HBA driver VIB package version
                #>
                Write-Verbose -Message ((Get-Date -Format G) + "`tGet VIB deatils for: " + $pciDevice.ModuleName)
                $vibName = $pciDevice.ModuleName -replace "_", "-"
                $driverVib = $esxcli.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object {$_.Name -eq "scsi-" + $VibName -or $_.Name -eq "sata-" + $VibName -or $_.Name -eq $VibName}
                $vibName = $driverVib.Name
                $vibVersion = $driverVib.Version
            }
            else {
                Write-Verbose -Message ((Get-Date -Format G) + "`tSkipping: " + $pciDevice.DeviceName)
                $firmwareVersion = $null
                $vibName = $null
                $vibVersion = $null
            } #END if/else
           
            <#
              Use a custom object to store
              collected data
            #>
            $outputCollection += [PSCustomObject]@{
                'Hostname'         = $vmhost.Name
                'Slot Description' = $pciDevice.SlotDescription
                'VMKernel Name'    = $pciDevice.VMKernelName
                'Device Name'      = $pciDevice.DeviceName
                'Vendor Name'      = $pciDevice.VendorName
                'Device Class'     = $pciDevice.DeviceClassName
                'PCI Address'      = $pciDevice.Address
                'VID'              = [String]::Format("{0:x4}", $device.VendorId)
                'DID'              = [String]::Format("{0:x4}", $device.DeviceId)
                'SVID'             = [String]::Format("{0:x4}", $device.SubVendorId)
                'SSID'             = [String]::Format("{0:x4}", $device.SubDeviceId)
                'Driver'           = $pciDevice.ModuleName
                'Driver Version'   = $driverVersion
                'Firmware Version' = $firmwareVersion
                'VIB Name'         = $vibName
                'VIB Version'      = $vibVersion
            } #END [PSCustomObject]
        } #END foreach
    } #END foreach
    Write-Verbose -Message ((Get-Date -Format G) + "`tMain code execution completed")
    
    <#
      Display skipped hosts and their connection status
    #>
    If ($skipCollection) {
        Write-Warning -Message "`tCheck Connection State or Host name "
        Write-Warning -Message "`tSkipped hosts: "
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
        }
        elseif ($ExportExcel) {
            $outputCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname IO_Device -NoNumberConversion * -AutoSize -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $outputCollection
        }
        else {
            $outputCollection | Format-List
        }#END if/else
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tNo information gathered")
    } #END if/else
} #END function