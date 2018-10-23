function Get-ESXIODevice {
    <#
     .SYNOPSIS
       Get ESXi IO VMKernel device information
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
       Contributor : @pdpelsem
       Version      : 2.4.5
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
       Get-ESXIODevice -VMhost devvm001.lab.local
     .PARAMETER Cluster
       The name(s) of the vSphere Cluster(s)
     .EXAMPLE
       Get-ESXIODevice -Cluster production-cluster
     .PARAMETER Datacenter
       The name(s) of the vSphere Virtual DataCenter(s)
     .EXAMPLE
       Get-ESXIODevice -Datacenter vDC001
     .PARAMETER ExportCSV
       Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-ESXIODevice -Cluster production-cluster -ExportCSV
     .PARAMETER ExportExcel
       Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
       ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
       File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-ESXIODevice -Cluster production-cluster -ExportExcel
     .PARAMETER folderPath
       Specificies an alternate folder path of where the exported file should be saved.
     .EXAMPLE
       Get-ESXIODevice -Cluster production-cluster -ExportExcel -folderPath C:\temp
     .PARAMETER PassThru
       Returns the object to console
     .EXAMPLE
       Get-ESXIODevice -VMhost devvm001.lab.local -PassThru
    #> 
    
    <#
     ----------------------------------------------------------[Declarations]----------------------------------------------------------
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false,
            ParameterSetName = "VMhost")]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String[]]$VMhost,
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
        [switch]$PassThru,
        $folderPath
    )
    
    $ioDeviceCollection = [System.Collections.ArrayList]@()
    $ioDeviceHclCollection = [System.Collections.ArrayList]@()
    $hclIoCollection = [System.Collections.ArrayList]@()
    $skipCollection = @()
    $vHostList = @()
    $date = Get-Date -format s
    $date = $date -replace ":", "-"
    $outputFile = "IODevice" + $date
    
    <#
     ----------------------------------------------------------[Execution]----------------------------------------------------------
    #>

    $stopWatch = [system.diagnostics.stopwatch]::startNew()

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
      Gather host list based on parameter used
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
        Write-Output "`tGathering host list from the following Cluster(s): " (@($Cluster) -join ',')
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
        Write-Output "`tGathering host list from the following DataCenter(s): " (@($DataCenter) -join ',')
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
      Query HCL website for IO device
      Information
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidating access to VMware HCL site...")
    $hclDataUrl = "https://www.vmware.com/resources/compatibility/js/data_io.js?"
    try {
        $webRequest = Invoke-WebRequest -Uri $hclDataUrl
    } 
    catch [System.Net.WebException] {
        $webRequest = $_.Exception.Response
    } #END try/catch
    if ([int]$webRequest.StatusCode -eq "200") {
        $webContent = $webRequest.Content
        $webContent = $webContent.Remove(0, $webContent.IndexOf('window.compdata'))
        $webContentSubstring = $webContent.Substring($webContent.IndexOf("=") + 2, $webContent.IndexOf("];") - $webContent.IndexOf("=") - 1).trim()
        $webContentArray = $webContentSubstring.Split("`n")
        foreach ($line in $webContentArray) {
            $line = $line.Trim()
            if ($line.StartsWith('["')) {
                $lineSubstring = $line.Substring($line.IndexOf('["') + 1, $line.IndexOf('["', 1) - $line.IndexOf('["') - 1)
                $lineSubstring = $lineSubstring.Replace('",' , '"|')
                $lineArray = $lineSubstring.split("|")

                <#
                  Use a custom object to store
                  collected data
                #>
                $output = [PSCustomObject]@{
                    'productid' = ($lineArray[0].Replace('"', "")).trim()
                    'brandname' = ($lineArray[1].Replace('"', "")).trim()
                    'model'     = ($lineArray[2].Replace('"', "")).trim()
                    'vid'       = ($lineArray[4].Replace('"', "")).trim()
                    'did'       = ($lineArray[5].Replace('"', "")).trim()
                    'svid'      = ($lineArray[6].Replace('"', "")).trim()
                    'ssid'      = ($lineArray[7].Replace('"', "")).trim()
                } #END [PSCustomObject]
                [void]$hclIoCollection.Add($output)
            } #END if
        } #END foreach
    }
    else {
        Write-Verbose -Message ("`tVMware HCL site: '$hclDataUrl' is NOT reachable/unavailable (Return code: " + ([int]$webRequest.StatusCode) + ")")
    } #END if/else   
    $hclDataUrl = "https://www.vmware.com/resources/compatibility/search.php?deviceCategory=io"
    $brandJson = $null
    try {
        $webRequest = Invoke-WebRequest -Uri $hclDataUrl
    } 
    catch [System.Net.WebException] {
        $webRequest = $_.Exception.Response
    } #END try/catch
    if ([int]$webRequest.StatusCode -eq "200") {
        $webElement = $webRequest.ParsedHtml.body.getElementsByTagName("script") | Where-Object { $_.type -eq "text/javascript"}
        foreach ($element in $webElement) {
            if ($null -ne $element.innerHTML -and $element.innerHTML.trim().StartsWith("var releases")) {
                $webElementHtml = $element.innerHTML.trim()
            } #END if
        } #END foreach
        $brandElement = $webElementHtml.Remove(0, $webElementHtml.IndexOf('var partners'))
        $brandSubstring = $brandElement.Substring($brandElement.IndexOf("=") + 1, $brandElement.IndexOf("};") + 1 - $brandElement.IndexOf("=") - 1).trim()
        $brandJson = $brandSubstring | ConvertFrom-Json -ErrorAction SilentlyContinue
    }
    else {
        Write-Verbose -Message ("`tVMware HCL site: '$hclDataUrl' is NOT reachable/unavailable (Return code: " + ([int]$webRequest.StatusCode) + ")")
    } #END if/else   

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
    
        <#
          Get IO Device details
        #>
        Write-Host "`tGathering information from $esxiHost ..."
        $esxiUpdateLevel = (Get-AdvancedSetting -Name "Misc.HostAgentUpdateLevel" -Entity $esxiHost -ErrorAction SilentlyContinue -ErrorVariable err).Value
        if ($esxiUpdateLevel) {
            $esxiVersion = "ESXi " + ($esxiHost.ApiVersion) + " U" + $esxiUpdateLevel
        }
        else {
            $esxiVersion = "ESXi " + ($esxiHost.ApiVersion)
            Write-Verbose -Message ((Get-Date -Format G) + "`tFailed to get ESXi Update Level, Error : " + $err)
        } #END if/else
        $pciDevices = $esxcli.hardware.pci.list.Invoke() | Where-Object {$_.VMKernelName -like "vmhba*" -or $_.VMKernelName -like "vmnic*" -or $_.VMKernelName -like "vmgfx*"} | Sort-Object -Property VMKernelName 
        foreach ($pciDevice in $pciDevices) {
            $device = $esxiHost | Get-VMHostPciDevice | Where-Object {$pciDevice.Address -match $_.Id}
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
            }
            elseif ($pciDevice.VMKernelName -like 'vmhba*') {

                <#
                  If HP Smart Array vmhba* (scsi-hpsa driver) then get Firmware version
                  else skip if VMkernnel is vmhba*. Can't get HBA Firmware from 
                  Powercli at the moment only through SSH or using Putty Plink+PowerCli.
                #>
                if ($pciDevice.DeviceName -match "smart array") {
                    Write-Verbose -Message ((Get-Date -Format G) + "`tGet Firmware version for: " + $pciDevice.VMKernelName)
                    $hpsa = $esxiHost.ExtensionData.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo | Where-Object {$_.Name -match "HP Smart Array"}
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
                Write-Verbose -Message ((Get-Date -Format G) + "`tGet VIB details for: " + $pciDevice.ModuleName)
                $vibName = $pciDevice.ModuleName -replace "_", "-"
                $driverVib = $esxcli.software.vib.list.Invoke() | Select-Object -Property Name, Version | Where-Object {$_.Name -eq "scsi-" + $VibName -or $_.Name -eq "sata-" + $VibName -or $_.Name -eq $VibName}
                $vibName = $driverVib.Name
            }
            else {
                Write-Verbose -Message ((Get-Date -Format G) + "`tSkipping: " + $pciDevice.DeviceName)
                $firmwareVersion = $null
                $vibName = $null
            } #END if/else

            <#
              Get HCL IDs and build URL
            #>
            $vid = [String]::Format("{0:x4}", $device.VendorId)
            $did = [String]::Format("{0:x4}", $device.DeviceId)
            $svid = [String]::Format("{0:x4}", $device.SubVendorId)
            $ssid = [String]::Format("{0:x4}", $device.SubDeviceId)
            $hclUrl = "https://www.vmware.com/resources/compatibility/search.php?deviceCategory=io&VID=$vid&DID=$did&SVID=$svid&SSID=$ssid&details=1"
            $productId = $hclIoCollection | Where-Object {$_.vid -match $vid -and $_.did -match $did -and $_.svid -match $svid -and $_.ssid -match $ssid} | Select-Object -ExpandProperty productid
            
            <#
              Use a custom object to store
              collected data
            #>
            $output = [PSCustomObject]@{
                'Hostname'         = $esxiHost.Name
                'Version'          = $esxiVersion
                'Slot Description' = $pciDevice.SlotDescription
                'VMKernel Name'    = $pciDevice.VMKernelName
                'Device Name'      = $pciDevice.DeviceName
                'Vendor Name'      = $pciDevice.VendorName
                'Device Class'     = $pciDevice.DeviceClassName
                'PCI Address'      = $pciDevice.Address
                'VID'              = $vid
                'DID'              = $did
                'SVID'             = $svid
                'SSID'             = $ssid
                'VIB Name'         = $vibName
                'Driver'           = $pciDevice.ModuleName
                'Driver Version'   = $driverVersion
                'Firmware Version' = $firmwareVersion
                'HCL URL'          = $hclUrl
                'ProductId'        = (@($productId) -join ',')
            } #END [PSCustomObject]
            [void]$ioDeviceCollection.Add($output)
        } #END foreach
    } #END foreach

    <#
      Get HCL IO device Details
    #>
    if ($hclIoCollection -and $brandJson) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering IO device HCL details...")
        $ioDevices = $ioDeviceCollection | Sort-Object -Property ProductId,Version -Unique
        foreach ($ioDevice in $ioDevices) {
            if ($ioDevice.ProductId) {
                $productIds = $ioDevice.ProductId.Split(',')
                foreach ($productId in $productIds) {
                    $vid = $ioDevice.VID
                    $did = $ioDevice.DID
                    $svid = $ioDevice.SVID
                    $ssid = $ioDevice.SSID
                    $hclDataUrl = "https://www.vmware.com/resources/compatibility/detail.php?deviceCategory=io&productid=$productId&deviceCategory=io&details=1&VID=$vid&DID=$did&SVID=$svid&SSID=$ssid&page=1&display_interval=10&sortColumn=Partner&sortOrder=Asc"
                    try {
                        $webRequest = Invoke-WebRequest -Uri $hclDataUrl
                    } 
                    catch [System.Net.WebException] {
                        $webRequest = $_.Exception.Response
                    } #END try/catch
                    if ([int]$webRequest.StatusCode -eq "200") {
                        $webElement = $webRequest.ParsedHtml.body.getElementsByTagName("script") | Where-Object { $_.type -eq "text/javascript"}
                        foreach ($element in $webElement) {
                            if ($null -ne $element.innerHTML -and $element.innerHTML.trim().StartsWith("var details")) {
                                $webElementHtml = $element.innerHTML.trim()
                            } #END if
                        } #END foreach    
                        $detailSubstring = $webElementHtml.Substring($webElementHtml.IndexOf("=[") + 1, $webElementHtml.IndexOf("];") + 1 - $webElementHtml.IndexOf("=[") - 1).trim()
                        $detailJson = $detailSubstring | ConvertFrom-Json -ErrorAction SilentlyContinue
                        $deviceList = $detailJson | Where-Object {$_.ReleaseVersion -eq $ioDevice.Version}

                        <#
                          Get supported features
                        #>
                        $featureElement = $webElementHtml.Remove(0, $webElementHtml.IndexOf('var cert_features'))
                        $featureSubstring = $featureElement.Substring($featureElement.IndexOf("={") + 1, $featureElement.IndexOf("};") + 1 - $featureElement.IndexOf("={") - 1).trim()
                        $featureJson = $featureSubstring | ConvertFrom-Json -ErrorAction SilentlyContinue

                        foreach ($device in $deviceList) {
                            $deviceInfo = $hclIoCollection | Where-Object {$_.productid -eq $device.Component_Id}
                            $partner = $deviceInfo.brandname
                            $deviceDriver = $device.DriverName + " Version " + $device.Version
                            $driverType = $device.inbox_async + "," + $device.VmklinuxOrNativeDriver
                            $certDetailId = $device.CertDetail_Id.ToString() + "-1"
                            $deviceFeatures = $null
                            if ($featureJson.$certDetailId.Count -gt 0) {
                                $deviceFeatures = $featureJson.$certDetailId[5]                                    
                            } #END if

                            <#
                              Use a custom object to store
                              collected data
                            #>
                            $output = [PSCustomObject]@{
                                'Model'                      = $deviceInfo.model
                                'Device Type'                = $device.DeviceType
                                'Brand Name'                 = $brandJson.$partner
                                'VID'                        = $deviceInfo.vid
                                'DID'                        = $deviceInfo.did
                                'SVID'                       = $deviceInfo.svid
                                'SSID'                       = $deviceInfo.ssid
                                'Release'                    = $device.ReleaseVersion
                                'Device Driver(s)'           = $deviceDriver
                                'Firmware Version'           = $device.FirmwareVersion
                                'Additional Firmare Version' = $device.AddlFirmwareVersion
                                'Type'                       = $driverType
                                'Features'                   = $deviceFeatures
                            } #END [PSCustomObject]
                            [void]$ioDeviceHclCollection.Add($output)
                        } #END foreach
                    }
                    else {
                        Write-Verbose -Message ("`tVMware HCL site: '$hclDataUrl' is NOT reachable/unavailable (Return code: " + ([int]$webRequest.StatusCode) + ")")
                    } #END if/else   
                } #END foreach
            }
            else {
                Write-Verbose -Message ((Get-Date -Format G) + "`t" + $ioDevice.'Device Name' + " has no Product ID. Skipping...")
                continue
            } #END if/else
            
        } #END foreach
    } #END if
    $stopWatch.Stop()
    Write-Verbose -Message ((Get-Date -Format G) + "`tMain code execution completed")
    Write-Verbose -Message  ((Get-Date -Format G) + "`tScript Duration: " + $stopWatch.Elapsed.Duration())
    
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
    if ($ioDeviceCollection -or $ioDeviceHclCollection) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tInformation gathered")
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tNo information gathered")
    } #END if/else
    
    <#
      Output to screen
      Export data to CSV, Excel
    #>
    if ($ioDeviceCollection) {
        Write-Host "`n" "ESXi IO Device:" -ForegroundColor Green
        if ($ExportCSV) {
            $ioDeviceCollection | Export-Csv ($outputFile + "IODevice.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "IODevice.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $ioDeviceCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname IO_Device -NoNumberConversion * -AutoSize -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $ioDeviceCollection
        }
        else {
            $ioDeviceCollection | Format-List
        }#END if/else
    } #END if
    if ($ioDeviceHclCollection) {
        Write-Host "`n" "VMware HCL Details:" -ForegroundColor Green
        if ($ExportCSV) {
            $ioDeviceHclCollection | Export-Csv ($outputFile + "HclDetails.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "HclDetails.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $ioDeviceHclCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname IO_HCL_Details -NoNumberConversion * -AutoSize -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $ioDeviceHclCollection
        }
        else {
            $ioDeviceHclCollection | Format-List
        }#END if/else
    } #END if
} #END function