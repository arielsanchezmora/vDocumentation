function Get-ESXSpeculativeExecution {
    <#
     .SYNOPSIS
       Get ESXi host compliance on VMSA-2018-0004 Security Advisory and BIOS version
     .DESCRIPTION
       Will validate ESXi host for VMware Security Advisory VMSA-2018-0004 Compliance (https://www.vmware.com/us/security/advisories/VMSA-2018-0004.html)
       and validate against BIOS version
     .NOTES
       Author     : Edgar Sanchez - @edmsanchez13
       Contributor: Ariel Sanchez - @arielsanchezmor
     .Link
       https://github.com/arielsanchezmora/vDocumentation
     .INPUTS
       If onlince CSV file is inaccesible then a CSV file path input must be provided to check against BIOS version
     .OUTPUTS
       CSV file
       Excel file
     .PARAMETER esxi
       The name(s) of the vSphere ESXi Host(s)
     .EXAMPLE
       Get-ESXSpeculativeExecution -esxi devvm001.lab.local
     .PARAMETER cluster
       The name(s) of the vSphere Cluster(s)
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster
     .PARAMETER datacenter
       The name(s) of the vSphere Virtual Datacenter(s).
     .EXAMPLE
       Get-ESXSpeculativeExecution -datacenter vDC001
       Get-ESXSpeculativeExecution -datacenter "all vdc" will gather all hosts in vCenter(s). This is the default if no Parameter (-esxi, -cluster, or -datacenter) is specified.
     .PARAMETER inputFile
       Specify input CSV file containing server BIOS version to validate against. Use this if you do not have access to the internet or wish to use your own offline version
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster -inputFile "c:\temp\BIOSUpdates.csv"
     .PARAMETER ExportCSV
       Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster -ExportCSV
     .PARAMETER ExportExcel
       Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
       ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
       File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster -ExportExcel
     .PARAMETER PatchCompliance
       Switch to get ESXi host compliance on VMSA-2018-0002
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster -PatchCompliance
     .PARAMETER BIOSCompliance
       Switch to get system BIOS version validated against CSV file
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster -BIOSCompliance
     .PARAMETER ReportOnVMs
       Switch to report on Virtual machines. This needs to be specified if needed, or it will not be included
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster -ReportOnVMS
  
     .PARAMETER folderPath
       Specify an alternate folder path where the exported data should be saved.
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster -ExportExcel -folderPath C:\temp
     .PARAMETER PassThru
       Switch to return object to command line
     .EXAMPLE
       Get-ESXSpeculativeExecution -esxi 192.168.1.100 -Hardware -PassThru
    #> 
    
    <#
     ----------------------------------------------------------[Declarations]----------------------------------------------------------
    #>
  
    [CmdletBinding()]
    param (
        $esxi,
        $cluster,
        $datacenter,
        $inputFile,
        [switch]$ExportCSV,
        [switch]$ExportExcel,
        [switch]$PatchCompliance,
        [switch]$BIOSCompliance,
        [switch]$ReportOnVMs,
        [switch]$PassThru,
        $folderPath
    )
    
    $PatchCollection = @()
    $BIOSCollection = @()
    $skipCollection = @()
    $VMCollection = @()
    $vHostList = @()
    $ReturnCollection = @()
    $csvCollection = @()
    $url = 'https://raw.githubusercontent.com/edmsanchez/vDocumentation/master/powershell/vDocumentation/BIOSUpdates.csv'
    $date = Get-Date -format s
    $date = $date -replace ":", "-"
    $outputFile = "SpeculativeExecution" + $date
    
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
      -PatchCompliance, -BIOSCompliance. By default all are executed
      unless one is specified. -ReportOnVMS needs to be mandatory specified for it to be used.
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate Cmdlet switches")
    if ($PatchCompliance -or $BIOSCompliance) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tA Cmdlet switch was specified")
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tA Cmdlet switch (-PatchCompliance -BIOSCompliance) was not specified")
        Write-Verbose -Message ((Get-Date -Format G) + "`tWill execute all (-PatchCompliance -BIOSCompliance)")
        $PatchCompliance = $true
        $BIOSCompliance = $true
    } #END if/else
    if ($ReportOnVMs) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tReportOnVMS switch was specified")
    } #END if
    
    <#
      Validate access to online CSV file or if inputFile parameter was specified for BIOS Compliance check
    #>
    if ($BIOSCompliance) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tInitializing -BIOSCompliance Cmdlet switch variables...")
        if ([string]::IsNullOrWhiteSpace($inputFile)) {
            Write-Verbose -Message ((Get-Date -Format G) + "`t-inputFile parameter is Null or Empty")
            Write-Verbose -Message ((Get-Date -Format G) + "`tValidating access to online CSV file")
            try {
                $webRequest = Invoke-WebRequest -Uri $url
            } 
            catch [System.Net.WebException] {
                $webRequest = $_.Exception.Response
            } #END try
            if ([int]$webRequest.StatusCode -eq "200") {
                $csvCollection = $webRequest | ConvertFrom-Csv
            }
            else {
                Write-Warning -Message ("`tonline CSV fie: '$url' is NOT reachable/unavailable (Return code: " + ([int]$webRequest.StatusCode) + ")")
                Write-Warning -Message "`tYou can download the file locally and use it offline with the -inputFile parameter. See help for more information"
                Write-Warning -Message "`tSkipping BIOS Compliance check..."
                $BIOSCompliance = $false
            } #END if/else   
        }
        else {
            if (Test-Path $inputFile) {
                Write-Verbose -Message ((Get-Date -Format G) + "`t'$inputFile' path found")
                $csvCollection = Import-Csv -Path $inputFile
            }
            else {
                Write-Warning -Message "`t'$inputFile' path not found."
                Write-Warning -Message "`tSkipping BIOS Compliance check..."
            } #END if/else
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
            $skipCollection += [PSCustomObject]@{
                'Hostname'         = $esxihost
                'Connection State' = $esxihost.ConnectionState
            } #END [PSCustomObject]
            continue
        } #END if/else
        $esxcli2 = Get-EsxCli -VMHost $esxihost -V2
    
        <#
          Get ESXi version details
        #>
        $vmhostView = $vmhost | Get-View
        $esxiVersion = $esxcli2.system.version.get.Invoke()
                    
        <#
          Get Patch compliance details
        #>
        if ($PatchCompliance) {
            Write-Host "`tValidating host compliance on $vmhost ..."

            <#
              Validate only against applicable ESXi Versions
              check for CPU-Microcode patch also
            #>
            if ($vmhost.ApiVersion -eq "5.5" -or $vmhost.ApiVersion -like "6.*") {
                if ($vmhost.ApiVersion -eq "5.5" -and $vmhost.Build -ge "7504623") {
                    $complianceStatus = "Compliant"
                    $esxMicrocode = "Present"
                    $hostCPUID = $vmhostview.Config.FeatureCapability | Where-Object {$_.FeatureName -eq "cpuid.IBRS" -or $_.Featurename -eq "cpuid.IBPB" -or $_.Featurename -eq "cpuid.STIBP"} | Select-Object -ExpandProperty FeatureName
                }
                elseif ($vmhost.ApiVersion -eq "6.0" -and $vmhost.Build -ge "7504637") {
                    $complianceStatus = "Compliant"
                    $vmhostPatch = $esxcli2.software.vib.list.Invoke() | Where-Object {$_.Name -eq "cpu-microcode"}
                    if ($vmhostPatch) {
                        if ((($vmhostPatch.Version).Split('.') | Select-Object -Last 1) -ge "7504637") {
                            $esxMicrocode = "Present"
                            $hostCPUID = $vmhostview.Config.FeatureCapability | Where-Object {$_.FeatureName -eq "cpuid.IBRS" -or $_.Featurename -eq "cpuid.IBPB" -or $_.Featurename -eq "cpuid.STIBP"} | Select-Object -ExpandProperty FeatureName
                        }
                        else {
                            $esxMicrocode = "NotPresent. ESXi600-201801402-BG Missing see VMSA-2018-0004"
                        } #END if/else
                    } #END if
                }
                elseif ($vmhost.ApiVersion -eq "6.5" -and $vmhost.Build -ge "7526125") {
                    $complianceStatus = "Compliant"
                    $vmhostPatch = $esxcli2.software.vib.list.Invoke() | Where-Object {$_.Name -eq "cpu-microcode"}
                    if ($vmhostPatch) {
                        if ((($vmhostPatch.Version).Split('.') | Select-Object -Last 1) -ge "7526125") {
                            $esxMicrocode = "Present"
                            $hostCPUID = $vmhostview.Config.FeatureCapability | Where-Object {$_.FeatureName -eq "cpuid.IBRS" -or $_.Featurename -eq "cpuid.IBPB" -or $_.Featurename -eq "cpuid.STIBP"} | Select-Object -ExpandProperty FeatureName
                        }
                        else {
                            $esxMicrocode = "NotPresent. ESXi650-201801402-BG Missing see VMSA-2018-0004"
                        } #END if/else
                    } #END if
                }
                else {
                    $complianceStatus = "NotCompliant"
                    $esxMicrocode = "NotPresent"
                    $hostCPUID = $null
                } #END if/ese
            }
            else {
                $complianceStatus = "ESXi upgrade needed"
                $esxMicrocode = "NotPresent"
                $hostCPUID = $null
            } #END if/else

            <#
              Get accurate last patched date if ESXi 6.5
              based on Date and time (UTC), which is
              converted to local time
            #>
            $vmhostPatch = $esxcli2.software.vib.list.Invoke() | Where-Object {$_.ID -match $vmhost.Build} | Select-Object -First 1
            if ($vmhost.ApiVersion -notmatch '6.5') {
                $lastPatched = Get-Date $vmhostPatch.InstallDate -Format d
            }
            else {
                Write-Verbose -Message ((Get-Date -Format G) + "`tESXi version " + $vmhost.ApiVersion + ". Gathering VIB " + $vmhostPatch.Name + " install date through ImageConfigManager" )
                $configManagerView = Get-View $vmhost.ExtensionData.ConfigManager.ImageConfigManager
                $softwarePackages = $configManagerView.fetchSoftwarePackages() | Where-Object {$_.CreationDate -ge $vmhostPatch.InstallDate}
                $dateInstalledUTC = ($softwarePackages | Where-Object {$_.Name -eq $vmhostPatch.Name -and $_.Version -eq $vmhostPatch.Version}).CreationDate
                $lastPatched = Get-Date ($dateInstalledUTC.ToLocalTime()) -Format d
            } #END if/else               
           
            <#
              Use a custom object to store
              collected data
            #>
            $PatchCollection += [PSCustomObject]@{
                'Hostname'             = $vmhost
                'Product'              = $vmhostView.Config.Product.Name
                'Version'              = $vmhostView.Config.Product.Version
                'Build'                = $vmhost.Build
                'Update'               = $esxiVersion.Update
                'Patch'                = $esxiVersion.Patch
                'Status'               = $complianceStatus
                'ESXi Microcode Patch' = $esxMicrocode
                'ESXi CPUID Features'  = (@($hostCPUID) -join ',')
                'Last Patched'         = $lastPatched
            } #END [PSCustomObject]
        } #END if
    
        <#
          Get ESXi BIOS details
        #>
        if ($BIOSCompliance) {
            Write-Host "`tGathering BIOS details from $vmhost ..."
            $BIOSReleaseDate = Get-Date (($vmhostview.Hardware.BiosInfo.ReleaseDate -split " ")[0])
            <#
              Get BIOS Compliance status
            #>
            $minVersion = $csvCollection | Where-Object {$_.Model -eq $vmhost.Model}

            if ($minVersion) {

                <#
                  if HP compare against Release Date
                  else compare against BIOS version
                #>
                if ($minVersion.Manufacturer -like "HP*") {
                    $csvDate = Get-Date $minVersion.BIOSReleaseDate
                    if ($BIOSReleaseDate -ge $csvDate) {
                        $BIOSComplianceStatus = "Compliant"
                    }
                    else {
                        $BIOSComplianceStatus = "NotCompliant - Update to version release: v" + $minVersion.BIOS + " " + $minVersion.BIOSReleaseDate
                    } #END if/else
                }
                else {
                    if ($vmhostview.Hardware.BiosInfo.BiosVersion -ge $minVersion.BIOS) {
                        $BIOSComplianceStatus = "Compliant"
                    }
                    else {
                        $BIOSComplianceStatus = "NotCompliant - Update to version: " + $minVersion.BIOS
                    } #END if/else
                } #END if/else
            }
            else {
                $BIOSComplianceStatus = "NotCompliant - Check with manufacturer"
            }
            
            <#
              Compare both ESXi and BIOS compliance Status
            #>
            if ($complianceStatus -eq "Compliant" -and $BIOSComplianceStatus -eq "Compliant" -and $esxMicrocode -eq "Present") {
                $spectreStatus = "True, Both BIOS and ESXi Microcode update are present"
            }
            elseif ($complianceStatus -eq "Compliant" -and $BIOSComplianceStatus -eq "Compliant" -and $esxMicrocode -like "NotPresent*") {
                $spectreStatus = "True, BIOS Update Present. ESXi Microcode update Notpresent"             
            }
            elseif ($complianceStatus -eq "Compliant" -and $BIOSComplianceStatus -like "NotCompliant*" -and $esxMicrocode -like "Present") {
                $spectreStatus = "True, BIOS Update NotPresent. ESXi Microcode update present"
            }
            else {
                $spectreStatus = $false
            } #END if/else

            <#
              Use a custom object to store
              collected data
            #>
            $BIOSCollection += [PSCustomObject]@{
                'Hostname'          = $vmhost
                'Make'              = $vmhost.Manufacturer
                'Model'             = $vmhost.Model
                'CPU Model'         = $vmhost.ProcessorType
                'Hyper-Threading'   = $vmhost.HyperthreadingActive
                'Max EVC Mode'      = $vmhost.MaxEVCMode
                'BIOS'              = $vmhostview.Hardware.BiosInfo.BiosVersion
                'BIOS Release Date' = (($vmhost.ExtensionData.Hardware.BiosInfo.ReleaseDate -split " ")[0])
                'Status'            = $BIOSComplianceStatus
                'SafeFromSpectre'   = $spectreStatus
            } #END [PSCustomObject]
        } #END if

        <#
          Report on VMs if switch was specified
        #>
        if ($ReportOnVMs) {
            $vmlist = $vmhost | Get-VM | Sort-Object -Property Name
            foreach ($vm in $vmlist) {
                Write-Host "`tGathering VM hypervisor-assisted guest mitigation details from $vm ..."
                $vmhost = Get-VMHost -Name $vm.VMHost
                $hardwareVersion = ($vm.Version.ToString()).Split('v')[1]
                $hostCPUID = $vmhost.ExtensionData.Config.FeatureCapability | Where-Object {$_.FeatureName -eq "cpuid.IBRS" -or $_.Featurename -eq "cpuid.IBPB" -or $_.Featurename -eq "cpuid.STIBP"} | Select-Object -ExpandProperty FeatureName
                $vmCPUID = $vm.ExtensionData.Runtime.FeatureRequirement | Where-Object {$_.FeatureName -eq "cpuid.IBRS" -or $_.Featurename -eq "cpuid.IBPB" -or $_.Featurename -eq "cpuid.STIBP"} | Select-Object -ExpandProperty FeatureName

                <#
                  Validate VM Hardware
                #>
                if ([int]$hardwareVersion -ge "9" -and $vm.PowerState -eq "PoweredOn") {
                    if ($hostCPUID -and $vmCPUID) {
                        $spectreStatus = $true
                    }
                    elseif ($hostCPUID -and !$vmCPUID) {
                        $spectreStatus = "False, You need to powercycle your VM"
                    } #END if/else
                }
                else {
                    if ([int]$hardwareVersion -ge "9" -and $vm.PowerState -eq "PoweredOff") {
                        $spectreStatus = "UnknownSinceOff"
                    }
                    else {
                        $spectreStatus = "Need to upgrade VM Hardware. See KB52085"
                    }
                } #END if

                <#
                  Use a custom object to store
                  collected data
                #>
                $VMCollection += [PSCustomObject]@{
                    'Name'                = $vm.Name
                    'Power State'         = $vm.PowerState
                    'Hardware Version'    = $vm.Version
                    'ESXi Host'           = $vmhost
                    'ESXi CPUID Features' = (@($hostCPUID) -join ',')
                    'VM CPUID Features'   = (@($vmCPUID) -join ',')
                    'SafeFromSpectre'     = $spectreStatus
                } #END [PSCustomObject]
            } #END foreach
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
    if ($PatchCollection -or $BIOSCollection -or $VMCollection) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tInformation gathered")
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tNo information gathered")
    } #END if/else
    
    <#
      Output to screen
      Export data to CSV, Excel
    #>
    if ($PatchCollection) {
        Write-Host "`n" "ESXi VMSA-2018-0004 Compliance:" -ForegroundColor Green
        if ($ExportCSV) {
            $PatchCollection | Export-Csv ($outputFile + "Patch.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "Patch.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $PatchCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Patch_Compliance -NoNumberConversion * -AutoSize -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $ReturnCollection += $PatchCollection 
            $ReturnCollection 
        }
        else {
            $PatchCollection | Format-List
        }#END if/else
    } #END if
    
    if ($BIOSCollection) {
        Write-Host "`n" "ESXi BIOS Compliance:" -ForegroundColor Green
        if ($ExportCSV) {
            $BIOSCollection | Export-Csv ($outputFile + "BIOS.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "BIOS.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $BIOSCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname BIOS_Compliance -NoNumberConversion * -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $ReturnCollection += $BIOSCollection
            $ReturnCollection  
        }
        else {
            $BIOSCollection | Format-List
        }#END if/else
    } #END if

    if ($VMCollection) {
        Write-Host "`n" "VM VMSA-2018-0004 Compliance:" -ForegroundColor Green
        if ($ExportCSV) {
            $VMCollection | Export-Csv ($outputFile + "VM.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "VM.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $VMCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname VM_Compliance -NoNumberConversion * -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $ReturnCollection += $VMCollection
            $ReturnCollection  
        }
        else {
            $VMCollection | Format-List
        }#END if/else
    } #END if
} #END function