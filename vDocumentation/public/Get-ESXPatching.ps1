function Get-ESXPatching {
    <#
     .SYNOPSIS
       Get ESXi patch compliance
     .DESCRIPTION
       Will get patch compliance for a vSphere Cluster, Datacenter or individual ESXi host
     .NOTES
       Author: Edgar Sanchez - @edmsanchez13
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
       Get-ESXPatching -esxi devvm001.lab.local
     .PARAMETER cluster
       The name(s) of the vSphere Cluster(s)
     .EXAMPLE
       Get-ESXPatching -cluster production-cluster
     .PARAMETER datacenter
       The name(s) of the vSphere Virtual Datacenter(s).
     .EXAMPLE
       Get-ESXPatching -datacenter vDC001
       Get-ESXPatching -datacenter "all vdc" will gather all hosts in vCenter(s). This is the default if no Parameter (-VMhost, -Cluster, or -Datacenter) is specified. 
     .PARAMETER baseline
      The name(s) of VUM basline(s) to use. By Default 'Critical Host Patches*', 'Non-Critical Host Patches*' are used if this parameter is not specified.
     .EXAMPLE
      Get-ESXPatching -cluster production-cluster -baseline 'Custom baseline Host Patches'
     .PARAMETER ExportCSV
       Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-ESXPatching -cluster production-cluster -ExportCSV
     .PARAMETER ExportExcel
       Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
       ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
       File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .PARAMETER folderPath
       Specify an alternate folder path where the exported data should be saved.
     .EXAMPLE
       Get-ESXPatching -cluster production-cluster -ExportExcel -folderPath C:\temp
     .PARAMETER PassThru
       Switch to return object to command line
     .EXAMPLE
       Get-ESXPatching -esxi 192.168.1.100 -Configuration -PassThru
    #> 
    
    <#
     ----------------------------------------------------------[Declarations]----------------------------------------------------------
    #>
  
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseOutputTypeCorrectly")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidGlobalVars")] # for $global:DefaultVIServers
    param (
        $esxi,
        $cluster,
        $datacenter,
        $baseline,
        [switch]$ExportCSV,
        [switch]$ExportExcel,
        [switch]$Patching,
        [switch]$PassThru,
        $folderPath
    )
    
    $patchingCollection = @()
    $lastPatchingCollection = @()
    $notCompliantPatchCollection = @()
    $skipCollection = @()
    $vHostList = @()
    $ReturnCollection = @()
    $date = Get-Date -format s
    $date = $date -replace ":", "-"
    $outputFile = "ESXiPatching" + $date
    
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
      Validate if baseline parameter was specified (-baseline).
      By default 'Critical Host Patches*', 'Non-Critical Host Patches*'
      VUM baselines are used.
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate baseline parameter")
    if ([string]::IsNullOrWhiteSpace($baseline)) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tA Baseline parameter (-baseline) was not specified. Will default to 'Critical Host Patches*', 'Non-Critical Host Patches*'")
        $baseline = 'Critical Host Patches*', 'Non-Critical Host Patches*'
    } #END if

    $patchBaseline = Get-PatchBaseline -Name $baseline.Trim() -ErrorAction SilentlyContinue
    if ($patchBaseline) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tUsing VUM Baseline: " + (@($patchBaseline.Name) -join ','))
    }
    else {
        Write-Error -Message ("Could not find any baseline(s) named " + (@($baseline) -join ',') + " on server " + $Global:DefaultViServers + ". Please check Baseline name and try again.")
        break
    } #END if/else

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
                    <#
                      Start Scan for Updates
                    #>
                    Write-Verbose -Message ((Get-Date -Format G) + "`tStarting Scan for Updates ...")
                    $scanEntity = Get-Datacenter
                    $scanEntity | Add-EntityBaseline -Baseline $patchBaseline -ErrorAction SilentlyContinue
                    $testComplianceTask = $scanEntity | Test-Compliance -UpdateType HostPatch -RunAsync -ErrorAction SilentlyContinue
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
                            <#
                              Start Scan for Updates
                            #>
                            Write-Verbose -Message ((Get-Date -Format G) + "`tStarting Scan for Updates ...")
                            $scanEntity = Get-Datacenter -Name $vDCname.Trim() -ErrorAction SilentlyContinue
                            $scanEntity | Add-EntityBaseline -Baseline $patchBaseline -ErrorAction SilentlyContinue
                            $testComplianceTask = $scanEntity | Test-Compliance -UpdateType HostPatch -RunAsync -ErrorAction SilentlyContinue
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
                    <#
                      Start Scan for Updates
                    #>
                    Write-Verbose -Message ((Get-Date -Format G) + "`tStarting Scan for Updates ...")
                    $scanEntity = Get-Cluster -Name $vClusterName.Trim() -ErrorAction SilentlyContinue
                    $scanEntity | Add-EntityBaseline -Baseline $patchBaseline -ErrorAction SilentlyContinue
                    $testComplianceTask = $scanEntity | Test-Compliance -UpdateType HostPatch -RunAsync -ErrorAction SilentlyContinue
                } #END if/else
            } #END foreach
        } #END if/else
    }
    else { 
        Write-Verbose -Message ((Get-Date -Format G) + "`tExecuting Cmdlet using esxi parameter")
        Write-Output -InputObject "`tGathering host list..."
        foreach ($invidualHost in $esxi) {
            $vHostList += $invidualHost.Trim() | Sort-Object -Property Name
            <#
              Start Scan for Updates
            #>
            if (Get-VMHost -Name $invidualHost.Trim() -ErrorAction SilentlyContinue) {
                Write-Verbose -Message ((Get-Date -Format G) + "`tStarting Scan for Updates ...")
                $scanEntity = $invidualHost.Trim()
                $scanEntity | Add-EntityBaseline -Baseline $patchBaseline -ErrorAction SilentlyContinue
                $testComplianceTask = Test-Compliance -Entity $scanEntity -UpdateType HostPatch -RunAsync -ErrorAction SilentlyContinue
            }
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
      Main code execution
      Get patch compliance details 
    #>
    if ($testComplianceTask) {
        $testComplianceTask = Get-Task -id $testComplianceTask.id -ErrorAction SilentlyContinue
    }# END if
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
    
        <#
          Get ESXi version details
        #>
        $esxcli = Get-EsxCli -VMHost $esxihost -V2
        $vmhostView = $vmhost | Get-View
        $esxiVersion = $esxcli.system.version.get.Invoke()
        <#
          Get ESXi Patch Compliance
        #>
        Write-Output -InputObject "`tGathering patch compliance from $vmhost ..."
        <#
          Get ESXi Software configuration
          and patch compliance
        #>
        $vmhostPatch = $esxcli.software.vib.list.Invoke() | Where-Object {$_.ID -match $vmhost.Build} | Select-Object -First 1
        $installedPatches = $esxcli.software.vib.list.Invoke() | Where-Object {$_.InstallDate -eq $vmhostPatch.InstallDate -and $_.Vendor -like "VMware*"}
        while ($testComplianceTask.PercentComplete -ne 100) {
            Write-Output -InputObject "`tWaiting on scan for updates to complete... " $testComplianceTask.PercentComplete "%"
            Start-Sleep -seconds 5
            $testComplianceTask = Get-Task -id $testComplianceTask.id
        }
            
        $vmPatchCompliance = $VMhost | Get-Compliance -Baseline $patchBaseline -Detailed
        foreach ($vmbaseline in $vmPatchCompliance) {
            <#
             Use a custom object to store
             collected data
            #>
            $patchingCollection += [PSCustomObject]@{
                'Hostname'     = $vmhost
                'Product'      = $vmhostView.Config.Product.Name
                'Version'      = $vmhostView.Config.Product.Version
                'Build'        = $vmhost.Build
                'Update'       = $esxiVersion.Update
                'Patch'        = $esxiVersion.Patch
                'Baseline'     = $vmbaseline.Baseline.Name
                'Compliance'   = $vmbaseline.Status
                'Last Patched' = Get-Date $vmhostPatch.InstallDate -Format d
            } #END [PSCustomObject]
                
            <#
              Get last installed
              patches
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering last installed patches...")
            foreach ($vmPatch in $installedPatches) {
                $lastInstalledPatches = $vmbaseline.CompliantPatches | Where-Object {($_.Name.Replace(',', '')).Split() -contains $vmPatch.Name}
                foreach ($lastInstalledPatch in $lastInstalledPatches) {
                    $dateReleased = Get-Date $lastInstalledPatch.ReleaseDate -Format d
                    $dateInstalled = Get-Date $vmPatch.InstallDate -Format d
                    $patchTimespan = (New-TimeSpan -Start $dateReleased -End $dateInstalled).Days
                    <#
                      Use a custom object to store
                      collected data
                    #>
                    $lastPatchingCollection += [PSCustomObject]@{
                        'Hostname'       = $vmhost
                        'Product'        = $vmhostView.Config.Product.Name
                        'Version'        = $vmhostView.Config.Product.Version
                        'Build'          = $vmhost.Build
                        'Baseline'       = $vmbaseline.Baseline.Name
                        'VIB Name'       = $vmPatch.Name
                        'Patch Name'     = $lastInstalledPatch.Name
                        'Release Date'   = $dateReleased
                        'Installed Date' = $dateInstalled
                        'Patch Timespan' = "$patchTimespan Day(s)"
                        'Vendor ID'      = $lastInstalledPatch.IdByVendor
                        'URL'            = ($lastInstalledPatch.Description | Select-String "(?<url>http://[\w|\.|/]*\w{1})").Matches[0].Groups['url'].Value
                    } #END [PSCustomObject]
                } #END foreach
            } #END foreach
                
            <#
              Get not compliant
              patches
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering not compliant patches...")
            $notCompliantPatches = $vmbaseline.NotCompliantPatches
            foreach ($notCompliantPatch in $notCompliantPatches) {
                <#
                  Use a custom object to store
                  collected data
                #>
                $notCompliantPatchCollection += [PSCustomObject]@{
                    'Hostname'     = $vmhost
                    'Product'      = $vmhostView.Config.Product.Name
                    'Version'      = $vmhostView.Config.Product.Version
                    'Build'        = $vmhost.Build
                    'Baseline'     = $vmbaseline.Baseline.Name
                    'Patch Name'   = $notCompliantPatch.Name
                    'Release Date' = Get-Date $notCompliantPatch.ReleaseDate -Format d
                    'Vendor ID'    = $notCompliantPatch.IdByVendor
                    'URL'          = ($notCompliantPatch.Description | Select-String "(?<url>http://[\w|\.|/]*\w{1})").Matches[0].Groups['url'].Value
                } #END [PSCustomObject]
            } #END foreach
        } #END foreach
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
    if ($patchingCollection -or $lastPatchingCollection -or $notCompliantPatchCollection) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tInformation gathered")
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tNo information gathered")
    } #END if/else
    
    <#
      Output to screen
      Export data to CSV, Excel
    #>
    if ($patchingCollection) {
        Write-Output -InputObject "`n" "ESXi Patch Compliance:" -ForegroundColor Green
        if ($ExportCSV) {
            $patchingCollection | Export-Csv ($outputFile + "PatchCompliance.csv") -NoTypeInformation
            Write-Output -InputObject "`tData exported to" ($outputFile + "PatchCompliance.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $patchingCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Patch_Compliance -NoNumberConversion * -BoldTopRow
            Write-Output -InputObject "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $ReturnCollection += $patchingCollection
            $ReturnCollection  
        }
        else {
            $patchingCollection | Format-List
        }#END if/else
    } #END if

    if ($lastPatchingCollection) {
        Write-Output -InputObject "`n" "ESXi Last Installed Patches:" -ForegroundColor Green
        if ($ExportCSV) {
            $lastPatchingCollection | Export-Csv ($outputFile + "LastInstalledPatches.csv") -NoTypeInformation
            Write-Output -InputObject "`tData exported to" ($outputFile + "LastInstalledPatches.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $lastPatchingCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Last_Installed_Patches -NoNumberConversion * -BoldTopRow
            Write-Output -InputObject "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $ReturnCollection += $lastPatchingCollection
            $ReturnCollection  
        }
        else {
            $lastPatchingCollection | Format-List
        }#END if/else
    } #END if

    if ($notCompliantPatchCollection) {
        Write-Output -InputObject "`n" "ESXi Missing Patches:" -ForegroundColor Green
        if ($ExportCSV) {
            $notCompliantPatchCollection | Export-Csv ($outputFile + "MissingPatches.csv") -NoTypeInformation
            Write-Output -InputObject "`tData exported to" ($outputFile + "MissingPatches.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $notCompliantPatchCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Missing_Patches -NoNumberConversion * -BoldTopRow
            Write-Output -InputObject "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $ReturnCollection += $notCompliantPatchCollection
            $ReturnCollection  
        }
        else {
            $notCompliantPatchCollection | Format-List
        }#END if/else
    } #END if    
} #END function