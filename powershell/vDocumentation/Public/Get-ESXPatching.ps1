function Get-ESXPatching {
    <#
     .SYNOPSIS
       Get ESXi patch compliance
     .DESCRIPTION
       Will get patch compliance for a vSphere Cluster, Datacenter or individual ESXi host
     .NOTES
       File Name    : Get-ESXPatching.ps1
       Author       : Edgar Sanchez - @edmsanchez13
       Contributor  : Ariel Sanchez - @arielsanchezmor
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
       Get-ESXPatching -VMhost devvm001.lab.local
     .PARAMETER Cluster
       The name(s) of the vSphere Cluster(s)
     .EXAMPLE
       Get-ESXPatching -Cluster production
     .PARAMETER Datacenter
       The name(s) of the vSphere Virtual Datacenter(s).
     .EXAMPLE
       Get-ESXPatching -Datacenter vDC001
     .PARAMETER baseline
      The name(s) of VUM basline(s) to use. By Default 'Critical Host Patches*', 'Non-Critical Host Patches*' are used if this parameter is not specified.
     .EXAMPLE
      Get-ESXPatching -Cluster production -baseline 'Custom baseline Host Patches'
     .PARAMETER ExportCSV
       Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-ESXPatching -Cluster production -ExportCSV
     .PARAMETER ExportExcel
       Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
       ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
       File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-ESXPatching -Cluster production -ExportExcel       
     .PARAMETER folderPath
       Specify an alternate folder path where the exported data should be saved.
     .EXAMPLE
       Get-ESXPatching -Cluster production -ExportExcel -folderPath C:\temp
     .PARAMETER PassThru
       Switch to return object to command line
     .EXAMPLE
       Get-ESXPatching -VMhost 192.168.1.100 -PassThru
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
        $baseline,
        [switch]$ExportCSV,
        [switch]$ExportExcel,
        [switch]$PassThru,
        $folderPath
    )
    
    $patchingCollection = [System.Collections.ArrayList]@()
    $lastPatchingCollection = [System.Collections.ArrayList]@()
    $notCompliantPatchCollection = [System.Collections.ArrayList]@()
    $skipCollection = @()
    $vHostList = @()
    $ReturnCollection = @()
    $date = Get-Date -format s
    $date = $date -replace ":", "-"
    $outputFile = "ESXiPatching" + $date
    
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

                <#
                  Start Scan for Updates
                #>
                Write-Verbose -Message ((Get-Date -Format G) + "`tStarting Scan for Updates ...")
                $scanEntity = $tempList
                $scanEntity | Attach-Baseline -Baseline $patchBaseline -ErrorAction SilentlyContinue
                $testComplianceTask = Test-Compliance -Entity $scanEntity -UpdateType HostPatch -RunAsync -ErrorAction SilentlyContinue
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

                <#
                  Start Scan for Updates
                #>
                Write-Verbose -Message ((Get-Date -Format G) + "`tStarting Scan for Updates ...")
                $scanEntity = Get-Cluster -Name $vClusterName.Trim() -ErrorAction SilentlyContinue
                $scanEntity | Attach-Baseline -Baseline $patchBaseline -ErrorAction SilentlyContinue
                $testComplianceTask = $scanEntity | Test-Compliance -UpdateType HostPatch -RunAsync -ErrorAction SilentlyContinue
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

                <#
                  Start Scan for Updates
                #>
                Write-Verbose -Message ((Get-Date -Format G) + "`tStarting Scan for Updates ...")
                $scanEntity = Get-Datacenter -Name $vDCname.Trim() -ErrorAction SilentlyContinue
                $scanEntity | Attach-Baseline -Baseline $patchBaseline -ErrorAction SilentlyContinue
                $testComplianceTask = $scanEntity | Test-Compliance -UpdateType HostPatch -RunAsync -ErrorAction SilentlyContinue
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
      Main code execution
      Get patch compliance details 
    #>
    if ($testComplianceTask) {
        $testComplianceTask = Get-Task -id $testComplianceTask.id -ErrorAction SilentlyContinue
    }# END if
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
    
        <#
          Get ESXi version details
        #>
        $esxcli = Get-EsxCli -VMHost $esxiHost -V2
        $vmhostView = $esxiHost | Get-View
        $esxiVersion = $esxcli.system.version.get.Invoke()

        <#
          Get ESXi Patch Compliance
          and details of sample/$vmhostPatch
        #>
        Write-Output "`tGathering patch compliance from $esxiHost ..."
        $vmhostPatch = $esxcli.software.vib.list.Invoke() | Where-Object {$_.ID -match $esxiHost.Build} | Select-Object -First 1
        $installedPatches = $esxcli.software.vib.list.Invoke() | Where-Object {$_.InstallDate -ge $vmhostPatch.InstallDate -and $_.Vendor -like "VMware*"}
        while ($testComplianceTask.PercentComplete -ne 100) {
            Write-Output ("`tWaiting on scan for updates to complete... " + $testComplianceTask.PercentComplete + "%")
            Start-Sleep -Seconds 5
            $testComplianceTask = Get-Task -id $testComplianceTask.id
        } #END while
            
        $vmPatchCompliance = $esxiHost | Get-Compliance -Baseline $patchBaseline -Detailed
        foreach ($vmbaseline in $vmPatchCompliance) {
            $samplePatch = $vmbaseline.CompliantPatches | Where-Object {($_.Name.Replace(',', '')).Split() -contains $vmhostPatch.Name}
            if ($samplePatch) {
                $patchProduct = $samplePatch.Product.Name
                $patchReleaseDate = $samplePatch.ReleaseDate
            } #END if

            <#
              Get accurate last patched date if ESXi 6.5
              based on Date and time (UTC), which is
              converted to local time
            #>
            if ($esxiHost.ApiVersion -notmatch '6.5') {
                $lastPatched = Get-Date $vmhostPatch.InstallDate -Format d
            }
            else {
                Write-Verbose -Message ((Get-Date -Format G) + "`tESXi version " + $esxiHost.ApiVersion + ". Gathering VIB " + $vmhostPatch.Name + " install date through ImageConfigManager" )
                $configManagerView = Get-View $vmhostView.ConfigManager.ImageConfigManager
                $softwarePackages = $configManagerView.fetchSoftwarePackages() | Where-Object {$_.CreationDate -ge $vmhostPatch.InstallDate}
                $dateInstalledUTC = ($softwarePackages | Where-Object {$_.Name -eq $vmhostPatch.Name -and $_.Version -eq $vmhostPatch.Version}).CreationDate
                $lastPatched = Get-Date ($dateInstalledUTC.ToLocalTime()) -Format d
            } #END if/else               

            <#
              Use a custom object to store
              collected data
            #>
            $output = [PSCustomObject]@{
                'Hostname'     = $esxiHost.Name
                'Product'      = $vmhostView.Config.Product.Name
                'Version'      = $vmhostView.Config.Product.Version
                'Build'        = $esxiHost.Build
                'Update'       = $esxiVersion.Update
                'Patch'        = $esxiVersion.Patch
                'Baseline'     = $vmbaseline.Baseline.Name
                'Compliance'   = $vmbaseline.Status
                'Last Patched' = $lastPatched
            } #END [PSCustomObject]
            [void]$patchingCollection.Add($output)
        } #END foreach

        <#
          Get last installed patches
        #>
        foreach ($vmbaseline in $vmPatchCompliance) {               
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering last installed patches...")
            $baselinePatches = Get-Patch -Baseline $vmbaseline.Baseline -Product $patchProduct -After $patchReleaseDate
            foreach ($vmPatch in $installedPatches) {
                $lastInstalledPatches = $baselinePatches | Where-Object {($_.Name.Replace(',', '')).Split() -contains $vmPatch.Name -and $_.ReleaseDate -eq $patchReleaseDate}
                foreach ($lastInstalledPatch in $lastInstalledPatches) {
                    <#
                      Determine if patch contains multiple VIBs
                      and update custom object so that
                      patch reports are accurate by Vendor ID
                    #>
                    $duplicateVendorID = $lastPatchingCollection | Where-Object {$_.Hostname -eq $esxiHost -and $_.'Vendor ID' -eq $lastInstalledPatch.IdByVendor}
                    if ($duplicateVendorID) {
                        if ($duplicateVendorID.'Patch Name' -eq $lastInstalledPatch.Name) {
                            Write-Verbose -Message ((Get-Date -Format G) + "`t" + $duplicateVendorID.'Vendor ID' + " already present in custom object. Updating VIB Name property with " + $vmPatch.Name)
                            $index = $lastPatchingCollection.IndexOf($duplicateVendorID)
                            $lastPatchingCollection[$index].'VIB Name(s)' += ", " + $vmPatch.Name
                            continue
                        } #END if
                    } #END if

                    <#
                      Get accurate patch install date if ESXi 6.5
                      based on Date and time (UTC), which is
                      converted to loal time
                    #>
                    if ($esxiHost.ApiVersion -notmatch '6.5') {
                        $dateInstalled = Get-Date $vmPatch.InstallDate -Format d
                    }
                    else {
                        Write-Verbose -Message ((Get-Date -Format G) + "`tESXi version " + $esxiHost.ApiVersion + ". Gathering VIB " + $vmPatch.Name + " install date through ImageConfigManager" )
                        $configManagerView = Get-View $vmhostView.ConfigManager.ImageConfigManager
                        $softwarePackages = $configManagerView.fetchSoftwarePackages() | Where-Object {$_.CreationDate -ge $vmPatch.InstallDate}
                        $dateInstalledUTC = ($softwarePackages | Where-Object {$_.Name -eq $vmPatch.Name -and $_.Version -eq $vmPatch.Version}).CreationDate
                        $dateInstalled = Get-Date ($dateInstalledUTC.ToLocalTime()) -Format d
                    } #END if/else

                    $dateReleased = Get-Date $lastInstalledPatch.ReleaseDate -Format d
                    $patchTimespan = (New-TimeSpan -Start $dateReleased -End $dateInstalled).Days
                    if ($lastInstalledPatch.Description -match 'http://' -or $lastInstalledPatch.Description -match 'https://') {
                        $referenceURL = ($lastInstalledPatch.Description | Select-String -Pattern "(?<url>https?://[\w|\.|/]*\w{1})").Matches[0].Groups['url'].Value
                    }
                    else {
                        Write-Verbose -Message ((Get-Date -Format G) + "`tFailed to get reference URL for patch: " + $lastInstalledPatch.Name)
                        Write-Verbose -Message ((Get-Date -Format G) + "`t" + $lastInstalledPatch.Description)
                        $referenceURL = $null
                    } #END if/else

                    <#
                      Use a custom object to store
                      collected data
                    #>
                    $output = [PSCustomObject]@{
                        'Hostname'       = $esxiHost.Name
                        'Product'        = $vmhostView.Config.Product.Name
                        'Version'        = $vmhostView.Config.Product.Version
                        'Build'          = $esxiHost.Build
                        'Baseline'       = $vmbaseline.Baseline.Name
                        'VIB Name(s)'    = $vmPatch.Name
                        'Patch Name'     = $lastInstalledPatch.Name
                        'Release Date'   = $dateReleased
                        'Installed Date' = $dateInstalled
                        'Patch Timespan' = "$patchTimespan Day(s)"
                        'Vendor ID'      = $lastInstalledPatch.IdByVendor
                        'URL'            = $referenceURL
                    } #END [PSCustomObject]
                    [void]$lastPatchingCollection.Add($output)
                } #END foreach
            } #END foreach
                
            <#
              Get not compliant
              patches
            #>
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering not compliant patches...")
            $notCompliantPatches = $vmbaseline.NotCompliantPatches
            foreach ($notCompliantPatch in $notCompliantPatches) {
                if ($notCompliantPatch.Description -match 'http://' -or $notCompliantPatch.Description -match 'https://') {
                    $referenceURL = ($notCompliantPatch.Description | Select-String -Pattern "(?<url>https?://[\w|\.|/]*\w{1})").Matches[0].Groups['url'].Value
                }
                else {
                    Write-Verbose -Message ((Get-Date -Format G) + "`tFailed to get reference URL for patch: " + $notCompliantPatch.Name)
                    Write-Verbose -Message ((Get-Date -Format G) + "`t" + $notCompliantPatch.Description)
                    $referenceURL = $null
                } #END if/else

                <#
                  Use a custom object to store
                  collected data
                #>
                $output = [PSCustomObject]@{
                    'Hostname'     = $esxiHost.Name
                    'Product'      = $vmhostView.Config.Product.Name
                    'Version'      = $vmhostView.Config.Product.Version
                    'Build'        = $esxiHost.Build
                    'Baseline'     = $vmbaseline.Baseline.Name
                    'Patch Name'   = $notCompliantPatch.Name
                    'Release Date' = Get-Date $notCompliantPatch.ReleaseDate -Format d
                    'Vendor ID'    = $notCompliantPatch.IdByVendor
                    'URL'          = $referenceURL
                } #END [PSCustomObject]
                [void]$notCompliantPatchCollection.Add($output)
            } #END foreach
        } #END foreach
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
        Write-Host "`n" "ESXi Patch Compliance:" -ForegroundColor Green
        if ($ExportCSV) {
            $patchingCollection | Export-Csv ($outputFile + "PatchCompliance.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "PatchCompliance.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $patchingCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Patch_Compliance -NoNumberConversion * -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
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
        Write-Host "`n" "ESXi Last Installed Patches:" -ForegroundColor Green
        if ($ExportCSV) {
            $lastPatchingCollection | Export-Csv ($outputFile + "LastInstalledPatches.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "LastInstalledPatches.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $lastPatchingCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Last_Installed_Patches -NoNumberConversion * -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
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
        Write-Host "`n" "ESXi Missing Patches:" -ForegroundColor Green
        if ($ExportCSV) {
            $notCompliantPatchCollection | Export-Csv ($outputFile + "MissingPatches.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "MissingPatches.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $notCompliantPatchCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname Missing_Patches -NoNumberConversion * -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
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