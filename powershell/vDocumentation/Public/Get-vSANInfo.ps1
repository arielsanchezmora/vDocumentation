function Get-vSANInfo {
    <#
     .SYNOPSIS
       Get basic vSAN Cluster information
     .DESCRIPTION
       Will get inventory information for a vSAN Cluster
       The following is gathered:
       vSAN Cluster Name, Cluster Type, Disk Claim Mode, Dedupe & Compression Enabled, Stretched Cluster Enabled, 
       Oldest Disk Format Version, Total Disks, Total Disk Groups, vSAN Capacity GB 
     .NOTES
       File Name    : Get-vSANInfo.ps1
       Author       : Graham Barker - @VirtualG_UK
       Contributor  : Edgar Sanchez - @edmsanchez13
       Contributor  : Ariel Sanchez - @arielsanchezmor
       Version      : 2.4.5
     .Link
       https://github.com/arielsanchezmora/vDocumentation
     .INPUTS
       No inputs required
     .OUTPUTS
       CSV file
       Excel file
     .PARAMETER Cluster
       The name(s) of the vSphere Cluster(s)
     .EXAMPLE
       Get-vSANInfo -Cluster production
     .PARAMETER ExportCSV
       Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-vSANInfo -Cluster production -ExportCSV
     .PARAMETER ExportExcel
       Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
       ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
       File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-vSANInfo -Cluster production -ExportExcel
     .PARAMETER folderPath
       Specify an alternate folder path where the exported data should be saved.
     .EXAMPLE
       Get-vSANInfo -ExportExcel -folderPath C:\temp
     .PARAMETER PassThru
       Switch to return object to command line
     .EXAMPLE
       Get-vSANInfo -Cluster production
    #> 
    
    <#
     ----------------------------------------------------------[Declarations]----------------------------------------------------------
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false,
            ParameterSetName = "Cluster")]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String[]]$Cluster,
        [switch]$ExportCSV,
        [switch]$ExportExcel,
        [switch]$PassThru,
        $folderPath
    )
    
    $configurationCollection = [System.Collections.ArrayList]@()
    $skipCollection = @()
    $vSANClusterList = @()
    $returnCollection = @()
    $date = Get-Date -format s
    $date = $date -replace ":", "-"
    $outputFile = "vSAN-info" + $date

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
      needs to be vCenter Server 6.5+
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate connection to a vCenter server and that it's running at least v6.5.0")
    if ($Global:DefaultViServers.Count -gt 0 -and $Global:DefaultViServers.Version -ge "6.5.0") {
        $vCenterVersion = $Global:DefaultViServers.Version
        Write-Host "`tConnected to vCenter: $Global:DefaultViServers, v$vCenterVersion" -ForegroundColor Green
    }
    else {
        Write-Error -Message "You must be connected to a vCenter Server running at least v6.5.0 server before running this Cmdlet."
        break
    } #END if/else
    
    <#
      Validate parameter (-Cluster) and gather cluster list.
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate parameters used")
    if ([string]::IsNullOrWhiteSpace($cluster) ) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tA parameter (-Cluster) was not specified. Will gather all clusters")
        Write-Output "`tGathering all clusters from the following vCenter(s): " $Global:DefaultViServers
        $vSANClusterList = Get-Cluster | Sort-Object -Property Name
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tExecuting Cmdlet using cluster parameter")
        Write-Output "`tGathering cluster list..."
        foreach ($vClusterName in $cluster) {
            $tempList = Get-Cluster -Name $vClusterName.Trim() -ErrorAction SilentlyContinue
            if ([string]::IsNullOrWhiteSpace($tempList)) {
                Write-Warning -Message "`tCluster with name $vClusterName was not found in $Global:DefaultViServers"
            }
            else {
                $vSANClusterList += $tempList | Sort-Object -Property Name
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
    foreach ($vSAN in $vSANClusterList) {
  
        <#
          Skip if cluster is not vSAN enabled
        #>
        if ($vSAN.VsanEnabled -eq $true) {
            <#
              Do nothing - Cluster is vSAN Enabled
            #>
        }
        else {
            <#
              Use a custom object to keep track of skipped
              clusters and continue to the next foreach loop
            #>
            $skipCollection += [pscustomobject]@{
                'Cluster' = $vSAN.Name
                'Status'  = "Not vSAN Enabled"
            } #END [PSCustomObject]
            continue
        } #END if/else

        <#
          Get vSAN oldest system version
        #>
        Write-Output "`tGathering configuration details from vSAN Cluster: $vSAN ..."
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering claimed disks configuration...")
        $clusterMoRef = $vSAN.ExtensionData.MoRef
        $vSanClusterHealth = Get-VSANView -Id "VsanVcClusterHealthSystem-vsan-cluster-health-system"
        $vSanSystemVersions = $vSanClusterHealth.VsanVcClusterQueryVerifyHealthSystemVersions($clusterMoRef)
        $oldestvSanSystemVersion = $vSanSystemVersions.HostResults | Select-Object -ExpandProperty Version | Sort-Object -Descending | Select-Object -Last 1

        <#
          Get vSAN configuration details
        #>
        $vSAN = $vSAN | Get-VsanClusterConfiguration
        $vSanDiskGroups = Get-VsanDiskGroup -Cluster $vSAN.name
        $vSanDisks = Get-VsanDisk -vSANDiskGroup $vSanDiskGroups
        $numberDisks = $vSanDisks.Count
        $vSanView = $vSAN | Get-View

        <#
          Get vSAN oldest disk format version
        #>
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering disk format Configuration...")
        $oldestDiskFormatVersion = $vSanDisks.DiskFormatVersion | Sort-Object -Unique | Select-Object -First 1
                   
        <#
          Get number of disk groups
        #>
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering disk group configuration...")
        $numberDiskGroups = $vSanDiskGroups.Count
    
        <#
          Get vSAN cluster type
        #>
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering cluster type Configuration...")
        $magneticDiskCounter = ($vSanDisks | Where-Object {$_.IsSsd -eq $true}).Count
        if ($magneticDiskCounter -gt 0) {
            $clusterType = "Hybrid"
        }
        else {
            $clusterType = "Flash"
        } #END if/else

        <#
          Get disk claim mode
        #>
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering disk claim mode configuration...")
        $diskClaimMode = $vSAN.VsanDiskClaimMode

        <#
          Get deduplication & compression
        #>
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering deduplication & compression configuration...")
        $deduplicationCompression = $vSAN.SpaceEfficiencyEnabled

        <#
          Get stretched cluster
        #>
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering stretched cluster configuration...")
        $stretchedCluster = $vSAN.StretchedClusterEnabled
            
        <#
          Get vSAN Capacity
        #>
        $vSanDS = Get-View $vSanView.Datastore | Where-Object {$_.Summary.Type -eq 'vsan'}
        $vSanDsCapacityGB = [math]::round(($vSanDS.Summary.Capacity) / 1GB, 2)
        $vSanProvisionedGB = [math]::round(($vSanDS.Summary.Capacity - $vSanDS.Summary.FreeSpace + $vSanDS.Summary.Uncommitted) / 1GB, 2)
        $vSanDsFreeGB = [math]::round(($vSanDS.Summary.FreeSpace) / 1GB, 2)

        <#
          Get vSAN Storage Policy
        #>
        $vSanPolicy = (Get-SpbmStoragePolicy -Name $vSAN.StoragePolicy).AnyOfRuleSets.allofrules
        $vSanFailureToTolerate = $vSanPolicy | Where-Object {$_.Capability -like "VSAN.hostFailuresToTolerate"} | Select-Object -ExpandProperty Value
        $vSanStripeWidth = $vSanPolicy | Where-Object {$_.Capability -like "VSAN.stripeWidth"} | Select-Object -ExpandProperty Value

        <#
          Use a custom object to store collected data
        #>
        $output = [PSCustomObject]@{
            'vSAN Cluster Name'                   = $vSAN.Name
            'Effective Hosts'                     = $vSanView.Summary.NumEffectiveHosts
            'Oldest vSAN Version'                 = $oldestvSanSystemVersion
            'Oldest Disk Format'                  = $oldestDiskFormatVersion
            'Cluster Type'                        = $clusterType
            'Disk Claim Mode'                     = $diskClaimMode
            'Deduplication & Compression Enabled' = $deduplicationCompression
            'Stretched Cluster Enabled'           = $stretchedCluster
            'Host Failures To Tolerate'           = $vSanFailureToTolerate
            'vSAN Stripe Width'                   = $vSanStripeWidth
            'Total vSAN Claimed Disks'            = $numberDisks
            'Total Disk Groups'                   = $numberDiskGroups
            'Total Capacity (GB)'                 = $vSanDsCapacityGB
            'Provisioned Space (GB)'              = $vSanProvisionedGB
            'Free Space (GB)'                     = $vSanDsFreeGB
        } #END [PSCustomObject]
        [void]$configurationCollection.Add($output)
    } #END foreach
    $stopWatch.Stop()
    Write-Verbose -Message ((Get-Date -Format G) + "`tMain code execution completed")
    Write-Verbose -Message  ((Get-Date -Format G) + "`tScript Duration: " + $stopWatch.Elapsed.Duration())

    <#
      Display skipped clusters and their vSAN status
    #>
    If ($skipCollection) {
        Write-Output "`n"
        Write-Warning -Message "`tCheck vSAN configuration or cluster name"
        Write-Warning -Message "`tSkipped cluster(s):"
        $skipCollection | Format-Table -AutoSize
    } #END if

    <#
      Validate output arrays
    #>
    if ($configurationCollection) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tInformation gathered")
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tNo information gathered")
    } #END if/else
    
    <#
      Output to screen
      Export data to CSV, Excel
    #>   
    if ($configurationCollection) {
        Write-Host "`n" "vSAN Configuration:" -ForegroundColor Green
        if ($ExportCSV) {
            $configurationCollection | Export-Csv ($outputFile + "Configuration.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "Configuration.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $configurationCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname vSAN_Configuration -NoNumberConversion * -BoldTopRow
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