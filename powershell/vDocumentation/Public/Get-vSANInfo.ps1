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
       Author     : Graham Barker - @VirtualG_UK
       Contributor: Edgar Sanchez - @edmsanchez13
       Contributor: Ariel Sanchez - @arielsanchezmor
     .Link
       https://github.com/arielsanchezmora/vDocumentation
     .INPUTS
       No inputs required
     .OUTPUTS
       CSV file
       Excel fil
     .PARAMETER cluster
       The name(s) of the vSphere Cluster(s)
     .EXAMPLE
       Get-vSANInfo -cluster production-cluster
     .PARAMETER ExportCSV
       Switch to export all data to CSV file. File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-vSANInfo -cluster production-cluster -ExportCSV
     .PARAMETER ExportExcel
       Switch to export all data to Excel file (No need to have Excel Installed). This relies on ImportExcel Module to be installed.
       ImportExcel Module can be installed directly from the PowerShell Gallery. See https://github.com/dfinke/ImportExcel for more information
       File is saved to the current user directory from where the script was executed. Use -folderPath parameter to specify a alternate location
     .EXAMPLE
       Get-vSANInfo -cluster production-cluster -ExportExcel
     .PARAMETER folderPath
       Specify an alternate folder path where the exported data should be saved.
     .EXAMPLE
       Get-vSANInfo -ExportExcel -folderPath C:\temp
     .PARAMETER PassThru
       Switch to return object to command line
     .EXAMPLE
       Get-vSANInfo -cluster production-cluster
    #> 
    
    <#
     ----------------------------------------------------------[Declarations]----------------------------------------------------------
    #>
  
    [CmdletBinding()]
    param (
        $cluster,
        [switch]$ExportCSV,
        [switch]$ExportExcel,
        [switch]$PassThru,
        $folderPath
    )
    
    #$hardwareCollection = @()
    $configurationCollection = @()
    $skipCollection = @()
    $vSANClusters = @()
    $returnCollection = @()
    $date = Get-Date -format s
    $date = $date -replace ":", "-"
    $dStore = 0;
    $outputFile = "vsan-info" + $date

    
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
      Validate parameter (-cluster) and gather cluster list.
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate parameters used")
    if ([string]::IsNullOrWhiteSpace($cluster) ) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tA parameter (-cluster) was not specified. Will gather all clusters")
        Write-Host "`tGathering all clusters from the following vCenter(s): " $Global:DefaultViServers
        $vSANClusters = Get-Cluster | Sort-Object -Property Name
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tExecuting Cmdlet using cluster parameter")
        Write-Host "`tGathering cluster list..."
        foreach ($vClusterName in $cluster) {
            $tempList = Get-Cluster -Name $vClusterName.Trim() -ErrorAction SilentlyContinue
            if ([string]::IsNullOrWhiteSpace($tempList)) {
                Write-Warning -Message "`tCluster with name $vClusterName was not found in $Global:DefaultViServers"
            }
            else {
                $vSANClusters += $tempList | Sort-Object -Property Name
            } #END if/else
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
    #>
    foreach ($vSANCluster in $vSANClusters) {
  
        <#
          Skip if cluster is not vSAN enabled
        #>
        if ($vSANCluster.VsanEnabled -eq $true) {
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
                'Cluster' = $vSANCluster.Name
                'Status'  = "Not vSAN Enabled"
            } #END [PSCustomObject]
            continue
        } #END if/else

        <#
          Get vSAN configuration details
        #>
        Write-Host "`tGathering configuration details from vSAN Cluster: $vSANCluster ..."
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering claimed disks configuration...")
        $vSAN = $vSANCluster | Get-VsanClusterConfiguration
        $vSanDiskGroups = Get-VsanDiskGroup -Cluster $vSAN.name
        $vSanDisks = Get-VsanDisk -vSANDiskGroup $vSanDiskGroups
        $numberDisks = $vSanDisks.Count

        <#
          Get disk format version configuration
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
          TODO: Must be an easier, more accurate & safer way to do this but cannot see anything in PowerCLI documentation
        #>

		# Begin Change
		# this does not work if you change the name of the vsan DS
		#$dStore = Get-Datastore -Name "vsanDatastore"
		# Changed to use the Datastore from the Cluster.
		# this may not work well if vsan and san used at the same time - have not been able to test a solution
		$dStore = get-cluster -name $vsan.name | get-datastore
		# end change
        $vSANCapacity = [math]::round(($dStore.CapacityGB), 2)
                    
        <#
          Use a custom object to store collected data
        #>
        $configurationCollection += [PSCustomObject]@{
            'vSAN Cluster Name'                   = $vSAN.Name
            'Cluster Type'                        = $clusterType
            'Disk Claim Mode'                     = $diskClaimMode
            'Deduplication & Compression Enabled' = $deduplicationCompression
            'Stretched Cluster Enabled'           = $stretchedCluster
            'Oldest Disk Format'                  = $oldestDiskFormatVersion
            'Total vSAN Claimed Disks'            = $numberDisks
            'Total Disk Groups'                   = $numberDiskGroups
            'Total Capacity (GB)'                 = $vSANCapacity
        } #END [PSCustomObject]
    } #END foreach
    Write-Verbose -Message ((Get-Date -Format G) + "`tMain code execution completed")

    <#
      Display skipped clusters and their vSAN status
    #>
    If ($skipCollection) {
        Write-Host "`n"
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