function Get-ESXSpeculativeExecution {
    <#
     .SYNOPSIS
       Get ESXi host mitigation status for Spectre
     .DESCRIPTION
       Will validate ESXi host for Spectre/Hypervisor-Assisted Guest Mitigation
       https://www.vmware.com/security/advisories/VMSA-2018-0004.html
     .NOTES
       File Name    : Get-ESXSpeculativeExecution.ps1
       Author       : Edgar Sanchez - @edmsanchez13
       Contributor  : Ariel Sanchez - @arielsanchezmor
       Version      : 2.4.3     
     .Link
       https://github.com/arielsanchezmora/vDocumentation
     .INPUTS
       If inputBiosFile is specified then an offline CSV file path input must be provided to check against BIOS version
       If inputMcuFile is specified then an offline CSV file path input must be provided to check against MCU revision
       If UseSSH is specified then SSH will be used to gather details around current MCU revisions. You will be prompted for username and password
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
     .PARAMETER inputBiosFile
       Specify input CSV file containing BIOS versions to validate against. Use this if you do not have access to the internet or wish to use your own offline version
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster -inputBiosFile "c:\temp\BIOSUpdates.csv"       
     .PARAMETER inputMcuFile
       Specify input CSV file containing Intel MCU version to validate against. Use this if you do not have access to the internet or wish to use your own offline version
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster -inputMcuFile "c:\temp\Intel_MCU.csv"
     .PARAMETER UseSSH
       Switch to use SSH to gahter specific MCU revisions details from the ESXi host. This needs to be specified if needed. This relies on Posh-SSH Module to be installed.
       Posh-SSH Module can be installed directly from the PowerShell Gallery. See https://github.com/darkoperator/Posh-SSH for more information
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster -ReportOnVMS       
     .PARAMETER ReportOnVMs
       Switch to report on virtual machines. This needs to be specified if needed.
     .EXAMPLE
       Get-ESXSpeculativeExecution -cluster production-cluster -ReportOnVMS       
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
        $inputBiosFile,
        $inputMcuFile,
        [switch]$UseSSH,
        [switch]$ReportOnVMs,
        [switch]$ExportCSV,
        [switch]$ExportExcel,
        [switch]$PassThru,
        $folderPath
    )
    
    $patchCollection = @()
    $skipCollection = @()
    $vmCollection = @()
    $vHostList = @()
    $returnCollection = @()
    $biosCsvCollection = @()
    $mcuCsvCollection = @()
    $biosUrl = 'https://raw.githubusercontent.com/edmsanchez/vDocumentation/master/powershell/vDocumentation/BIOSUpdates.csv'
    $mcuUrl = 'https://raw.githubusercontent.com/edmsanchez/vDocumentation/master/powershell/vDocumentation/Intel_MCU.csv'
    $date = Get-Date -format s
    $date = $date -replace ":", "-"
    $outputFile = "SpeculativeExecution" + $date
    <#
      VMSA-2018-0004.3 Build IDs
    #>
    $esx55Build = "7967571"
    $esx60Build = "7967664"
    $esx65Build = "7967591"
    
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
      Validate UseSSH switch and dependencies
    #>
    Write-Verbose -Message ((Get-Date -Format G) + "`tValidate UseSSH switch")
    if ($UseSSH) {
        if (Get-Module -ListAvailable -Name Posh-SSH) {
            Write-Verbose -Message ((Get-Date -Format G) + "`tPosh-SSH Module available")
            $poshSSH = $true
            $rootCreds = Get-Credential -UserName root -Message "Enter ESXi SSH Credentials"
        }
        else {
            Write-Warning -Message "`tPosh-SSH Module missing. Will not be able to retrieve MCU revision"
            Write-Warning -Message "`tPosh-SSH Module can be installed directly from the PowerShell Gallery"
            Write-Warning -Message "`tSee https://github.com/darkoperator/Posh-SSH for more information"
            $poshSSH = $false
        } #END if/else
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`t-UseSSH switch was not specified.")
        $poshSSH = $false
    } #END if/else

    <#
      Validate if inputBiosFile/inputMcuFile parameter was specified
      and if the CSV file(s) is reachable
    #>
    if ([string]::IsNullOrWhiteSpace($inputBiosFile)) {
        Write-Verbose -Message ((Get-Date -Format G) + "`t-inputBiosFile parameter is Null or Empty")
        Write-Verbose -Message ((Get-Date -Format G) + "`tValidating access to online CSV file")
        try {
            $webRequest = Invoke-WebRequest -Uri $biosUrl
        } 
        catch [System.Net.WebException] {
            $webRequest = $_.Exception.Response
        } #END try
        if ([int]$webRequest.StatusCode -eq "200") {
            $biosCsvCollection = $webRequest | ConvertFrom-Csv
        }
        else {
            Write-Warning -Message ("`tonline CSV file: '$biosUrl' is NOT reachable/unavailable (Return code: " + ([int]$webRequest.StatusCode) + ")")
            Write-Warning -Message "`tYou can download the file locally and use it offline with the -inputBiosFile parameter. See help for more information"
            Write-Warning -Message "`tSkipping BIOS Compliance check..."
        } #END if/else   
    }
    else {
        if (Test-Path $inputBiosFile) {
            Write-Verbose -Message ((Get-Date -Format G) + "`t'$inputBiosFile' path found")
            $biosCsvCollection = Import-Csv -Path $inputBiosFile
        }
        else {
            Write-Warning -Message "`t'$inputBiosFile' path not found."
            Write-Warning -Message "`tSkipping BIOS Compliance check..."
        } #END if/else
    } #END if/else
    if ([string]::IsNullOrWhiteSpace($inputMcuFile)) {
        Write-Verbose -Message ((Get-Date -Format G) + "`t-inputMcuFile parameter is Null or Empty")
        Write-Verbose -Message ((Get-Date -Format G) + "`tValidating access to online CSV file")
        try {
            $webRequest = Invoke-WebRequest -Uri $mcuUrl
        } 
        catch [System.Net.WebException] {
            $webRequest = $_.Exception.Response
        } #END try
        if ([int]$webRequest.StatusCode -eq "200") {
            $mcuCsvCollection = $webRequest | ConvertFrom-Csv
        }
        else {
            Write-Warning -Message ("`tonline CSV file: '$mcuUrl' is NOT reachable/unavailable (Return code: " + ([int]$webRequest.StatusCode) + ")")
            Write-Warning -Message "`tYou can download the file locally and use it offline with the -inputMcuFile parameter. See help for more information"
            Write-Warning -Message "`tSkipping Intel MCU Compliance check..."
        } #END if/else   
    }
    else {
        if (Test-Path $inputMcuFile) {
            Write-Verbose -Message ((Get-Date -Format G) + "`t'$inputMcuFile' path found")
            $mcuCsvCollection = Import-Csv -Path $inputMcuFile
        }
        else {
            Write-Warning -Message "`t'$inputMcuFile' path not found."
            Write-Warning -Message "`tSkipping Intel MCU Compliance check..."
        } #END if/else
    } #END if/else    
    
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
            $skipCollection += [PSCustomObject]@{
                'Hostname'         = $vmhost.Name
                'Connection State' = $vmhost.ConnectionState
            } #END [PSCustomObject]
            continue
        } #END if/else

        <#
          Get ESXi CPUID details
        #>
        Write-Host "`tGathering details from $vmhost ..."
        $esxcli = Get-EsxCli -VMHost $vmhost -V2
        $vmhostView = $vmhost | Get-View
        $mgmtIP = $vmhost | Get-VMHostNetworkAdapter | Where-Object {$_.ManagementTrafficEnabled -eq 'True'} | Select-Object -ExpandProperty IP
        $cpuList = $esxcli.hardware.cpu.list.Invoke() | Where-Object {$_.Id -eq "0"}
        $cpuModel = $vmhost.ProcessorType -replace '\s+', ' '
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering ESXi CPUID details...")
        $hostCpuPcid = $false
        $hostMcuCPuid = $null
        $hostFeatureCapability = $vmhostView.Config.FeatureCapability
        $hostCpuid = $hostFeatureCapability | Where-Object {$_.FeatureName -eq "cpuid.IBRS" -and $_.Value -eq "1" -or $_.Featurename -eq "cpuid.IBPB" -and $_.Value -eq "1" -or $_.Featurename -eq "cpuid.STIBP" -and $_.Value -eq "1"} | Select-Object -ExpandProperty FeatureName
        $hostInvPcid = $hostFeatureCapability | Where-Object {$_.FeatureName -eq "cpuid.INVPCID" -and $_.Value -eq "1"} | Select-Object -ExpandProperty FeatureName
        $hostPcid = $hostFeatureCapability | Where-Object {$_.Featurename -eq "cpuid.PCID" -and $_.Value -eq "1"} | Select-Object -ExpandProperty FeatureName
        if ($hostInvPcid -and $hostPcid) {
            $hostCpuPcid = $true
        } #End if        
        if ($hostCpuid) {
            $hostMcuCPuid = (@($hostCpuid.split('.') | Where-Object {$_ -ne "cpuid"}) -join ',')
        } #END if
        
        <#
          Get accurate last patched date if ESXi 6.5
          based on Date and time (UTC), which is
          converted to local time
        #>
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering last patched date...")
        $esxPatches = $esxcli.software.vib.list.Invoke()
        $vmhostPatch = $esxPatches | Where-Object {$_.ID -match $vmhost.Build} | Select-Object -First 1
        if ($vmhost.ApiVersion -notmatch '6.5') {
            $lastPatched = Get-Date $vmhostPatch.InstallDate -Format d
        }
        else {
            Write-Verbose -Message ((Get-Date -Format G) + "`tESXi version " + $vmhost.ApiVersion + ". Gathering VIB " + $vmhostPatch.Name + " install date through ImageConfigManager" )
            $configManagerView = Get-View $vmhostView.ConfigManager.ImageConfigManager
            $softwarePackages = $configManagerView.fetchSoftwarePackages() | Where-Object {$_.CreationDate -ge $vmhostPatch.InstallDate}
            $dateInstalledUTC = ($softwarePackages | Where-Object {$_.Name -eq $vmhostPatch.Name -and $_.Version -eq $vmhostPatch.Version}).CreationDate
            $lastPatched = Get-Date ($dateInstalledUTC.ToLocalTime()) -Format d
        } #END if/else
        
        <#
          Get ESXi VMSA-2018-0004.3
          path details
        #>
        $esxFrameworkStatus = $null
        $esxv2McuStatus = $null
        if ($vmhost.ApiVersion -eq "5.5") {
            $esxFrameworkStatus = "Framework Missing"
            $esxv2McuStatus = "v2 MCU Missing"
            if ($vmhost.Build -ge $esx55Build) {
                $esxFrameworkPatch = $esxPatches | Where-Object {$_.Name -eq "esx-base"}
                $esxv2McuPatch = $esxPatches | Where-Object {$_.Name -eq "cpu-microcode"}
                if ($esxFrameworkPatch) {
                    if (($esxFrameworkPatch.Version.Split('.') | Select-Object -Last 1) -ge $esx55Build) {
                        $esxFrameworkStatus = "Framework Installed"
                    } #END if
                } #END if
                if ($esxv2McuPatch) {
                    if (($esxv2McuPatch.Version.Split('.') | Select-Object -Last 1) -ge $esx55Build) {
                        $esxv2McuStatus = "v2 MCU Installed"
                    } #END if
                } #END if
            } #END if
        } #END if
        if ($vmhost.ApiVersion -eq "6.0") {
            $esxFrameworkStatus = "Framework Missing"
            $esxv2McuStatus = "v2 MCU Missing"
            if ($vmhost.Build -ge $esx60Build) {
                $esxFrameworkPatch = $esxPatches | Where-Object {$_.Name -eq "esx-base"}
                $esxv2McuPatch = $esxPatches | Where-Object {$_.Name -eq "cpu-microcode"}
                if ($esxFrameworkPatch) {
                    if (($esxFrameworkPatch.Version.Split('.') | Select-Object -Last 1) -ge $esx60Build) {
                        $esxFrameworkStatus = "Framework Installed"
                    } #END if
                } #END if
                if ($esxv2McuPatch) {
                    if (($esxv2McuPatch.Version.Split('.') | Select-Object -Last 1) -ge $esx60Build) {
                        $esxv2McuStatus = "v2 MCU Installed"
                    } #END if
                } #END if
            } #END if
        } #END if
        if ($vmhost.ApiVersion -eq "6.5") {
            $esxFrameworkStatus = "Framework Missing"
            $esxv2McuStatus = "v2 MCU Missing"
            if ($vmhost.Build -ge $esx65Build) {
                $esxFrameworkPatch = $esxPatches | Where-Object {$_.Name -eq "esx-base"}
                $esxv2McuPatch = $esxPatches | Where-Object {$_.Name -eq "cpu-microcode"}
                if ($esxFrameworkPatch) {
                    if (($esxFrameworkPatch.Version.Split('.') | Select-Object -Last 1) -ge $esx65Build) {
                        $esxFrameworkStatus = "Framework Installed"
                    } #END if
                } #END if
                if ($esxv2McuPatch) {
                    if (($esxv2McuPatch.Version.Split('.') | Select-Object -Last 1) -ge $esx65Build) {
                        $esxv2McuStatus = "v2 MCU Installed"
                    } #END if
                } #END if
            } #END if
        } #END if

        <#
          Get UpTime
        #>
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering ESXi UpTime details...")
        $bootTimeUTC = $vmhostView.Runtime.BootTime
        $bootTime = $bootTimeUTC.ToLocalTime()
        $upTime = New-TimeSpan -Seconds $vmhostView.Summary.QuickStats.Uptime
        $upTimeDays = $upTime.Days
        $upTimeHours = $upTime.Hours
        $upTimeMinutes = $upTime.Minutes

        <#
          Get ESXi MCU Revisions
        #>
        $mcuUpdate = $null
        $mcuOriginal = $null
        $mcuCurrent = $null
        $biosReleaseDate = $null
        if ($poshSSH) {
            Write-Verbose -Message ((Get-Date -Format G) + "`tGathering ESXi MCU revisions and BIOS release date through SSH...")
            $sshServerFWException = $vmhost | Get-VMHostFirewallException -Name "SSH Server"
            if ($sshServerFWException.Enabled -eq $false) {
                $sshServerFWException | Set-VMHostFirewallException -Enabled $true -Confirm:$false | Out-Null
            } #END if
            $vmServices = $vmhost | Get-VMHostService
            $vmServices | Where-Object {$_.Key -eq "TSM-SSH"} | Start-VMHostService -Confirm:$false | Out-Null
            $sshSession = New-SSHSession -ComputerName $vmhost -Credential $rootCreds -AcceptKey -ConnectionTimeout 90 -KeepAliveInterval 5 -ErrorAction SilentlyContinue
            if ($sshSession) {
                Write-Verbose -Message ((Get-Date -Format G) + "`tSSH session established...")
                $sshCommand = Invoke-SSHCommand -Command "vsish -e cat /hardware/cpu/cpuList/0 | grep microcode -A 2" -SessionId $sshSession.SessionId
                if ($sshCommand.ExitStatus -eq '0') {
                    $mcuUpdate = ($sshCommand.Output | Where-Object {$_ -match "Number of microcode"}).Split(':')[1]
                    $mcuOriginal = ($sshCommand.Output | Where-Object {$_ -match "Original Revision"}).Split(':')[1]
                    $mcuCurrent = ($sshCommand.Output | Where-Object {$_ -match "Current Revision"}).Split(':')[1]
                } #END if
                $sshCommand = Invoke-SSHCommand -Command "esxcfg-info --hardware | grep -i 'bios releasedate'" -SessionId $sshSession.SessionId
                if ($sshCommand.ExitStatus -eq '0') {
                    Write-Verbose -Message ((Get-Date -Format G) + $sshCommand.Output)
                    $intDate = @()
                    $stringDate = (($sshCommand.Output).Split('.') | Select-Object -Last 1 -ErrorAction SilentlyContinue).Split('T')[0]
                    foreach ($string in $stringDate.Split('-')) {
                        $intDate += [int]$string
                    } #END foreach
                    $biosReleaseDate = Get-Date (@($intDate) -join '/') -ErrorAction SilentlyContinue
                } #END if
                Remove-SSHSession $sshSession | Out-Null
            }
            else {
                Write-Warning -Message ("`tFailed to establish SSH connection: " + $Error[0].Exception.Message)
            } #END if/else
            $vmServices | Where-Object {$_.Key -eq "TSM-SSH"} | Stop-VMHostService -Confirm:$false | Out-Null
            if ($mcuUpdate -eq "0") {
                $mcuUpdate = "Inactive"
            }
            elseif ($mcuUpdate -eq "1") {
                $mcuUpdate = "Active"
            } #END if/elseif
        } #END if
        if ($biosReleaseDate) {
            Write-Verbose -Message ((Get-Date -Format G) + "`tBIOS release date was gathered through SSH...")
        }
        else {
            if ($vmhostView.Hardware.BiosInfo.ReleaseDate) {
                $biosReleaseDate = Get-Date $vmhostView.Hardware.BiosInfo.ReleaseDate
            }
            else {
                Write-Verbose -Message ((Get-Date -Format G) + "`tFailed to gather BIOS release date, its null...")
            } #END if/else
        } #END if/else
        
        <#
          Check for BIOS
          version details
        #>        
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering BIOS details...")
        if ($biosCsvCollection) {
            $minVersion = $biosCsvCollection | Where-Object {$_.Model -eq $vmhost.Model}
            if ($minVersion) {

                <#
                  If HP compare against Release Date
                  else compare against BIOS version
                #>
                if ($minVersion.Manufacturer -like "HP*") {
                    $biosCsvDate = [DateTime]$minVersion.BIOSReleaseDate
                    if ($biosReleaseDate -ge $biosCsvDate) {
                        $biosComplianceStatus = "Proper BIOS installed"
                    }
                    else {
                        $biosComplianceStatus = "BIOS update available - install v" + $minVersion.BIOS + " (" + $minVersion.BIOSReleaseDate + ")"
                    } #END if/else
                }
                else {
                    if ($vmhostView.Hardware.BiosInfo.BiosVersion -ge $minVersion.BIOS) {
                        $biosComplianceStatus = "Proper BIOS installed"
                    }
                    else {
                        $biosComplianceStatus = "BIOS update available - install v" + $minVersion.BIOS
                    } #END if/else
                } #END if/else
            }
            else {
                $biosComplianceStatus = "Unknown - Check with manufacturer"
            } #END if/else
        } #END if
        if ($biosReleaseDate) {
            $biosReleaseDate = $biosReleaseDate.ToShortDateString()
        } #END if
            
        <#
          Check for CPU Signatures
          Intel Sighting
        #>
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering Intel MCU and sighting details...")
        $cpuidEAX = ($esxcli.hardware.cpu.cpuid.get.Invoke(@{cpu = 0}) | Where-Object {$_.Level -eq 1}).EAX
        $cpuidHEX = [System.Convert]::ToString($cpuidEAX, 16)
        if ($cpuList.Family -eq "6") {
            if ($mcuCsvCollection) {
                $intelMcuList = $mcuCsvCollection | Where-Object {$_.CPUID -eq $cpuidHEX}
                if ($intelMcuList) {
                    $cpuAffected = $null
                    if ($intelMcuList.Count -gt 1) {
                        if ($mcu = $intelMcuList | Where-Object {$cpuModel.Contains($_.Name.Split('')[3])}) {
                            $intelProduct = $mcu.Product
                            $prodStatus = $mcu.Status
                            $affectedMCU = $mcu.AffectedMCU
                            $prodMCU = $mcu.ProductionMCU
                            if ($affectedMCU -eq "N/A") {
                                $cpuAffected = $false
                            }
                            else {
                                if ($mcuCurrent) {
                                    $cpuAffected = $affectedMCU.Contains($mcuCurrent)
                                }
                                else {
                                    $cpuAffected = "Unknown"        
                                } #END if/else
                            } #END if/else
                        }
                        else {
                            Write-Verbose -Message ((Get-Date -Format G) + "`tNo Match found for: $cpuModel")
                            $intelProduct = "Unknown"
                            $prodStatus = "Unknown"
                            $affectedMCU = "Unknown"
                            $prodMCU = "Unknown"
                            $cpuAffected = "Unknown"
                        } #END if/else
                    }
                    else {
                        $intelProduct = $intelMcuList.Product
                        $prodStatus = $intelMcuList.Status
                        $affectedMCU = $intelMcuList.AffectedMCU
                        $prodMCU = $intelMcuList.ProductionMCU
                        if ($affectedMCU -eq "N/A") {
                            $cpuAffected = $false
                        }
                        else {
                            if ($mcuCurrent) {
                                $cpuAffected = $affectedMCU.Contains($mcuCurrent)
                            }
                            else {
                                $cpuAffected = "Unknown"        
                            } #END if/else
                        } #END if/else
                    } #END if/else
                }
                else {
                    Write-Verbose -Message ((Get-Date -Format G) + "`t$cpuModel not found in Intel MCU Guidance list provided.")
                    $intelProduct = "Unknown"
                    $prodStatus = "Unknown"
                    $affectedMCU = "Unknown"
                    $prodMCU = "Unknown"
                    $cpuAffected = "Unknown"
                } #END if/else
            } #END if
        }
        else {
            Write-Verbose -Message ((Get-Date -Format G) + "`tNot an Intel CPU, skipping...")
            $intelProduct = "N/A"
            $prodStatus = "N/A"
            $affectedMCU = "N/A"
            $prodMCU = "N/A"
            $cpuAffected = "Unknown - Check with manufacturer"
        } #END if/Else

        <#
          Check for VMware
          MCU workaround
        #>
        Write-Verbose -Message ((Get-Date -Format G) + "`tGathering details for VMware MCU workaround...")
        $vmhostName = $vmhost.Name
        $url = "https://$vmhostName/host/vmware_config"
        $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)
        $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
        $spec.Method = "httpGet"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)
        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie = New-Object System.Net.Cookie
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhostName
        $websession.Cookies.Add($cookie)
        $result = Invoke-WebRequest -Uri $url -WebSession $websession
        $esxconfig = $result.content
        foreach ($line in $esxconfig.Split("`n")) {
            if ($line -eq 'cpuid.7.edx = "----:00--:----:----:----:----:----:----"') {
                $intelWorkaround = $true
                break
            }
            else {
                $intelWorkaround = $false
            } #END if/else
        } #END foreach

        <#
          Use a custom object to store
          collected data
        #>
        $patchCollection += [PSCustomObject]@{
            'Hostname'                      = $vmhost.Name
            'Management IP'                 = $mgmtIP
            'Cluster'                       = $vmhost.Parent            
            'Product'                       = $vmhostView.Config.Product.Name
            'Version'                       = $vmhostView.Config.Product.Version
            'Build'                         = $vmhost.Build            
            'Last patched'                  = $lastPatched
            'Boot time'                     = $bootTime
            'Uptime'                        = "$upTimeDays Day(s), $upTimeHours Hr(s), $upTimeMinutes Min(s)"            
            'Make'                          = $vmhost.Manufacturer
            'Model'                         = $vmhost.Model
            'CPU model'                     = $cpuModel
            'CPUID'                         = $cpuidHEX
            'Max EVC mode'                  = $vmhost.MaxEVCMode
            'BIOS version'                  = $vmhostView.Hardware.BiosInfo.BiosVersion
            'BIOS release date'             = $biosReleaseDate
            'BIOS guidance'                 = $biosComplianceStatus
            'ESXi guidance'                 = $esxFrameworkStatus + "/" + $esxv2McuStatus
            'MCU CPUID'                     = $hostMcuCPuid
            'PCID/INVPCID'                  = $hostCpuPcid
            'ESXi applied MCU'              = $mcuUpdate
            'MCU BIOS rev.'                 = $mcuOriginal
            'MCU boot rev.'                 = $mcuCurrent
            'Intel product'                 = $intelProduct
            'Intel MCU status'              = $prodStatus
            'Intel MCU(s) at risk'          = $affectedMCU
            'Intel production MCU'          = $prodMCU
            'MCU boot rev. at risk'         = $cpuAffected
            'VMware MCU workaround applied' = $intelWorkaround
        } #END [PSCustomObject]      

        <#
          Report on VMs if switch was specified
        #>
        if ($ReportOnVMs) {
            $vmlist = $vmhost | Get-VM | Sort-Object -Property Name
            if ($vmhost.Parent.EVCMode) {
                $clusEvcMode = $vmhost.Parent.EVCMode
            }
            else {
                $clusEvcMode = "Disabled"
            } #END if/else
            foreach ($vm in $vmlist) {
                $vmMcuCpuid = $null
                $vmCpuPcid = $null
                $powerOnEvent = $null
                Write-Host "`tGathering VM hypervisor-assisted guest mitigation details from $vm ..."
                $hardwareVersion = ($vm.Version.ToString()).Split('v')[1]              
                $vmFeatureRequirement = $vm.ExtensionData.Runtime.FeatureRequirement
                $vmCpuid = $vmFeatureRequirement | Where-Object {$_.FeatureName -eq "cpuid.IBRS" -or $_.Featurename -eq "cpuid.IBPB" -or $_.Featurename -eq "cpuid.STIBP"} | Select-Object -ExpandProperty FeatureName
                $vmPcid = $vmFeatureRequirement | Where-Object {$_.FeatureName -eq "cpuid.INVPCID" -or $_.Featurename -eq "cpuid.PCID"} | Select-Object -ExpandProperty FeatureName
                if ($vm.Guest.OSFullName) {
                    $vmGuestOs = $vm.Guest.OSFullName
                }
                else {
                    $vmGuestOs = $vm.ExtensionData.Config.GuestFullName
                } #END if

                <#
                  Validate VM Hardware
                #>
                Write-Verbose -Message ((Get-Date -Format G) + "`tValidating VM Hardware")
                if ([int]$hardwareVersion -ge "9" -and $vm.PowerState -eq "PoweredOn") {
                    if ($hostCpuid) {
                        if ($vmCpuid) {
                            $spectreStatus = "Supported/Enabled"
                        }
                        else {
                            $spectreStatus = "Supported/Disabled"
                        } #END if/else
                    }
                    else {
                        $spectreStatus = "NotSupported/Disabled"
                    } #END if/else
                    if ($hostCpuPcid) {
                        if ([int]$hardwareVersion -ge "11") {
                            if ($vmPcid.Count -eq "2") {
                                $pcidStatus = "Supported/Enabled"
                            }
                            else {
                                $pcidStatus = "Supported/Disabled"
                            } #END if/else
                        }
                        else {
                            $pcidStatus = "Supported/Upgrade VM Hardware"
                        } #END if/else
                    }
                    else {
                        $pcidStatus = "NotSupported/NA"
                    } #END if/else
                    $powerOnEvents = $vm | Get-VIEvent -MaxSamples ([int]::MaxValue) -Types Info | Where-Object {$_ -is [VMware.Vim.VmPoweredOnEvent]}
                    if ($powerOnEvents) {
                        $sortedEvents = Sort-Object -InputObject $powerOnEvents -Property CreatedTime -Descending
                        $powerOnEvent = ($sortedEvents | Select-Object -First 1).CreatedTime
                    } #END if
                }
                else {
                    if ([int]$hardwareVersion -ge "9" -and $vm.PowerState -eq "PoweredOff") {
                        $spectreStatus = "UnknownSinceOff"
                        $pcidStatus = "UnknownSinceOff"
                    }
                    else {
                        $spectreStatus = "Supported/Upgrade VM Hardware"
                        if ($hostCpuPcid) {
                            $pcidStatus = "Supported/Upgrade VM Hardware"
                        }
                        else {
                            $pcidStatus = "NotSupported/NA"
                        } #END if
                    } #END if/else
                } #END if/else
                if ($vmCpuid) {
                    $vmMcuCpuid = (@($vmCpuid.split('.') | Where-Object {$_ -ne "cpuid"}) -join ',')
                } #END if/else
                if ($vmPcid) {
                    $vmCpuPcid = (@($vmPcid.split('.') | Where-Object {$_ -ne "cpuid"}) -join ',')
                } #END if/else

                <#
                  Use a custom object to store
                  collected data
                #>
                $vmCollection += [PSCustomObject]@{
                    'Name'                                 = $vm.Name
                    'Power state'                          = $vm.PowerState
                    'VM guest OS'                          = $vmGuestOs
                    'ESXi'                                 = $vmhost.Name
                    'Cluster'                              = $vmhost.Parent
                    'Cluster EVC mode'                     = $clusEvcMode
                    'Hardware version'                     = $vm.Version
                    'ESXi MCU CPUID'                       = $hostMcuCPuid
                    'ESXi PCID/INVPCID'                    = $hostCpuPcid
                    'VM MCU CPUID'                         = $vmMcuCpuid
                    'VM PCID/INVPCID'                      = $vmCpuPcid
                    'Last PoweredOn'                       = $powerOnEvent
                    'Hypervisor-Assisted Guest mitigation' = $spectreStatus
                    'PCID optimization'                    = $pcidStatus
                } #END [PSCustomObject]
            } #END foreach
        } #END if
    } #END foreach
    Write-Verbose -Message ((Get-Date -Format G) + "`tMain code execution completed")
    
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
    if ($patchCollection -or $vmCollection) {
        Write-Verbose -Message ((Get-Date -Format G) + "`tInformation gathered")
    }
    else {
        Write-Verbose -Message ((Get-Date -Format G) + "`tNo information gathered")
    } #END if/else
    
    <#
      Output to screen
      Export data to CSV, Excel
    #>
    if ($patchCollection) {
        Write-Host "`n" "ESXi Speculative Execution:" -ForegroundColor Green
        if ($ExportCSV) {
            $patchCollection | Export-Csv ($outputFile + "ESXi.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "ESXi.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $patchCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname ESXi -NoNumberConversion * -AutoSize -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $returnCollection += $patchCollection 
            $returnCollection 
        }
        else {
            $patchCollection | Format-List
        }#END if/else
    } #END if

    if ($vmCollection) {
        Write-Host "`n" "VM Hypervisor-Assisted Guest Mitigation:" -ForegroundColor Green
        if ($ExportCSV) {
            $VMCollection | Export-Csv ($outputFile + "VM.csv") -NoTypeInformation
            Write-Host "`tData exported to" ($outputFile + "VM.csv") "file" -ForegroundColor Green
        }
        elseif ($ExportExcel) {
            $VMCollection | Export-Excel ($outputFile + ".xlsx") -WorkSheetname VM -NoNumberConversion * -BoldTopRow
            Write-Host "`tData exported to" ($outputFile + ".xlsx") "file" -ForegroundColor Green
        }
        elseif ($PassThru) {
            $returnCollection += $VMCollection
            $returnCollection  
        }
        else {
            $vmCollection | Format-List
        }#END if/else
    } #END if
} #END function