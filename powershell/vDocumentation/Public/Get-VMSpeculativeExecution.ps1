function Get-VMSpeculativeExecution {
    <#
     .SYNOPSIS
       Get VM compliance on VMSA-2018-0004 Security Advisory
     .DESCRIPTION
       Will validate VM for Spectre/Hypervisor-Assisted Guest Mitigation
       https://www.vmware.com/security/advisories/VMSA-2018-0004.html
       https://www.vmware.com/security/advisories/VMSA-2018-0012.1.html
     .NOTES
       File Name    : Get-VMSpeculativeExecution.ps1
       Author       : Edgar Sanchez - @edmsanchez13
       Contributor  : Ariel Sanchez - @arielsanchezmor
       Version      : 2.4.4
     .Link
       https://github.com/arielsanchezmora/vDocumentation
     .PARAMETER VM
       Specifies the virtual machine(s) which you want to validate compliance on.
     .EXAMPLE
       C:\PS> Get-VM -Server $VCServer | Get-VMSpeculativeExecution

       Retrieves the compliance status on Spectre of all virtual machines which run in the $VCServer vCenter Server.
     
     .EXAMPLE
       C:\PS> Get-VM -Name "labvm001" | Get-VMSpeculativeExecution

       Name                                 : labvm001
       Power state                          : PoweredOn
       VM Guest OS                          : Microsoft Windows Server 2008 (32-bit)
       ESXi Host                            : aluvm014.emea.convergys.com
       Cluster EVC mode                     : Disabled
       Hardware version                     : v13
       ESXi MCU CPUID                       : IBPB,IBRS,SSBD,STIBP
       ESXi PCID/INVPCID                    : True
       VM MCU CPUID                         : STIBP,SSBD,IBRS,IBPB
       VM PCID/INVPCID                      : PCID,INVPCID
       Last PoweredOn                       : 7/31/2018 5:36:47 AM
       Hypervisor-Assisted Guest mitigation : Supported/Enabled
       PCID optimization                    : Supported/Enabled

     .EXAMPLE
       C:\PS> Get-VM -Location "MyClusterName" | Get-VMSpeculativeExecution

       Retrieves the compliance status on Spectre of all virtual machines which run in the "MyClusterName" cluster.
 
     .EXAMPLE
       C:\PS> Get-VMHost "MyESXiHostName" | Get-VM | Get-VMSpeculativeExecution

       Retrieves the compliance status on Spectre of all virtual machines which run on the "MyESXiHostName" ESXi host.
     
       .EXAMPLE
      C:\PS> Get-VMHost "MyESXiHostName" | Get-VM | Get-VMSpeculativeExecution | Export-Excel "VMValidation.xlsx" -WorkSheetname "VMresults"

       Retrieves the compliance status on Spectre of all virtual machines which run on the "MyESXiHostName" ESXi host and exports to Excel

     .NOTES
       This advanced function assumes that you are connected to at least one vCenter Server system.
    #> 
    
    <#
     ----------------------------------------------------------[Declarations]----------------------------------------------------------
    #>
  
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeLine = $true,
            ValueFromPipelinebyPropertyName = $True,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine[]] $VM
    )
     
    <#
      Main code execution
    #>
    Process {
        $vmCollection = @()
        $hostCpuidCollection = @()
        $hostList = $vm | Select-Object -ExpandProperty VMhost -Unique
        foreach ($esxhost in $hostList) {
            $hostCpuidCollection += [PSCustomObject]@{
                'Name'              = $esxhost.Name
                'FeatureCapability' = $esxhost.ExtensionData.Config.FeatureCapability
            } #END [PSCustomObject]
        } #END foreach

        foreach ($oneVM in $VM) {
            $vmhost = $vm.VMHost
            if ($vmhost.Parent.EVCMode) {
                $clusEvcMode = $vmhost.Parent.EVCMode
            }
            else {
                $clusEvcMode = "Disabled"
            } #END if/else
            $hostCpuPcid = $false
            $hostMcuCpuid = $null
            $vmMcuCpuid = $null
            $vmCpuPcid = $null
            $powerOnEvent = $null
            $hardwareVersion = ($oneVM.Version.ToString()).Split('v')[1]
            $hostFeatureCapability = $hostCpuidCollection | Where-Object {$_.Name -eq $oneVM.VMHost.Name} | Select-Object -ExpandProperty FeatureCapability
            $hostCpuid = $hostFeatureCapability | Where-Object {$_.FeatureName -eq "cpuid.IBRS" -and $_.Value -eq "1" -or $_.Featurename -eq "cpuid.IBPB" -and $_.Value -eq "1" -or $_.Featurename -eq "cpuid.STIBP" -and $_.Value -eq "1" -or $_.Featurename -eq "cpuid.SSBD" -and $_.Value -eq "1"} | Select-Object -ExpandProperty FeatureName
            $hostInvPcid = $hostFeatureCapability | Where-Object {$_.FeatureName -eq "cpuid.INVPCID" -and $_.Value -eq "1"} | Select-Object -ExpandProperty FeatureName
            $hostPcid = $hostFeatureCapability | Where-Object {$_.Featurename -eq "cpuid.PCID" -and $_.Value -eq "1"} | Select-Object -ExpandProperty FeatureName
            $vmFeatureRequirement = $oneVM.ExtensionData.Runtime.FeatureRequirement
            $vmCpuid = $vmFeatureRequirement | Where-Object {$_.FeatureName -eq "cpuid.IBRS" -or $_.Featurename -eq "cpuid.IBPB" -or $_.Featurename -eq "cpuid.STIBP" -or $_.Featurename -eq "cpuid.SSBD"} | Select-Object -ExpandProperty FeatureName
            $vmPcid = $vmFeatureRequirement | Where-Object {$_.FeatureName -eq "cpuid.INVPCID" -or $_.Featurename -eq "cpuid.PCID"} | Select-Object -ExpandProperty FeatureName
            if ($oneVM.Guest.OSFullName) {
                $vmGuestOs = $oneVM.Guest.OSFullName
            }
            else {
                $vmGuestOs = $oneVM.ExtensionData.Config.GuestFullName
            } #END if
            if ($hostInvPcid -and $hostPcid) {
                $hostCpuPcid = $true
            } #End if        

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
            if ($hostCpuid) {
                $hostMcuCpuid = (@($hostCpuid.split('.') | Where-Object {$_ -ne "cpuid"}) -join ',')
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
            $info = New-Object PSObject
            $info | Add-Member -type NoteProperty -Name 'Name' -Value $oneVM.Name
            $info | Add-Member -type NoteProperty -Name 'Power state' -Value $oneVM.PowerState
            $info | Add-Member -type NoteProperty -Name 'VM Guest OS' -Value $vmGuestOs
            $info | Add-Member -type NoteProperty -Name 'ESXi Host' -Value $vmhost.Name
            $info | Add-Member -type NoteProperty -Name 'Cluster EVC mode' -Value $clusEvcMode
            $info | Add-Member -type NoteProperty -Name 'Hardware version' -Value $oneVM.Version
            $info | Add-Member -type NoteProperty -Name 'ESXi MCU CPUID' -Value $hostMcuCPuid
            $info | Add-Member -type NoteProperty -Name 'ESXi PCID/INVPCID' -Value $hostCpuPcid
            $info | Add-Member -type NoteProperty -Name 'VM MCU CPUID' -Value $vmMcuCpuid
            $info | Add-Member -type NoteProperty -Name 'VM PCID/INVPCID' -Value $vmCpuPcid
            $info | Add-Member -type NoteProperty -Name 'Last PoweredOn' -Value $powerOnEvent
            $info | Add-Member -type NoteProperty -Name 'Hypervisor-Assisted Guest mitigation' -Value $spectreStatus
            $info | Add-Member -type NoteProperty -Name 'PCID optimization' -Value $pcidStatus

            $vmCollection += $info        
        } #END foreach        
        $vmCollection
    } #END process
} #END function