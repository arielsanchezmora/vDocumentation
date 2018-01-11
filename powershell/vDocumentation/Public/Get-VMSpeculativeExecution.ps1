function Get-VMSpeculativeExecution {
    <#
     .SYNOPSIS
       Get VM compliance on VMSA-2018-0004 Security Advisory
     .DESCRIPTION
       Will validate VM for VMware Security Advisory VMSA-2018-0004 Compliance (https://www.vmware.com/us/security/advisories/VMSA-2018-0004.html)
     .NOTES
       Author     : Edgar Sanchez - @edmsanchez13
       Contributor: Ariel Sanchez - @arielsanchezmor
     .Link
       https://github.com/arielsanchezmora/vDocumentation
     .PARAMETER VM
       Specifies the virtual machine(s) which you want to validate compliance on.
     .EXAMPLE
       C:\PS> Get-VM -Server $VCServer | Get-VMSpeculativeExecution

       Retrieves the compliance status on Spectre of all virtual machines which run in the $VCServer vCenter Server.
     
     .EXAMPLE
       C:\PS> Get-VM -Name "labvm001" | Get-VMSpeculativeExecution

       Name                : labvm001
       Power State         : PoweredOn
       Hardware Version    : v10
       ESXi Host           : labhost010
       ESXi CPUID Features : cpuid.IBPB,cpuid.IBRS,cpuid.STIBP
       VM CPUID Features   : cpuid.STIBP,cpuid.IBRS,cpuid.IBPB
       SafeFromSpectre     : True

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
        $VMCollection = @()
        foreach ($oneVM in $VM) {
            $vmhost = Get-VMHost -Name $oneVM.VMHost
            $hardwareVersion = ($oneVM.Version.ToString()).Split('v')[1]
            $hostCPUID = $vmhost.ExtensionData.Config.FeatureCapability | Where-Object {$_.FeatureName -eq "cpuid.IBRS" -or $_.Featurename -eq "cpuid.IBPB" -or $_.Featurename -eq "cpuid.STIBP"} | Select-Object -ExpandProperty FeatureName
            $vmCPUID = $oneVM.ExtensionData.Runtime.FeatureRequirement | Where-Object {$_.FeatureName -eq "cpuid.IBRS" -or $_.Featurename -eq "cpuid.IBPB" -or $_.Featurename -eq "cpuid.STIBP"} | Select-Object -ExpandProperty FeatureName

            <#
              Validate VM Hardware
            #>
            if ([int]$hardwareVersion -ge "9" -and $oneVM.PowerState -eq "PoweredOn") {
                if ($hostCPUID -and $vmCPUID) {
                    $spectreStatus = $true
                }
                elseif ($hostCPUID -and !$vmCPUID) {
                    $spectreStatus = "False, You need to powercycle your VM"
                } #END if/else
            }
            else {
                if ([int]$hardwareVersion -ge "9" -and $oneVM.PowerState -eq "PoweredOff") {
                    $spectreStatus = "UnknownSinceOff"
                }
                else {
                    $spectreStatus = "Need to upgrade VM Hardware. See KB52085"
                }
            } #END if

            $info = New-Object PSObject
            $info | Add-Member -type NoteProperty -Name 'Name' -Value $oneVM.Name
            $info | Add-Member -type NoteProperty -Name 'Power State' -Value $oneVM.PowerState
            $info | Add-Member -type NoteProperty -Name 'Hardware Version' -Value $oneVM.Version
            $info | Add-Member -type NoteProperty -Name 'ESXi Host' -Value $oneVM.VMHost
            $info | Add-Member -type NoteProperty -Name 'ESXi CPUID Features' -Value (@($hostCPUID) -join ',')
            $info | Add-Member -type NoteProperty -Name 'VM CPUID Features' -Value (@($vmCPUID) -join ',')
            $info | Add-Member -type NoteProperty -Name 'SafeFromSpectre' -Value $spectreStatus
            
            $VMCollection += $info        
        } #END foreach
        $VMCollection
    } #END process
} #END function

