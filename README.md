# vDocumentation

vDocumentation provides a community-created set of PowerCLI scripts that produce documentation of vSphere environments in CSV or Excel file format.

# TL;DR

## First time usage on a *brand new machine*

_Paste in a PowerShell window that has been Run as Administrator and answer Y_

**Set-ExecutionPolicy RemoteSigned**  
**Set-PowerCLIConfiguration -InvalidCertificateAction Ignore**

![Run PowerShell as Administrator](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/PowerShell_as_administrator.png)

![Enable remote scripts and ignore certificate warnings](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/enable_RemoteSigned_Invalid_Certificate.png)

_You can now close the PowerShell window that ran as Administrator_ 

_In a new **normal** PowerShell console paste all of the below answering Y (this only affects your user, and it may take a while)_

**Install-Module -Name VMware.PowerCLI -Scope CurrentUser**  
**Install-Module ImportExcel -scope CurrentUser**  
**Install-Module vDocumentation -Scope CurrentUser**  

![Install PowerCLI, ImportExcel and vDocumentation modules](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/install_PowerCLI_ImportExcel_vDocumentation.png)

_vDocumentation is now installed! You can verify with_

**Get-Module vDocumentation -ListAvailable | Format-Table -AutoSize**

![Confirm vDocumentation installation](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/Confirm_vDocumentation_installation.png)


## The vDocumentation module gives you new PowerCLI Commands you can use to create documentation of a vSphere environment

_Make sure to connect to vCenter using PowerCLI_

**Connect-VIServer [IP_or_FQDN_of_vCenter]**  # Connect to one or many vCenters

_You are now able to use these commands:_

|Command|Description|
|----------------|---|
|**Get-ESXInventory**|Document host inventory and host config info|
|**Get-ESXIODevice**|Document information from HBAs, MACs, NICs and other PCIe devices including firmware & drivers|
|**Get-ESXNetworking**|Document networking configuration info such as NICs, vSwitches, VMKernel interface configuration|
|**Get-ESXStorage**|Document storage information and configurations such as iSCSI, FibreChannel, Datastores & Multipathing|

_Each script will output the corresponding data to terminal, and optionally create a file (XLSX, CSV) with the command name and a timestamp. You can use command switches to customize CSV or Excel output, file path (default is powershell working directory), and the command scope (report all vCenter or just cluster/host)._

## Upgrading from a previous version

_If the prompt returns without doing anything, you are running latest._

**Update-Module VMware.PowerCLI**  
**Update-Module ImportExcel**  
**Update-Module vDocumentation**

![Upgrade Commands](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/upgrade_commands.png)

## Uninstalling the vDocumentation script

**Uninstall-Module vDocumentation**

![Uninstall vDocumentation](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/uninstall_vDocumentation.png)

# Module Changelog

__v1.04__ new functionality added:  
 Updated export-excel so that it does no number conversion (IP addresses are now text) on any of the columns and it auto sizes them. Thanks to [@magneet_nl](https://twitter.com/Magneet_nl) for helping us discover this bug!

__v1.03__ new functionality added:  
 Get-ESXInventory: Added RAC Firmware version, BIOS release date.  
 Get-ESXIODevice: Added support to get HP Smart Array Firmware from PowerCLI  
 
__1.02__ Formatting & Manifest changes

__1.01__ Changes to support displaying datastore multipathing

__1.0__ First release to PowerShell Gallery with 4 commands: Get-ESXInventory, Get-ESXIODevice, Get-ESXNetworking & Get-ESXStorage


# vDocumentation backstory

Hi! I'm Ariel Sanchez (https://twitter.com/arielsanchezmor) and this is the result of a dream and the power of the vCommunity. I started a documentation template effort, which can be found [here](https://sites.google.com/site/arielsanchezmora/home/vmware/free-vmware-documentation-templates). There is a lot of work pending to be able to call the effort complete, but one very important component that my friend [Edgar Sanchez](https://github.com/edmsanchez) ( https://twitter.com/edmsanchez13 ) has advanced dramatically is the PowerCLI scripting. This repository stores them, and publishes them to the world so they can start being used. We open-sourced and placed in GitHub so they can be further improved by the vCommunity!

The main motivation for this project was the sad state of reliable documentation available to many vSphere administrators. It is demoralizing to start a new job, ask for documentation, and find there is none. It's sometimes worse that if there is documentation, it turns out to be outdated, or even worse, plain wrong! And it's also demoralizing to be tasked with creating documentation, realizing that creating it manually would take a long time, and that collecting and customizing all the scripts will take a long time as well.

Thus, our goal is to be able to easily produce documentation "direct from vCenter" that is relevant to what your manager or another VMware administrator wants to see. The best part is, you only need to run the scripts and they create the needed CSV or Excel file for you. This means you can update your documentation at a moment's notice, and even better, review it to identify things in your environment that may not have been easily visible before.

The license on these scripts is a MIT style license - use as you will. Like all the PowerCLI greats have told us before, steal and modify whatever you find useful. We definitely have stolen from all over the internet to create these (and have tried to credit those who we stole from). Special shout-outs to Luc Dekens, William Lam, Alan Renouf, Kyle Ruddy - and many more in the vCommunity.

Our goal is that this project is useful to others and it will be accepted in the official VMware PowerCLI examples. Please, let us know if you found this useful, had trouble running it, or anything that you want to see changed. We are new to GitHub but actively learning - use GitHub or reach out to us on twitter or in the VMware Code Slack (https://code.vmware.com/web/code/join)

To a future where walking into a new place and asking for documentation is greeted with "Yup, we use vDocumentation" and the interested party replies "Perfect!" :)

# Usage

. Once you have installed the module, you will be able to use the following functions:

__Get-ESXInventory__

__Get-ESXIODevice__

__Get-ESXNetworking__

__Get-ESXStorage__

Refer to the code's comments in the [vDocument Module File](https://github.com/arielsanchezmora/vDocumentation/blob/master/powershell/vDocument/vDocument.psm1) for full usage and examples, or use Get-Help and the module name:

__Get-Help Get-ESXInventory__

NAME
    Get-ESXInventory

SYNOPSIS
    Get basic ESXi host information


SYNTAX
    Get-ESXInventory [[-esxi] <Object>] [[-cluster] <Object>] [[-datacenter] <Object>] [-ExportCSV] [-ExportExcel]
    [-Hardware] [-Configuration] [[-folderPath] <Object>] [<CommonParameters>]


DESCRIPTION
    Will get inventory information for a vSphere Cluster, Datacenter or individual ESXi host
    The following is gathered:
    Hostname, Management IP, RAC IP, ESXi Version information, Hardware information
    and Host configuration


RELATED LINKS
    https://github.com/arielsanchezmora/vDocumentation

REMARKS
    To see the examples, type: "get-help Get-ESXInventory -examples".
    For more information, type: "get-help Get-ESXInventory -detailed".
    For technical information, type: "get-help Get-ESXInventory -full".
    For online help, type: "get-help Get-ESXInventory -online"


__Get-Help Get-ESXIODevice__

NAME
    Get-ESXIODevice

SYNOPSIS
    Get ESXi vmnic* and vmhba* VMKernel device information


SYNTAX
    Get-ESXIODevice [[-esxi] <Object>] [[-cluster] <Object>] [[-datacenter] <Object>] [-ExportCSV] [-ExportExcel]
    [[-folderPath] <Object>] [<CommonParameters>]


DESCRIPTION
    Will get PCI/IO Device information including HCL IDs for the below VMkernel name(s):
    Network Controller - vmnic*
    Storage Controller - vmhba*
    Graphic Device - vmgfx*
    All this can be gathered for a vSphere Cluster, Datacenter or individual ESXi host


RELATED LINKS
    https://github.com/arielsanchezmora/vDocumentation

REMARKS
    To see the examples, type: "get-help Get-ESXIODevice -examples".
    For more information, type: "get-help Get-ESXIODevice -detailed".
    For technical information, type: "get-help Get-ESXIODevice -full".
    For online help, type: "get-help Get-ESXIODevice -online"


__Get-Help Get-ESXNetworking__

NAME
    Get-ESXNetworking

SYNOPSIS
    Get ESXi Networking Details.


SYNTAX
    Get-ESXNetworking [[-esxi] <Object>] [[-cluster] <Object>] [[-datacenter] <Object>] [-ExportCSV] [-ExportExcel]
    [-VirtualSwitches] [-VMkernelAdapters] [-PhysicalAdapters] [[-folderPath] <Object>] [<CommonParameters>]


DESCRIPTION
    Will get Physical Adapters, Virtual Switches, and Port Groups
    All this can be gathered for a vSphere Cluster, Datacenter or individual ESXi host


RELATED LINKS
    https://github.com/arielsanchezmora/vDocumentation

REMARKS
    To see the examples, type: "get-help Get-ESXNetworking -examples".
    For more information, type: "get-help Get-ESXNetworking -detailed".
    For technical information, type: "get-help Get-ESXNetworking -full".
    For online help, type: "get-help Get-ESXNetworking -online"


__Get-Help Get-ESXStorage__

NAME
    Get-ESXStorage

SYNOPSIS
    Get ESXi Storage Details


SYNTAX
    Get-ESXStorage [[-esxi] <Object>] [[-cluster] <Object>] [[-datacenter] <Object>] [-ExportCSV] [-ExportExcel]
    [-StorageAdapters] [-Datastores] [[-folderPath] <Object>] [<CommonParameters>]


DESCRIPTION
    Will get iSCSI Software and Fibre Channel Adapter (HBA) details including Datastores
    All this can be gathered for a vSphere Cluster, Datacenter or individual ESXi host


RELATED LINKS
    https://github.com/arielsanchezmora/vDocumentation

REMARKS
    To see the examples, type: "get-help Get-ESXStorage -examples".
    For more information, type: "get-help Get-ESXStorage -detailed".
    For technical information, type: "get-help Get-ESXStorage -full".
    For online help, type: "get-help Get-ESXStorage -online"


# Licensing

Copyright (c) <2017> Ariel Sanchez and Edgar Sanchez

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

# Installation

The scripts run inside a PowerShell window using PowerCLI modules. Powershell is available in all modern windows OS, with PowerShell core available for Mac and Linux. Make sure you have the latest PowerCLI installed (you can check here for a video on how to install https://blogs.vmware.com/PowerCLI/2017/05/powercli-6-5-1-install-walkthrough.html)

From the video, these are the useful commands you should have completed before installing vDocumentation:

_$psversiontable_ [enter]  =  gives you the PowerShell version

_get-module VMware* -ListAvailable_ [enter]  =  Lists all installed PowerCLI modules, if return empty, install PowerCLI

## Installing PowerCLI
  _Find-Module -Name VMware.PowerCLI_  =  checks connectivity to PowerShell Gallery and updates NuGet if needed (yes is default)  
  _Install-Module -Name VMware.PowerCLI -Scope CurrentUser_  =  install PowerCLI as long as you answer Y or A

## Execution Policy and Certificate Warnings

 Make sure that your execution policy allows you to run scripts downloaded from the internet. You do this with a command run in a powershell window that has been launched with "Run as Administrator"
 
 _Set-ExecutionPolicy RemoteSigned_

and click Y or A

Unless you have proper certificates in your vSphere environment, some of the data collections may fail silently due to a certificate warning. Run this command so you never have to wonder:

_Set-PowerCLIConfiguration -InvalidCertificateAction Ignore_

Y is default

## Excel Module

While not required, having this module installed is recommended, as you can export direct to Excel. [Read about ImportExcel module.](https://github.com/dfinke/ImportExcel)

_Install-Module ImportExcel -scope CurrentUser_

## Adding the vDocumentation module

vDocumentation was created as a PowerShell module as well, and it's published in the PowerShell Gallery, so we can use the Install-Module command:

![install_vDocumentation_1.03](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/install_vDocumentation_1.03.png)

If you can't use the online method, use this manual process:

  1 Download the two files inside the vDocumentation folder.  
  2 Browse to the %USERPROFILE%\Documents\WindowsPowerShell\Modules and copy the files inside a folder named vDocumentation  
  3 Close all PowerShell windows  
  4 Launch PowerShell again, you should be able to use the vDocumentation functions now


## One method to copy the needed files from Github to your PC using PowerShell:

Execute these lines in a PowerShell window that is in your home directory (tested with PS 5)

_mkdir Documents\WindowsPowerShell\Modules\vDocumentation_

_(new-object Net.WebClient).DownloadString("https://raw.githubusercontent.com/arielsanchezmora/vDocumentation/master/powershell/vDocument/vDocument.psd1") > Documents\WindowsPowerShell\Modules\vDocument\vDocumentation.psd1_

_(new-object Net.WebClient).DownloadString("https://raw.githubusercontent.com/arielsanchezmora/vDocumentation/master/powershell/vDocument/vDocument.psm1") > Documents\WindowsPowerShell\Modules\vDocument\vDocumentation.psm1_

_exit_
