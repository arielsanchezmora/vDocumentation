# vDocumentation

vDocumentation provides a community-created set of PowerCLI scripts that produce infrastructure documentation of vSphere environments in CSV or Excel file format. It was presented for general public use in VMworld 2017, session SER2077BU. You can watch the video here

https://www.youtube.com/watch?v=-KK0ih8tuTo

Original slides are [here](https://www.dropbox.com/s/f5e9hpxgzz0unq1/vmworld2017-Ariel%20and%20Edgar%20Sanchez-SER2077BU-Achieve%20Maximum%20vSphere%20Stability%20with%20PowerCLI%20Assisted%20Documentation%20From%20Buildout%20to%20Daily%20Administration.pptx?dl=0) as well as the [mindmap](https://www.dropbox.com/s/19jdgup6ldah3u9/SER2077BU%20Achieve%20maximum%20vSphere%20stability%20with%20PowerCLI%20assisted%20documentation%20%20from%20buildout%20to%20daily%20administration-mindmap201707231829EST.png?dl=0) we used to create this talk. We are passionate about this subject so please use the slides or let us know what you would like to add to the MindMap, and we can continue improving this presentation.

# Changelog

__v2.3.0__ Very meaty update, with a new cmdlet aimed at verifying the first wave of vSphere mitigations against Meltdown and Spectre developed by project lead Edgar Sanchez (twitter <a href="https://twitter.com/edmsanchez13/" target="_blank"> @edmsanchez13</a>. A much better overview of the new function can be found on his blog  <a href="https://virtualcornerstone.com/2018/01/08/validating-compliance-of-vmsa-2018-0002-and-bios-update/" target="_blank"> virtualcornerstone.com</a>)

 *Additions:*  
 - Added **Get-ESXSpeculativeExecution** Cmdlet to check compliane for VMSA-2018-0002 Security Advisory and BIOS version. He is already working on additional checks for v2.3.1

 *Bug Fixes:*  
 - Code fix in Get-ESXPatching - now able to get reference URLss for description fields containing "https" (example: ESXi650-201712103-SG is https://kb.vmware.com/kb/000051196)  
 - Get-ESXInventory  - Added more details for ESXi Install source

        * Device Model
        * Boot Device
        * Runtime Name
        * Device Path

__v2.2.0__ Another meaty update, with a new vSAN cmdlet donated by Graham Barker (twitter <a href="https://twitter.com/VirtualG_UK" target="_blank"> @VirtualG_UK</a> website  <a href="https://virtualg.uk/" target="_blank"> virtualg.uk</a>)! This brings the total number of vDocumentation cmdlets to six from our initial launch of 4! 

 *Additions:*  
- Added new Cmdlet: **Get-vSANInfo**, NOTE! It depends on _Get-VsanClusterConfiguration_ which only works on vSphere 6.5! Documentation update and examples coming soon!  
- Added RAC MAC to **Get-ESXInventory**
        
 *Bug Fixes:*  
- Minor code fixes in Get-ESXInventory, Get-ESXIODevice, and Get-ESXPatching

__v2.1.0__ Meaty update, our first new cmdlet since the project's debut!  

 *Additions:*  
- Added new Cmdlet: **Get-ESXPatching**, documentation update and examples coming soon!  
- Added the following to **Get-ESXInventory**, Configuration tab: SSH and ESXi Shell Service details requested by akozlow in Issue #19, and Boot Time
        
 *Bug Fixes:*  
- Fixed reported issue #16 by DaveBF 'VMHostNeworkInfo type is deprecated' in Get-ESXNetworking Cmdlet
- Fixed issue for Uptime in Get-ESXInventory where it was not being calculated correctly

__v2.0.0__ Major update, on the backend, mostly safe for actual users  

 *Code cleaning:*  
 Each script module exists now in its own .ps1 file which will allow easier editing by the community  
 Scripts code optimization and formatting updated  
 [@jpsider](https://github.com/jpsider) championed the removal of the CLS command that would clear screen before starting screen output, and contributed the code, which was included in this release.  
 
 *Removed:*  
 Get-ESXInventory function (and thus, a report column) removed: Deprecated script Cmdlet - Software/Patch Name(s) from host configuration has been deprecated. What Patches gets pushed can be manually verified using the Build ID  
 
 *Additions:*  
  [@jpsider](https://github.com/jpsider) championed the addition of a -passthru option and contributed the code, which was included in this release.  
 Get-ESXInventory - Host Configuration script now has the following:  
 - Gather ESXi Installation Type and Boot source
 - Gather ESXi Image Profile
 - Gather ESXi Software Acceptance Level
 - Gather ESXi Uptime (thanks to the person who asked in #SER2077BU, send us your name to give you credit!)
 - Gather ESXi Install Date
 
 Get-ESXIODevice - NIC and HBA script now has the following:
 - Updated string match to check for HPSA firmware, as it changed between 5.5, and 6.0 and possibly between firmware versions.
 
 *Bug Fixes:*  
 Fixed Get-ESXNetworking script Cmdlet when querying UCS environment, or 3rd party Distributed switches.  While the information retrieved is not the same (due to the powershell command, not because of vDocumentation) the script will no longer fail, and will produce what it can.
 
__v1.0.4__ new functionality added:  

 Updated export-excel so that it does no number conversion (IP addresses are now text) on any of the columns and it auto sizes them. Thanks to [@magneet_nl](https://twitter.com/Magneet_nl) for helping us discover this bug!

__v1.0.3__ new functionality added:  

 Get-ESXInventory: Added RAC Firmware version, BIOS release date.  
 Get-ESXIODevice: Added support to get HP Smart Array Firmware from PowerCLI  
 
__1.0.2__ Formatting & Manifest changes

__1.0.1__ Changes to support displaying datastore multipathing

__1.0.0__ First release to PowerShell Gallery with 4 commands: Get-ESXInventory, Get-ESXIODevice, Get-ESXNetworking & Get-ESXStorage


# Quickstart instructions

## First time usage on a *brand new machine* with PowerShell 5.x and an open internet connection

_Paste in a PowerShell window that has been Run as Administrator and answer Y_

**Set-ExecutionPolicy RemoteSigned**  

![Run PowerShell as Administrator](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/PowerShell_as_administrator.png)

_You can now close the PowerShell window that ran as Administrator_ 

_In a new, **normal** PowerShell console, paste the below commands answering Y (this only affects your user, and it may take a while)_

**Install-Module -Name VMware.PowerCLI -Scope CurrentUser**  
**Install-Module ImportExcel -scope CurrentUser**  
**Install-Module vDocumentation -Scope CurrentUser**  

![Install PowerCLI, ImportExcel and vDocumentation modules](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/install_PowerCLI_ImportExcel_vDocumentation.png)

_vDocumentation is now installed! You can verify with_

**Get-Module vDocumentation -ListAvailable | Format-List**

![Confirm vDocumentation installation](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/Confirm_vDocumentation_installation2.png)


## The vDocumentation module gives you five _new_ PowerCLI Commands you can use to create documentation of a vSphere environment

_Before you can use them, connect to your vCenter(s) using PowerCLI_

**Connect-VIServer [IP_or_FQDN_of_vCenter]**      _# Connect to one, or repeat for many vCenters_

_When prompted for credentials use a vCenter Administrator-level account. Once connected you can execute these commands:_

|Command|Description|
|----------------|---|
|**Get-ESXInventory**|Document host hardware inventory and host configuration|
|**Get-ESXIODevice**|Document information from HBAs, NICs and other PCIe devices including PCI IDs, MACs, firmware & drivers|
|**Get-ESXNetworking**|Document networking configuration information such as NICs, vSwitches, VMKernel details|
|**Get-ESXStorage**|Document storage configurations such as iSCSI details, FibreChannel, Datastores & Multipathing|
|**Get-ESXPatching**|Document installed and pending patches, including related time and KB information|

_Each script will output the corresponding data to terminal, and optionally create a file (XLSX, CSV) with the command name and a timestamp. You can use command switches to customize CSV or Excel output, file path (default is powershell working directory), and the command scope (report on all connected vCenters or just cluster or host)._

## Command switch options

_Running a command **without** switches will_
- report on all virtual datacenters in all connected vCenters
- output to PowerShell terminal only
- include all data tabs for each command

_To change this behaviour use these switches:_

|Scope|Switch|Description|
|---|---|---|
|Target|**-esxi**|Get information from a particular host (for several, use commas)|
|Target|**-cluster**|Get information from a particular cluster (for several, use commas)|
|Target|**-datacenter**|Get information from a particular virtual datacenter (for several, use commas)|
|Output|**-folderPath**|Specify the path to save the file name|
|Output|**-ExportCSV**|The output will be written to a CSV file|
|Output|**-ExportExcel**|The output will be written to a XLSX file (if ImportExcel module is not installed will do CSV)|
|Info Tab|**-Hardware**|For Get-ESXInventory: explicitly outputs the Hardware tab|
|Info Tab|**-Configuration**|For Get-ESXInventory: explicitly outputs the Configuration tab|
|Info Tab|**-VirtualSwitches**|For Get-ESXNetworking: explicitly outputs the VirtualSwitches tab|
|Info Tab|**-VMkernelAdapters**|For Get-ESXNetworking: explicitly outputs the VMkernelAdapters tab|
|Info Tab|**-PhysicalAdapters**|For Get-ESXNetworking: explicitly outputs the PhysicalAdapters tab|
|Info Tab|**-StorageAdapters**|For Get-ESXStorage: explicitly outputs the StorageAdapters tab|
|Info Tab|**-Datastores**|For Get-ESXStorage: explicitly outputs the Datastores tab|


You can see the full syntax with the **Get-Help** command

`get-help Get-ESXInventory -ShowWindow`

![Get-Help Example](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/get-help_example.png)

## Example Outputs

Get-ESXInventory -Hardware

![Get-ESXInventory -Hardware](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/get-esxinventory-hardware_output.png)

Get-ESXInventory -Configuration

![Get-ESXInventory -Configuration](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/get-esxinventory-configuration_output.png)

Get-ESXIODevice _(only has one tab)_

![Get-ESXIODevice](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/Get-ESXIODevice_output.png)

Get-ESXNetworking -VirtualSwitches _(standard switch)_

![Get-ESXNetworking -VirtualSwitches](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/Get-ESXNetworking-VirtualSwitches_output.png)

Get-ESXNetworking -VirtualSwitches _(distributed switch)_

![Get-ESXNetworking -VirtualSwitches](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/Get-ESXNetworking-VirtualSwitches_output2.png)

Get-ESXNetworking -VMKernelAdapter

![Get-ESXNetworking -VMKernelAdapter](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/Get-ESXNetworking-VMkernelAdapters_output.png)

Get-ESXNetworking -PhysicalAdapters

![Get-ESXNetworking -PhysicalAdapters](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/Get-ESXNetworking-PhysicalAdapters_output.png)

Get-ESXStorage -StorageAdapters

![Get-ESXStorage -StorageAdapters](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/Get-ESXStorage-StorageAdapters_output.png)

Get-ESXStorage -Datastores

![Get-ESXStorage -Datastores](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/Get-ESXStorage-Datastores_output.png)

**iSCSI output thanks to [@michael_rudloff](https://twitter.com/michael_rudloff), see his ![full output](https://github.com/arielsanchezmora/vDocumentation/blob/master/example-outputs/michael_rudloff/README.MD)**

![michael_iscsi_physical](https://github.com/arielsanchezmora/vDocumentation/blob/master/example-outputs/michael_rudloff/michael_iscsi_physical.png)  
![michael_iscsi_VMKernel](https://github.com/arielsanchezmora/vDocumentation/blob/master/example-outputs/michael_rudloff/michael_iscsi_VMKernel.png)  
![michael_iscsi_iSCSI](https://github.com/arielsanchezmora/vDocumentation/blob/master/example-outputs/michael_rudloff/michael_iscsi_iSCSI.png)  
![michael_iscsi_datastore](https://github.com/arielsanchezmora/vDocumentation/blob/master/example-outputs/michael_rudloff/michael_iscsi_datastore.png)  

**CSV outputs thanks to [@magneet_nl](https://twitter.com/Magneet_nl), see his ![full output](https://github.com/arielsanchezmora/vDocumentation/blob/master/example-outputs/Magneet_nl/README.MD)**

![magneet_csv1](https://github.com/arielsanchezmora/vDocumentation/blob/master/example-outputs/Magneet_nl/magneet_csv_hardware_1.png)   
![magneet_csv2](https://github.com/arielsanchezmora/vDocumentation/blob/master/example-outputs/Magneet_nl/magneet_csv_hardware_2.png)  
![magneet_csv3](https://github.com/arielsanchezmora/vDocumentation/blob/master/example-outputs/Magneet_nl/magneet_csv_hardware_3.png)  
![magneet_csv4](https://github.com/arielsanchezmora/vDocumentation/blob/master/example-outputs/Magneet_nl/magneet_csv_hardware_4.png)  



**[Document your vSphere environment? Yes you can!](https://notesfrommwhite.net/2017/08/16/document-your-vsphere-environment-yes-you-can/) Blog article with Excel outputs thanks to [@mwVme](https://twitter.com/mwVme)**  


## Uninstalling the vDocumentation script

**Uninstall-Module vDocumentation**

![Uninstall vDocumentation](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/uninstall_vDocumentation.png)

## Upgrading from a previous version

There is a known limitation in just upgrading through the PowerShell Gallery: using the Update-Module command installs a new version but does **not** remove the old version. While PowerShell/PowerCLI will use the latest module, if you wish to only have the latest listed in your computer, uninstall all existing vDocumentation modules **before** installing the latest by using Uninstall-Module as many times as needed, before using **Install-Module** as with a new installation.  

However, in an effort to keep it simple, you can just use the following commands (and again, it does seem it always uses the latest version). If the prompt returns without doing anything, you are already running the latest.

**Update-Module VMware.PowerCLI**  
**Update-Module ImportExcel**  
**Update-Module vDocumentation**

![Upgrade Commands](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/upgrade_commands.png)


## FAQ

What if I don't have internet?

- _A great guide to follow is https://blogs.vmware.com/PowerCLI/2017/04/powercli-install-process-powershell-gallery.html_

How do I know which PowerShell version I am running?

|OS|Default Version|Upgradeable to 5.x|
|---|---|---|
|Windows 7|2.0|Yes, manually|
|Windows Server 2008 R2|2.0|Yes, manually|
|Windows 8|3.0|Yes, manually|
|Windows Server 2012|3.0|Yes, manually|
|Windows 10|5.0|Included|
|Windows Server 2016|5.0|Included|

_To upgrade follow links such as https://docs.microsoft.com/en-us/powershell/scripting/setup/windows-powershell-system-requirements?view=powershell-5.1_

What if I can't run PowerShell 5.x?

- _If ESXi and vCenter are hardened to only talk on TLS v1.2, you need .Net 4.5 or above for PowerShell to support this._

What is the ImportExcel module?

- _[Read about ImportExcel](https://github.com/dfinke/ImportExcel)_

Does this run on PowerCLI core?

- _We'd love to know! We haven't tested it yet; expect an update soon._

Why do I get a warning about deprecated features when running the script?

- _This is native from PowerCLI as they plan future changes. vDocumentation does not use any feature that is known to be in deprecation plans. You can disable the warnings with `Set-PowerCLIConfiguration -DisplayDeprecationWarnings $false -Scope User`_

I get certificate warnings

- _You can disable self-signed certificate warnings with the following command, or install proper certs ;)_

**Set-PowerCLIConfiguration -InvalidCertificateAction Ignore**

![Enable remote scripts and ignore certificate warnings](https://github.com/arielsanchezmora/vDocumentation/blob/master/pictures/enable_RemoteSigned_Invalid_Certificate.png)

I get this error "Get-EsxCli : A parameter cannot be found that matches parameter name 'V2'" why?

- _This probably means you are running a version of PowerCLI that is older than 6.3. We encourage uninstalling all versions and then using the latest version - that should take care of this error, which comes from a feature that was added in PowerCLI 6.3_



# vDocumentation backstory

Hi! I'm Ariel Sanchez (https://twitter.com/arielsanchezmor) and this is the result of a dream and the power of the vCommunity. I started a documentation template effort, which can be found [here](https://sites.google.com/site/arielsanchezmora/home/vmware/free-vmware-documentation-templates). There is a lot of work pending to be able to call the effort complete, but one very important component that my friend [Edgar Sanchez](https://github.com/edmsanchez) ( https://twitter.com/edmsanchez13 ) has advanced dramatically is the PowerCLI scripting. This repository stores them, and publishes them to the world so they can start being used. We open-sourced and placed in GitHub so they can be further improved by the vCommunity!

The main motivation for this project was the sad state of vSphere infrastructure documentation accessible to many vSphere administrators. It is demoralizing to start a new job, ask for documentation, and find there is none. The situation is bad enough when the documentation is outdated, but even worse when it's plain wrong. It's also challenging to be tasked with creating documentation, realizing that creating it manually would take a long time, and that collecting and customizing all the scripts will take a long time as well.

Thus, our goal is to be able to easily produce documentation "direct from vCenter" that is relevant to what your manager or another VMware administrator wants to see. The best part is, you only need to run the scripts and they create the needed CSV or Excel file for you. This means you can update your documentation at a moment's notice, and even better, review it to identify things in your environment that may not have been easily visible before.

The license on these scripts is a MIT style license - use as you will. Like all the PowerCLI greats have told us before, steal and modify whatever you find useful. We definitely have stolen from all over the internet to create these (and have tried to credit those who we stole from). Special shout-outs to Luc Dekens, William Lam, Alan Renouf, Kyle Ruddy - and many more in the vCommunity.

Our goal is that this project is useful to others and it will be accepted in the official VMware PowerCLI examples. Please, let us know if you found this useful, had trouble running it, or anything that you want to see changed. We are new to GitHub but actively learning - use GitHub or reach out to us on twitter or in the VMware Code Slack (https://code.vmware.com/web/code/join)

To a future where walking into a new place and asking for documentation is greeted with "Yup, we use vDocumentation" and the interested party replies "Perfect!" :)

# Syntax

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
    

__Get-Help Get-ESXPatching__  

NAME
    Get-ESXPatching

SYNOPSIS
    Get ESXi patch compliance


SYNTAX
    Get-ESXPatching [[-esxi] <Object>] [[-cluster] <Object>] [[-datacenter] <Object>] [[-baseline] <Object>]
    [-ExportCSV] [-ExportExcel] [-Patching] [-PassThru] [[-folderPath] <Object>] [<CommonParameters>]


DESCRIPTION
    Will get patch compliance for a vSphere Cluster, Datacenter or individual ESXi host


RELATED LINKS
    https://github.com/arielsanchezmora/vDocumentation

REMARKS
    To see the examples, type: "get-help Get-ESXPatching -examples".
    For more information, type: "get-help Get-ESXPatching -detailed".
    For technical information, type: "get-help Get-ESXPatching -full".
    For online help, type: "get-help Get-ESXPatching -online"


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

_Install-Module vDocumentation -scope CurrentUser_

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
