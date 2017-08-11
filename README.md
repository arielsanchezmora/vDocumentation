# vDocumentation
vDocumentation provides a community-created set of PowerCLI scripts that produce documentation of vSphere environments in Excel file format.

Hi! I'm Ariel Sanchez (https://twitter.com/arielsanchezmor) and this is my brainchild. I started a documentation template effort, which can be found [here](https://sites.google.com/site/arielsanchezmora/home/vmware/free-vmware-documentation-templates). There is a lot of work to do in it, but one very important component that my friend [Edgar Sanchez](https://github.com/edmsanchez) ( https://twitter.com/edmsanchez13 ) has completed is the PowerCLI scripting. This repository stores them, and publishes them to the world so they can start being used and improved by the community!

The main motivation for this project was the sad state of reliable documentation available to many vSphere administrators. It is demoralizing to start a new job, ask for documentation, and find there is none, or what there is turns out to be outdated, or even worse, wrong! And it's also demoralizing to be tasked with creating documentation, realizing that creating it manually would take a long time, and that figuring out all the scripts will take probably longer.

Thus, easy to create documentation "direct from vCenter" that is relevant to what your manager or another VMware administrator wants to see is our goal. The best part is, you only need to run the scripts and they create the needed Excel file for you. This means you can update your documentation at a moment's notice, and even better, use it to identify things in your environment that may not have been easily visible before.

The license on these scripts is a BSD style license - use as you will. Like all the PowerCLI greats have told us before, steal and modify whatever you find useful. We definitely have stolen from all over the internet to create these (and have tried to credit those who we stole from). Special shout-outs to Luc Dekens, William Lam, Alan Renouf, Kyle Ruddy - and many more in the vCommunity.

Our goal is that this project is useful to others and it will be accepted in the official VMware PowerCLI examples. Please, let us know if you found this useful, had trouble running it, or anything that you want to see changed. We are new to GitHub but actively learning - use GitHub or reach out to us on twitter or on the VMware Code Slack (https://code.vmware.com/web/code/join)

To a future where walking into a new place and asking for documentation is greeted with "Yup, we use vDocument" and the interested party replies "Perfect!" :)

# Module Changelog

v1.0.3 new functionality added:
 Get-ESXInventory: Added RAC Firmware version, BIOS release date. 
 Get-ESXIODevice: Added support to get HP Smart Array Firmware from PowerCLI
 
1.0.2 Formatting & Manifest changes

1.0.1 Changes to support displaying datastore multipathing

1.0 First release to PowerShell Gallery with 4 commands: Get-ESXInventory, Get-ESXIODevice, Get-ESXNetworking & Get-ESXStorage


# Usage

. Once you have installed the module, you will be able to use the following functions:

__Get-ESXInventory__

__Get-ESXIODevice__

__Get-ESXNetworking__

Refer to the code's comments in the [vDocument Module File](https://github.com/arielsanchezmora/vDocumentation/blob/master/powershell/vDocument/vDocument.psm1) for full usage and examples, or use Get-Help and the module name:

__Get-Help Get-ESXInventory__

NAME
    Get-ESXInventory

SYNOPSIS
    Get basic ESXi host information


SYNTAX
    Get-ESXInventory [[-esxi] <Object>] [[-cluster] <Object>] [[-datacenter] <Object>] [-ExportCSV] [-ExportExcel]
    [[-folderPath] <Object>] [<CommonParameters>]


DESCRIPTION
    Will get inventory information for a vSphere Cluster, Datacenter or individual ESXi host
    The following is gathered:
    Hostname, Management IP, RAC IP, ESXi Version information, Hardware information


RELATED LINKS
    https://github.com/edmsanchez/vDocumentation

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
    https://github.com/edmsanchez/vDocumentation

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
    https://github.com/edmsanchez/vDocumentation

REMARKS
    To see the examples, type: "get-help Get-ESXNetworking -examples".
    For more information, type: "get-help Get-ESXNetworking -detailed".
    For technical information, type: "get-help Get-ESXNetworking -full".
    For online help, type: "get-help Get-ESXNetworking -online"



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

vDocumentation are powershell modules as well, but are not yet in the PowerShell Gallery, so we can't use the Install-Module command.  For now, use this manual process:

  1 Download the two files inside the vDocumentation folder.
  
  2 Browse to the %USERPROFILE%\Documents\WindowsPowerShell\Modules and copy the files inside a folder named vDocumentation
  
  3 Close all PowerShell windows
  
  4 Launch PowerShell, you should be able to use the vDocumentation functions now


## One method to copy the needed files from Github to your PC using PowerShell:

Execute these lines in a PowerShell window that is in your home directory (tested with PS 5)

_mkdir Documents\WindowsPowerShell\Modules\vDocument_

_(new-object Net.WebClient).DownloadString("https://raw.githubusercontent.com/arielsanchezmora/vDocumentation/master/powershell/vDocument/vDocument.psd1") > Documents\WindowsPowerShell\Modules\vDocument\vDocument.psd1_

_(new-object Net.WebClient).DownloadString("https://raw.githubusercontent.com/arielsanchezmora/vDocumentation/master/powershell/vDocument/vDocument.psm1") > Documents\WindowsPowerShell\Modules\vDocument\vDocument.psm1_

_exit_
