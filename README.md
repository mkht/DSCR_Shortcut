DSCR_Shortcut
====

PowerShell DSC Resource to create shortcut file.

## Install
You can install Resource through [PowerShell Gallery](https://www.powershellgallery.com/packages/DSCR_Shortcut/).
```Powershell
Install-Module -Name DSCR_Shortcut
```

## Resources
* **cShortcut**
PowerShell DSC Resource to create shortcut file.

## Properties
### cShortcut
+ [string] **Ensure** (Write):
    + Specify whether or not a shortcut file exists
    + The default value is Present. { Present | Absent }

+ [string] **Path** (Key):
    + The path of the shortcut file.
    + If the path ends with something other than `.lnk`, The extension will be added automatically to the end of the path

+ [string] **Target** (Required):
    + The target path of the shortcut.

+ [string] **Arguments** (Write):
    + The arguments of the shortcut.

+ [string] **WorkingDirectory** (Write):
    + The working directory of the shortcut.

+ [string] **WindowStyle** (Write):
    + You can select window style. { normal | maximized | minimized }
    + The default value is `normal`

## Examples
+ **Example 1**: Create a shortcut to the Internet Explore InPrivate mode to the Administrator's desktop
```Powershell
Configuration Example1
{
    Import-DscResource -ModuleName DSCR_Shortcut
    cShortcut IE_Desktop
    {
        Path = 'C:\Users\Administrator\Desktop\PrivateIE.lnk'
        Target = "C:\Program Files\Internet Explorer\iexplore.exe"
        Arguments = '-private'
    }
}
```
