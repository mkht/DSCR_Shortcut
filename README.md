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
    + The default value is `Present`. { Present | Absent }

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

+ [string] **Description** (Write):
    + The description of the shortcut.

+ [string] **Icon** (Write):
    + The path of the icon resource.

+ [string[]] **HotKey** (Write):
    + HotKey (Shortcut Key) of the shortcut
    + HotKey works only for shortcuts on the desktop or in the Start menu.
    + The syntax is: `"{KeyModifier} + {KeyName}"` ( e.g. `"Alt+Ctrl+Q"`, `"Shift+F9"` )
    + If the hotkey not working after configuration, try to reboot.

## Examples
+ **Example 1**: Create a shortcut to the Internet Explore InPrivate mode to the Administrator's desktop
```Powershell
Configuration Example1
{
    Import-DscResource -ModuleName DSCR_Shortcut
    cShortcut IE_Desktop
    {
        Path      = 'C:\Users\Administrator\Desktop\PrivateIE.lnk'
        Target    = "C:\Program Files\Internet Explorer\iexplore.exe"
        Arguments = '-private'
    }
}
```

+ **Example 2**: WindowStyle, WorkingDirectory, Description, Icon, Hotkey
```Powershell
Configuration Example2
{
    Import-DscResource -ModuleName DSCR_Shortcut
    cShortcut IE_Desktop
    {
        Path             = 'C:\Users\Administrator\Desktop\PrivateIE.lnk'
        Target           = "C:\Program Files\Internet Explorer\iexplore.exe"
        Arguments        = '-private'
        WindowStyle      = 'maximized'
        WorkingDirectory = 'C:\work'
        Description      = 'This is a shortcut to the IE'
        Icon             = 'shell32.dll,277'
        HotKey           = 'Ctrl+Shift+U'
    }
}
```

## ChangeLog
### v1.3.0
+ Add `Description` property #1
+ Add `HotKey` property #2
+ Add `Icon` property #3
