$TestShortcut = [PSObject]@{
    Path             = 'C:\Windows\Temp\cShortcut_Integration\test.lnk'
    Target           = 'C:\Program Files\Internet Explorer\iexplore.exe'
    Arguments        = '-private'
    WindowStyle      = 'maximized'
    WorkingDirectory = 'C:\work'
    Description      = 'This is a shortcut to the IE'
    Icon             = 'shell32.dll,277'
    HotKey           = 'Ctrl+Shift+U'
    HotKeyCode       = 0x0355
    AppUserModelID   = 'Microsoft.InternetExplorer.Default'
}



Configuration cShortcut_Config {
    Import-DscResource -ModuleName DSCR_Shortcut

    Node localhost {
        cShortcut Integration_Test
        {
            Path             = $TestShortcut.Path
            Target           = $TestShortcut.Target
            Arguments        = $TestShortcut.Arguments
            WindowStyle      = $TestShortcut.WindowStyle
            WorkingDirectory = $TestShortcut.WorkingDirectory
            Description      = $TestShortcut.Description
            Icon             = $TestShortcut.Icon
            HotKey           = $TestShortcut.HotKey
            AppUserModelID   = $TestShortcut.AppUserModelID
        }
    }
}
