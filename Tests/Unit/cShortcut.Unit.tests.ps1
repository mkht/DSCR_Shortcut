#Requires #Requires -Module @{ ModuleName = 'Pester'; RequiredVersion = '4.10.1' }

$script:moduleRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $script:moduleRoot '\DSCResources\cShortcut\cShortcut.psm1') -Force
$global:TestData = Join-Path (Split-Path -Parent $PSScriptRoot) '\TestData'

#region Begin Testing
InModuleScope 'cShortcut' {
    $script:TestGuid = [Guid]::NewGuid()

    #region Tests for Get-TargetResource
    Describe 'cShortcut/Get-TargetResource' -Tag 'Unit' {

        BeforeAll {
            $ErrorActionPreference = 'Stop'
        }

        Context 'Absent' {
            It 'Return Ensure = "Absent" if the specified path does not exist.' {
                $PathNotExist = Join-Path $TestDrive '\NotExist\Nothing.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                $getParam = @{
                    Path   = $PathNotExist
                    Target = $Target
                }

                $Result = Get-TargetResource @getParam

                $Result.Ensure | Should -Be 'Absent'
                $Result.Path | Should -Be $PathNotExist
                $Result.Target | Should -Be $null
            }

            It 'Add .lnk extension if the specified path does not ends with ".lnk".' {
                $PathWithoutLnkExt = Join-Path $TestDrive '\NotExist\SomeName.txt'
                $Target = 'C:\Windows\System32\notepad.exe'

                $getParam = @{
                    Path   = $PathWithoutLnkExt
                    Target = $Target
                }

                $Result = Get-TargetResource @getParam

                $Result.Path | Should -Be ($PathWithoutLnkExt + '.lnk')
            }

            It 'The Path parameter is mandatory.' {
                (Get-Command 'cShortcut\Get-TargetResource').Parameters['Path'].Attributes.Mandatory | Should -BeTrue
            }

            It 'The Target parameter is mandatory.' {
                (Get-Command 'cShortcut\Get-TargetResource').Parameters['Target'].Attributes.Mandatory | Should -BeTrue
            }
        }

        Context 'Present' {

            Mock Get-Shortcut {
                @{
                    TargetPath       = 'Mock_TargetPath'
                    WorkingDirectory = 'Mock_WorkingDirectory'
                    Arguments        = 'Mock_Arguments'
                    Description      = 'Mock_Description'
                    IconLocation     = 'Mock_IconLocation'
                    HotKey           = 0x0141
                    WindowStyle      = 'maximized'
                    AppUserModelID   = 'Mock_AppUserModelID'
                }
            }

            Mock ConvertTo-HotKeyString { return 'Shift+B' }

            It 'Return correct properties when the specified path exists.' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                New-Item $PathExists -ItemType File -Force

                $getParam = @{
                    Path   = $PathExists
                    Target = $Target
                }

                $Result = Get-TargetResource @getParam

                Assert-MockCalled 'Get-Shortcut' -Times 1 -Exactly -Scope It
                Assert-MockCalled 'ConvertTo-HotKeyString' -Times 1 -Exactly -Scope It

                $Result.Ensure | Should -Be 'Present'
                $Result.Path | Should -Be $PathExists
                $Result.Target | Should -Be 'Mock_TargetPath'
                $Result.WorkingDirectory | Should -Be 'Mock_WorkingDirectory'
                $Result.Arguments | Should -Be 'Mock_Arguments'
                $Result.Description | Should -Be 'Mock_Description'
                $Result.Icon | Should -Be 'Mock_IconLocation'
                $Result.HotKey | Should -Be 'Shift+B'
                $Result.HotKeyCode | Should -Be 0x0141
                $Result.WindowStyle | Should -Be 'maximized'
                $Result.AppUserModelID | Should -Be 'Mock_AppUserModelID'
            }
        }
    }
    #endregion Tests for Get-TargetResource


    #region Tests for Test-TargetResource
    Describe 'cShortcut/Test-TargetResource' -Tag 'Unit' {

        BeforeAll {
            $ErrorActionPreference = 'Stop'
        }

        Context 'Get-TargetResource returns "Absent"' {

            Mock Get-TargetResource { return @{Ensure = 'Absent' } }

            It 'Return $true when the Ensure is Absent and the Path does not exist.' {
                $PathNotExist = Join-Path $TestDrive '\NotExist\Nothing.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                $testParam = @{
                    Ensure = 'Absent'
                    Path   = $PathNotExist
                    Target = $Target
                }

                Test-TargetResource @testParam | Should -Be $true
                Assert-MockCalled -CommandName 'Get-TargetResource' -Exactly -Times 0 -Scope It
            }

            It 'Return $false when the Ensure is Absent and the Path exists.' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                New-Item $PathExists -ItemType File -Force

                $testParam = @{
                    Ensure = 'Absent'
                    Path   = $PathExists
                    Target = $Target
                }

                Test-TargetResource @testParam | Should -Be $false
                Assert-MockCalled -CommandName 'Get-TargetResource' -Exactly -Times 0 -Scope It
            }

            It 'Return $false when the Ensure is Absent and the Path exists (without extension).' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $PathWithoutExtension = Join-Path $TestDrive '\Exist\shortcut'
                $Target = 'C:\Windows\System32\notepad.exe'

                New-Item $PathExists -ItemType File -Force

                $testParam = @{
                    Ensure = 'Absent'
                    Path   = $PathWithoutExtension
                    Target = $Target
                }

                Test-TargetResource @testParam | Should -Be $false
                Assert-MockCalled -CommandName 'Get-TargetResource' -Exactly -Times 0 -Scope It
            }

            It 'Return $false when the Ensure is Present' {
                $PathNotExist = Join-Path $TestDrive '\NotExist\Nothing.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                $testParam = @{
                    Ensure = 'Present'
                    Path   = $PathNotExist
                    Target = $Target
                }

                Test-TargetResource @testParam | Should -Be $false
                Assert-MockCalled -CommandName 'Get-TargetResource' -Exactly -Times 1 -Scope It
            }
        }

        Context 'Get-TargetResource returns "Present"' {

            Mock Get-TargetResource {
                @{
                    Ensure           = 'Present'
                    Path             = Join-Path $TestDrive '\Exist\shortcut.lnk'
                    Target           = 'C:\Windows\System32\notepad.exe'
                    WorkingDirectory = 'Mock_WorkingDirectory'
                    Arguments        = 'Mock_Arguments'
                    Description      = 'Mock_Description'
                    Icon             = 'Mock_IconLocation,0'
                    HotKey           = 'Shift+B'
                    HotKeyCode       = 0x0142
                    WindowStyle      = 'maximized'
                    AppUserModelID   = 'Mock_AppUserModelID'
                }
            }

            Mock Format-HotKeyString { return 'Shift+A' } -ParameterFilter { $HotKey -eq 'Shift+A' }
            Mock Format-HotKeyString { return 'Shift+B' } -ParameterFilter { $HotKey -eq 'Shift+B' }
            Mock ConvertFrom-HotKeyString { return 0x0141 } -ParameterFilter { $HotKey -eq 'Shift+A' }
            Mock ConvertFrom-HotKeyString { return 0x0142 } -ParameterFilter { $HotKey -eq 'Shift+B' }

            It 'Return $true when the Ensure is Present and match all specified properties (single property)' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                New-Item $PathExists -ItemType File -Force

                $testParam = @{
                    Ensure = 'Present'
                    Path   = $PathExists
                    Target = $Target
                }

                Test-TargetResource @testParam | Should -Be $true
                Assert-MockCalled -CommandName 'Get-TargetResource' -Exactly -Times 1 -Scope It
                Assert-MockCalled -CommandName 'Format-HotKeyString' -Exactly -Times 0 -Scope It
                Assert-MockCalled -CommandName 'ConvertFrom-HotKeyString' -Exactly -Times 0 -Scope It
            }

            It 'Return $true when the Ensure is Present and match all specified properties (two properties)' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                New-Item $PathExists -ItemType File -Force

                $testParam = @{
                    Ensure      = 'Present'
                    Path        = $PathExists
                    Target      = $Target
                    Description = 'Mock_Description'
                }

                Test-TargetResource @testParam | Should -Be $true
                Assert-MockCalled -CommandName 'Get-TargetResource' -Exactly -Times 1 -Scope It
                Assert-MockCalled -CommandName 'Format-HotKeyString' -Exactly -Times 0 -Scope It
                Assert-MockCalled -CommandName 'ConvertFrom-HotKeyString' -Exactly -Times 0 -Scope It
            }

            It 'Return $true when the Ensure is Present and match all specified properties (all properties)' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                New-Item $PathExists -ItemType File -Force

                $testParam = @{
                    Ensure           = 'Present'
                    Path             = $PathExists
                    Target           = $Target
                    WorkingDirectory = 'Mock_WorkingDirectory'
                    Arguments        = 'Mock_Arguments'
                    Description      = 'Mock_Description'
                    Icon             = 'Mock_IconLocation'
                    HotKey           = 'Shift+B'
                    WindowStyle      = 'maximized'
                    AppUserModelID   = 'Mock_AppUserModelID'
                }

                Test-TargetResource @testParam | Should -Be $true
                Assert-MockCalled -CommandName 'Get-TargetResource' -Exactly -Times 1 -Scope It
                Assert-MockCalled -CommandName 'Format-HotKeyString' -Exactly -Times 0 -Scope It -ParameterFilter { $HotKey -eq 'Shift+B' }
                Assert-MockCalled -CommandName 'ConvertFrom-HotKeyString' -Exactly -Times 1 -Scope It -ParameterFilter { $HotKey -eq 'Shift+B' }
            }

            It 'Return $false when the Ensure is Present and does not match some specified properties (single property)' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $Target = 'C:\Windows\System32\hostname.exe'

                New-Item $PathExists -ItemType File -Force

                $testParam = @{
                    Ensure = 'Present'
                    Path   = $PathExists
                    Target = $Target
                }

                Test-TargetResource @testParam | Should -Be $false
                Assert-MockCalled -CommandName 'Get-TargetResource' -Exactly -Times 1 -Scope It
                Assert-MockCalled -CommandName 'Format-HotKeyString' -Exactly -Times 0 -Scope It
                Assert-MockCalled -CommandName 'ConvertFrom-HotKeyString' -Exactly -Times 0 -Scope It
            }

            It 'Return $false when the Ensure is Present and does not match some specified properties (two properties specified, one match, one does not.)' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                New-Item $PathExists -ItemType File -Force

                $testParam = @{
                    Ensure      = 'Present'
                    Path        = $PathExists
                    Target      = $Target
                    WindowStyle = 'minimized'
                }

                Test-TargetResource @testParam | Should -Be $false
                Assert-MockCalled -CommandName 'Get-TargetResource' -Exactly -Times 1 -Scope It
                Assert-MockCalled -CommandName 'Format-HotKeyString' -Exactly -Times 0 -Scope It
                Assert-MockCalled -CommandName 'ConvertFrom-HotKeyString' -Exactly -Times 0 -Scope It
            }

            It 'Return $false when the Ensure is Present and does not match some specified properties (all properties specified, only one does not match.)' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                New-Item $PathExists -ItemType File -Force

                $testParam = @{
                    Ensure           = 'Present'
                    Path             = $PathExists
                    Target           = $Target
                    WorkingDirectory = 'Mock_WorkingDirectory'
                    Arguments        = 'Mock_Arguments'
                    Description      = 'Mock_Description'
                    Icon             = 'Mock_IconLocation'
                    HotKey           = 'Shift+A'    # not match
                    WindowStyle      = 'maximized'
                    AppUserModelID   = 'Mock_AppUserModelID'
                }

                Test-TargetResource @testParam | Should -Be $false
                Assert-MockCalled -CommandName 'Get-TargetResource' -Exactly -Times 1 -Scope It
                Assert-MockCalled -CommandName 'Format-HotKeyString' -Exactly -Times 0 -Scope It -ParameterFilter { $HotKey -eq 'Shift+A' }
                Assert-MockCalled -CommandName 'ConvertFrom-HotKeyString' -Exactly -Times 1 -Scope It -ParameterFilter { $HotKey -eq 'Shift+A' }
            }
        }
    }
    #endregion Tests for Test-TargetResource


    #region Tests for Set-TargetResource
    Describe 'cShortcut/Set-TargetResource' -Tag 'Unit' {

        BeforeAll {
            $ErrorActionPreference = 'Stop'
        }

        Context 'Absent' {
            Mock Remove-Item {}
            Mock Remove-Item {} -Verifiable -ParameterFilter { $LiteralPath -eq (Join-Path $TestDrive '\Exist\shortcut.lnk') }

            It 'Remove file.' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                New-Item $PathExists -ItemType File -Force

                $setParam = @{
                    Ensure = 'Absent'
                    Path   = $PathExists
                    Target = $Target
                }

                { Set-TargetResource @setParam } | Should -Not -Throw
                Assert-VerifiableMock
            }

            It 'Remove file (without extension).' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $PathWithoutExtension = Join-Path $TestDrive '\Exist\shortcut'
                $Target = 'C:\Windows\System32\notepad.exe'

                New-Item $PathExists -ItemType File -Force

                $setParam = @{
                    Ensure = 'Absent'
                    Path   = $PathWithoutExtension
                    Target = $Target
                }

                { Set-TargetResource @setParam } | Should -Not -Throw
                Assert-VerifiableMock
            }
        }

        Context 'Present' {
            Mock Update-Shortcut {}
            Mock Update-Shortcut {} -ParameterFilter { $Path -eq (Join-Path $TestDrive '\Exist\shortcut.lnk') -and $Icon -eq 'Mock_Icon,0' }

            It 'Create shortcut.' {
                $PathExists = Join-Path $TestDrive '\Exist\shortcut.lnk'
                $Target = 'C:\Windows\System32\notepad.exe'

                $setParam = @{
                    Ensure = 'Present'
                    Path   = $PathExists
                    Target = $Target
                    Icon   = 'Mock_Icon'
                }

                { Set-TargetResource @setParam } | Should -Not -Throw
                Assert-MockCalled -CommandName 'Update-Shortcut' -Times 1 -Exactly -Scope It -ParameterFilter { $Path -eq (Join-Path $TestDrive '\Exist\shortcut.lnk') -and $Icon -eq 'Mock_Icon,0' }
            }

            It 'Create shortcut. (without Extension)' {
                $PathWithoutExtension = Join-Path $TestDrive '\Exist\shortcut'
                $Target = 'C:\Windows\System32\notepad.exe'

                $setParam = @{
                    Ensure = 'Present'
                    Path   = $PathWithoutExtension
                    Target = $Target
                    Icon   = 'Mock_Icon,0'
                }

                { Set-TargetResource @setParam } | Should -Not -Throw
                Assert-MockCalled -CommandName 'Update-Shortcut' -Times 1 -Exactly -Scope It -ParameterFilter { $Path -eq (Join-Path $TestDrive '\Exist\shortcut.lnk') -and $Icon -eq 'Mock_Icon,0' }
            }
        }
    }
    #endregion Tests for Set-TargetResource

    #region Tests for Get-Shortcut
    Describe 'cShortcut/Get-Shortcut' -Tag 'Unit' {

        BeforeAll {
            $ErrorActionPreference = 'Stop'
            Get-ChildItem $global:TestData | Copy-Item -Destination $TestDrive -Force -Recurse
        }

        It 'Get correct properties from existence shortcut file.' {
            $LnkFile = Join-Path $TestDrive 'NotePad.lnk'
            try {
                $Shortcut = Get-Shortcut -Path $LnkFile
                $Shortcut.GetType().FullName | Should -Be 'ShellLink'
                $Shortcut.TargetPath | Should -BeExactly '%windir%\system32\notepad.exe'
                $Shortcut.Description | Should -BeExactly '基本的なテキストの形式を使ったテキスト ファイルの表示と編集ができます。'
                $Shortcut.Arguments | Should -BeExactly 'C:\test.txt'
                $Shortcut.WorkingDirectory | Should -BeExactly '%HOMEDRIVE%%HOMEPATH%'
                $Shortcut.WindowStyle | Should -Be 3
                $Shortcut.Hotkey | Should -Be 0x0641
                $Shortcut.AppUserModelID | Should -BeExactly 'System.NotePad'
            }
            finally {
                if ($Shortcut -is [IDisposable]) {
                    $Shortcut.Dispose()
                    $Shortcut = $null
                }
            }
        }

        It 'Load shortcut file as Read Only' {
            $LnkFile = Join-Path $TestDrive 'NotePad.lnk'
            try {
                $Shortcut = Get-Shortcut -Path $LnkFile -ReadOnly
                $Shortcut.GetType().FullName | Should -Be 'ShellLink'
                { $Shortcut.TargetPath = 'C:\Windows\System32\HOSTNAME.EXE' } | Should -Throw
            }
            finally {
                if ($Shortcut -is [IDisposable]) {
                    $Shortcut.Dispose()
                    $Shortcut = $null
                }
            }
        }
    }
    # #endregion Tests for Get-Shortcut


    #region Tests for Format-HotKeyString
    Describe 'cShortcut/Format-HotKeyString' -Tag 'Unit' {

        BeforeAll {
            $ErrorActionPreference = 'Stop'
        }

        It 'Throw error if input string has less than 2 elements. (except Function keys)' {
            $HotKey = 'Ctrl'
            { Format-HotKeyString -HotKey $HotKey } | Should -Throw

            $HotKey = 'F'
            { Format-HotKeyString -HotKey $HotKey } | Should -Throw
        }

        It 'Throw error if input string has more than 5 elements.' {
            $HotKey = ('Ctrl', 'Shift', 'Alt', 'A', 'F8') -join '+'
            { Format-HotKeyString -HotKey $HotKey } | Should -Throw
        }

        It 'Throw error if the first element of an input is not modifier.  (except Function keys)' {
            $HotKey = ('A', 'Shift') -join '+'
            { Format-HotKeyString -HotKey $HotKey } | Should -Throw
        }

        It 'Accept single element if it is Function Keys. (F1 to F24)' {
            $HotKey = 'F1'
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly 'F1'

            $HotKey = 'F10'
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly 'F10'

            $HotKey = 'F24'
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly 'F24'
        }

        It 'Throw error if the specified F-Keys out of range. (F0 and F25)' {
            $HotKey = 'F0'
            { Format-HotKeyString -HotKey $HotKey } | Should -Throw

            $HotKey = 'F25'
            { Format-HotKeyString -HotKey $HotKey } | Should -Throw
        }

        It 'Returns empty string if an input is empty or white spaces string.' {
            $HotKey = [string]::Empty
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly ''

            $HotKey = '      '
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly ''
        }

        It 'Returns correct formatted string' {
            $HotKey = ('Alt', 'F12') -join '+'
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly 'Alt+F12'

            $HotKey = ('Ctrl', 'F12') -join '+'
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly 'Ctrl+F12'

            $HotKey = ('Ctrl', 'Alt', 'F12') -join '+'
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly 'Ctrl+Alt+F12'

            $HotKey = ('Ctrl', 'Shift', 'Alt', 'F12') -join '+'
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly 'Ctrl+Shift+Alt+F12'
        }

        It 'Returns correct formatted string with priority sorting' {
            $HotKey = ('Shift', 'Ctrl', 'Alt', 'F12') -join '+'
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly 'Ctrl+Shift+Alt+F12'

            $HotKey = ('Shift', 'F12', 'Alt', 'Ctrl') -join '+'
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly 'Ctrl+Shift+Alt+F12'
        }

        It 'Ignore casing' {
            $HotKey = ('cTrL', 'SHIFT', 'aLt', 'f12') -join '+'
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly 'cTrL+SHIFT+aLt+f12'
        }

        It 'Trim whitespaces' {
            $HotKey = ('  Ctrl', '   Shift', '   Alt   ', 'F12   ') -join '+'
            Format-HotKeyString -HotKey $HotKey | Should -BeExactly 'Ctrl+Shift+Alt+F12'
        }
    }
    #endregion Tests for Format-HotKeyString



    #region Tests for ConvertFrom-HotKeyString
    Describe 'cShortcut/ConvertFrom-HotKeyString' -Tag 'Unit' {

        BeforeAll {
            $ErrorActionPreference = 'Stop'
        }

        Mock Format-HotKeyString { return $HotKey }

        It 'Returns 0x0000 if the input is empty.' {
            $HotKey = [string]::Empty
            ConvertFrom-HotKeyString -HotKey $HotKey | Should -Be 0x0000
            Assert-MockCalled -CommandName 'Format-HotKeyString' -Times 0 -Exactly -Scope It
        }

        It 'Returns correct value. (Alt+F12)' {
            $HotKey = ('Alt', 'F12') -join '+'
            ConvertFrom-HotKeyString -HotKey $HotKey | Should -Be 0x047b
            Assert-MockCalled -CommandName 'Format-HotKeyString' -Times 1 -Exactly -Scope It
        }

        It 'Returns correct value. (Ctrl+Shift+F9)' {
            $HotKey = 'Ctrl+Shift+F9'
            ConvertFrom-HotKeyString -HotKey $HotKey | Should -Be 0x0378
            Assert-MockCalled -CommandName 'Format-HotKeyString' -Times 1 -Exactly -Scope It
        }

        It 'Returns correct value. (Ctrl+Alt+@)' {
            $HotKey = 'Ctrl+Alt+@'
            ConvertFrom-HotKeyString -HotKey $HotKey | Should -Be 0x06c0
            Assert-MockCalled -CommandName 'Format-HotKeyString' -Times 1 -Exactly -Scope It
        }

        It 'Throw exception if the keycode is not valid. (Ctrl+F999)' {
            $HotKey = 'Ctrl+F999'
            { ConvertFrom-HotKeyString -HotKey $HotKey } | Should -Throw
            Assert-MockCalled -CommandName 'Format-HotKeyString' -Times 1 -Exactly -Scope It
        }
    }
    #endregion Tests for ConvertFrom-HotKeyString

    #region Tests for ConvertTo-HotKeyString
    Describe 'cShortcut/ConvertTo-HotKeyString' -Tag 'Unit' {

        BeforeAll {
            $ErrorActionPreference = 'Stop'
        }

        Mock Format-HotKeyString { return $HotKey }

        It 'Returns empty string if an input code is 0' {
            $HotKeyCode = 0
            ConvertTo-HotKeyString -HotKeyCode $HotKeyCode | Should -BeNullOrEmpty
            Assert-MockCalled -CommandName 'Format-HotKeyString' -Times 0 -Exactly -Scope It
        }

        It 'Returns correct string. (Alt+F12)' {
            $HotKeyCode = 0x047b
            ConvertTo-HotKeyString -HotKeyCode $HotKeyCode | Should -BeExactly 'Alt+F12'
            Assert-MockCalled -CommandName 'Format-HotKeyString' -Times 1 -Exactly -Scope It
        }

        It 'Returns correct string. (F9)' {
            $HotKeyCode = 0x0078
            ConvertTo-HotKeyString -HotKeyCode $HotKeyCode | Should -BeExactly 'F9'
            Assert-MockCalled -CommandName 'Format-HotKeyString' -Times 1 -Exactly -Scope It
        }

        It 'Returns correct value. (Ctrl+Alt+@)' {
            # This test depends on the current keyboard layout.
            if ((Get-Culture).KeyboardLayoutID -ne 1041) {
                Set-ItResult -Skipped
            }
            else {
                $HotKeyCode = 0x06c0
                ConvertTo-HotKeyString -HotKeyCode $HotKeyCode | Should -BeExactly 'Ctrl+Alt+@'
                Assert-MockCalled -CommandName 'Format-HotKeyString' -Times 1 -Exactly -Scope It
            }
        }
    }
    #endregion Tests for ConvertTo-HotKeyString
}
#endregion End Testing
