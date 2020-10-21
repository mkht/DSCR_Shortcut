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
                    WindowStyle      = 'maximized'
                    AppUserModelID   = 'Mock_AppUserModelID'
                }
            }

            Mock Format-HotKeyString { return 'Shift+A' } -ParameterFilter { $HotKey -eq 'Shift+A' }
            Mock Format-HotKeyString { return 'Shift+B' } -ParameterFilter { $HotKey -eq 'Shift+B' }

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
                Assert-MockCalled -CommandName 'Format-HotKeyString' -Exactly -Times 1 -Scope It -ParameterFilter { $HotKey -eq 'Shift+B' }
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
                Assert-MockCalled -CommandName 'Format-HotKeyString' -Exactly -Times 1 -Scope It -ParameterFilter { $HotKey -eq 'Shift+A' }
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
}
#endregion End Testing
