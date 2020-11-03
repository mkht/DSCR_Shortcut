#Requires #Requires -Module @{ ModuleName = 'Pester'; RequiredVersion = '4.10.1' }

$script:dscModuleName = 'DSCR_Shortcut'
$script:dscResourceName = 'cShortcut'
$script:dscModuleRoot = Resolve-Path '..\..'

Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath '..\TestHelper.psm1')

Initialize-TestEnvironment -ModuleName $script:dscModuleName -ModuleRoot $script:dscModuleRoot

try {
    Describe 'cShortcut Integration Tests' {

        Describe "$($script:dscResourceName)_Integration" {
            $configFile = Join-Path -Path $PSScriptRoot -ChildPath "$($script:dscResourceName).config.ps1"
            . $configFile -Verbose -ErrorAction Stop

            BeforeAll {
                Remove-Item -Path 'C:\Windows\Temp\cShortcut_Integration' -Force -Recurse -ErrorAction SilentlyContinue
            }

            It 'Should compile and apply configuration without throwing' {
                {
                    & "$($script:dscResourceName)_Config" -OutputPath $TestDrive
                    Start-DscConfiguration `
                        -Path $TestDrive `
                        -ComputerName localhost `
                        -Wait `
                        -Verbose `
                        -Force
                } | Should -Not -Throw
            }

            It 'Test-DscConfiguration should returns True.' {
                $testResult = Test-DscConfiguration -Path $TestDrive -ComputerName localhost
                $testResult | Should -Be $true
            }

            It 'should be able to call Get-DscConfiguration without throwing' {
                { Get-DscConfiguration -Verbose -ErrorAction Stop } | Should -Not -Throw
            }

            It 'Should have set the resource and all the parameters should match' {
                $current = Get-DscConfiguration | Where-Object -FilterScript {
                    $_.ConfigurationName -eq "$($script:dscResourceName)_Config"
                }
                $current.Target           | Should -Be $TestShortcut.Target
                $current.Arguments        | Should -Be $TestShortcut.Arguments
                $current.WindowStyle      | Should -Be $TestShortcut.WindowStyle
                $current.WorkingDirectory | Should -Be $TestShortcut.WorkingDirectory
                $current.Description      | Should -Be $TestShortcut.Description
                $current.Icon             | Should -Be $TestShortcut.Icon
                $current.HotKeyCode       | Should -Be $TestShortcut.HotKeyCode
                $current.AppUserModelID   | Should -Be $TestShortcut.AppUserModelID
            }
        }
    }
}
finally {
    Remove-Item -Path 'C:\Windows\Temp\cShortcut_Integration' -Force -Recurse -ErrorAction SilentlyContinue
    Restore-TestEnvironment -ModuleName $script:dscModuleName -ModuleRoot $script:dscModuleRoot
}
