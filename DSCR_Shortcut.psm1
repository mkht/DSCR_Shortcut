$modulePath = $PSScriptRoot
$subModulePath = @(
    '\DSCResources\cShortcut\cShortcut.psm1'
)

$subModulePath.ForEach( {
        Import-Module (Join-Path $modulePath $_)
    })
