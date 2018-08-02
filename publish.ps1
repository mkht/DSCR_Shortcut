Param (
    [Parameter(Mandatory = $true)]
    [string]$NugetApiKey,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ExcludeDir = @('.git', '.vscode'),

    [switch]$WhatIf
)

$ModuleDir = $PSScriptRoot
$ModuleName = Split-Path $ModuleDir -Leaf
$Destination = Join-Path $env:TEMP $ModuleName

if (Test-Path $Destination) {
    Remove-Item $Destination -Force -Recurse -ErrorAction Stop
}

if ($ExcludeDir -notcontains '.git') {
    $ExcludeDir += '.git'
}

try {
    robocopy $ModuleDir $Destination /MIR /XD ($ExcludeDir -join ' ') /XF "publish.ps1" >$null

    Set-Location $Destination
    Publish-Module -Path ./ -NuGetApiKey $NugetApiKey -Verbose -WhatIf:$WhatIf
}
finally {
    Set-Location $ModuleDir
    Remove-Item $Destination -Force -Recurse -ErrorAction Continue
}
