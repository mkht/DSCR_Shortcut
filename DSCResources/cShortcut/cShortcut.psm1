Enum Ensure{
    Absent
    Present
}

Enum WindowStyle
{
    undefined = 0
    normal    = 1
    maximized = 3
    minimized = 7
}

function Get-TargetResource {
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = [Ensure]::Present,

        [parameter(Mandatory = $true)]
        [System.String]
        $Path,

        [parameter(Mandatory = $true)]
        [System.String]
        $Target,

        [parameter()]
        [System.String]
        $WorkingDirectory,

        [parameter()]
        [System.String]
        $Arguments,

        [parameter()]
        [System.String]
        $Description,

        [ValidateSet("normal", "maximized", "minimized")]
        [System.String]
        $WindowStyle = [WindowStyle]::normal
    )

    if (-not $Path.EndsWith('.lnk')) {
        Write-Verbose ("File extentison is not 'lnk'. Automatically add extension")
        $Path = $Path + '.lnk'
    }

    $Ensure = [Ensure]::Present

    # check file exists
    if (-not (Test-Path $Path)) {
        Write-Verbose 'File not found.'
        $Ensure = [Ensure]::Absent
    }
    else {
        $shortcut = Get-Shortcut $Path -ErrorAction Continue
    }
    $returnValue = @{
        Ensure           = $Ensure
        Path             = $Path
        Target           = $shortcut.TargetPath
        WorkingDirectory = $shortcut.WorkingDirectory
        Arguments        = $shortcut.Arguments
        Description        = $shortcut.Description
        WindowStyle      = [WindowStyle]::undefined
    }

    if($shortcut.WindowStyle -as [WindowStyle]){
        $returnValue.WindowStyle = [WindowStyle]$shortcut.WindowStyle
    }

    $returnValue
} # end of Get-TargetResource

function Set-TargetResource {
    [CmdletBinding()]
    param
    (
        [parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = [Ensure]::Present,

        [parameter(Mandatory = $true)]
        [System.String]
        $Path,

        [parameter(Mandatory = $true)]
        [System.String]
        $Target,

        [parameter()]
        [System.String]
        $WorkingDirectory,

        [parameter()]
        [System.String]
        $Arguments,

        [parameter()]
        [System.String]
        $Description,

        [ValidateSet("normal", "maximized", "minimized")]
        [System.String]
        $WindowStyle = [WindowStyle]::normal
    )

    if (-not $Path.EndsWith('.lnk')) {
        Write-Verbose ("File extentison is not 'lnk'. Automatically add extension")
        $Path = $Path + '.lnk'
    }

    # Ensure = "Absent"
    if ($Ensure -eq [Ensure]::Absent) {
        Write-Verbose ('Remove shortcut file "{0}"' -f $Path)
        Remove-Item $Path -Force
    }
    else {
        # Ensure = "Present"
        $arg = $PSBoundParameters
        $arg.Remove('Ensure')
        New-Shortcut @arg
    }

} # end of Set-TargetResource

function Test-TargetResource {
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure = [Ensure]::Present,

        [parameter(Mandatory = $true)]
        [System.String]
        $Path,

        [parameter(Mandatory = $true)]
        [System.String]
        $Target,

        [parameter()]
        [System.String]
        $WorkingDirectory,

        [parameter()]
        [System.String]
        $Arguments,

        [parameter()]
        [System.String]
        $Description,

        [ValidateSet("normal", "maximized", "minimized")]
        [System.String]
        $WindowStyle = [WindowStyle]::normal
    )

    <#  想定される状態パターンと返却するべき値
        1. ショートカットがあるべき(Present)
            1-A. ショートカットなし => 更新の必要あり($false)
            1-B. ショートカットはあるがプロパティが正しくない => 更新の必要あり($false)
            1-C. ショートカットあり、プロパティ一致 => 何もする必要なし($true)
        2. ショートカットはあるべきではない(Absent)
            2-A. ショートカットなし => 何もする必要なし($true)
            2-B. ショートカットあり => 削除の必要あり($false)
    #>

    # 拡張子つける
    if (-not $Path.EndsWith('.lnk')) {
        Write-Verbose ("File extentison is not 'lnk'. Automatically add extension")
        $Path = $Path + '.lnk'
    }

    $ReturnValue = $false

    switch ($Ensure) {
        'Absent' {
            # ファイルがなければ$true あれば$false
            $ReturnValue = (-not (Test-Path $Path -PathType Leaf))
        }
        'Present' {
            $Info = Get-TargetResource -Ensure $Ensure -Path $Path -Target $Target
            if ($Info.Ensure -eq [Ensure]::Absent) {
                $ReturnValue = $false
            }
            else {
                $ReturnValue = ($Info.Target -eq $Target)`
                 -and ($Info.WorkingDirectory -eq $WorkingDirectory)`
                  -and ($Info.Arguments -eq $Arguments)`
                   -and ($Info.Description -eq $Description)`
                    -and ($Info.WindowStyle -eq $WindowStyle)
            }
        }
    }
    Write-Verbose "Test returns $ReturnValue"
    return $ReturnValue
} # end of Test-TargetResource

function New-Shortcut {
    [CmdletBinding()]
    [OutputType([System.__ComObject])]
    param
    (
        # Set Target full path to create shortcut
        [parameter(
            Position = 0,
            Mandatory,
            ValueFromPipelineByPropertyName)]
        [Alias('Target')]
        [string]$TargetPath,

        # set file path to create shortcut. If the path not ends with '.lnk', extension will be add automatically.
        [parameter(
            Position = 1,
            Mandatory,
            ValueFromPipelineByPropertyName)]
        #[validateScript({Test-Path (Split-Path $_ -Parent)})]
        [string]$Path,

        # Set Description for shortcut.
        [parameter(
            ValueFromPipelineByPropertyName)]
        [Alias('Comment')]
        [string]$Description,

        # Set Arguments for shortcut.
        [parameter(
            ValueFromPipelineByPropertyName)]
        [string]$Arguments,

        # Set WorkingDirectory for shortcut.
        [parameter(
            ValueFromPipelineByPropertyName)]
        # [validateScript({Test-Path $_})]
        [string]$WorkingDirectory,

        # Set WindowStyle for shortcut.
        [parameter(
            ValueFromPipelineByPropertyName)]
        [ValidateSet('normal', 'maximized', 'minimized')]
        [string]$WindowStyle = [WindowStyle]::normal,

        # set if you want to show create shortcut result
        [switch]$PassThru
    )

    begin {
        $extension = ".lnk"
        $wsh = New-Object -ComObject Wscript.Shell
    }

    process {
        # set Path for Shortcut
        if (-not $Path.EndsWith('.lnk')) {
            $Path = $Path + $extension
        }

        if (-not (Test-Path (Split-Path $Path -Parent))) {
            Write-Verbose ("Create parent folder")
            New-Item -Path (Split-Path $Path -Parent) -ItemType Directory -Force -ErrorAction Stop
        }

        $fileName = Split-Path $Path -Leaf  # Filename of shortcut
        $Directory = Resolve-Path (Split-Path $Path -Parent) # Directory of shortcut
        $Path = Join-Path $Directory $fileName  # Fullpath of shortcut

        #Remove existing shortcut
        if (Test-Path $path) {
            Write-Verbose ("Remove existing shortcut file")
            Remove-Item $path -Force -ErrorAction SilentlyContinue
        }

        # Call Wscript to create Shortcut
        Write-Verbose ("Trying to create Shortcut for name '{0}'" -f $path)
        try {
            $shortCut = $wsh.CreateShortCut($path)
            $shortCut.TargetPath = $TargetPath
            $shortCut.Description = $Description
            $shortCut.WindowStyle = [int][WindowStyle]$WindowStyle
            $shortCut.Arguments = $Arguments
            $shortCut.WorkingDirectory = $WorkingDirectory
            $shortCut.Save()
            Write-Verbose ('Shortcut file created successfully')
        }
        catch [Exception] {
            Write-Error $_.Exception
        }

        if ($PSBoundParameters.PassThru) {
            $shortCut
        }
    }

    end {}
}

function Get-Shortcut {
    [CmdletBinding()]
    [OutputType([System.__ComObject])]
    param
    (
        # Path of shortcut file
        [parameter(
            Position = 0,
            Mandatory,
            ValueFromPipeline)]
        [validateScript( {$_ | % {Test-Path $_}})]
        [string[]]$Path
    )
    begin {
        $wsh = New-Object -ComObject Wscript.Shell
    }

    Process {
        $Path.ForEach( {
                $fullPath = Resolve-Path $_
                Write-Verbose ('Trying to get file properties from "{0}"' -f $fullPath)
                $wsh.CreateShortcut($fullPath.Path)
            })
    }

    End {}
}

Export-ModuleMember -Function *-TargetResource
