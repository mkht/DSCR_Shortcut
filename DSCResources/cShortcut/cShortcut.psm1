# Import ShellLink class
$ShellLinkPath = Join-Path $PSScriptRoot '..\..\Libs\ShellLink\ShellLink.cs'
if (Test-Path -LiteralPath $ShellLinkPath -PathType Leaf) {
    Add-Type -TypeDefinition (Get-Content -LiteralPath $ShellLinkPath -Raw -Encoding UTF8) -Language 'CSharp' -ErrorAction Stop
}

Enum Ensure {
    Absent
    Present
}

Enum WindowStyle {
    undefined = 0
    normal = 1
    maximized = 3
    minimized = 7
}

function Get-TargetResource {
    [CmdletBinding()]
    [OutputType([Hashtable])]
    param
    (
        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [string]
        $Ensure = [Ensure]::Present,

        [Parameter(Mandatory)]
        [string]
        $Path,

        [Parameter(Mandatory)]
        [string]
        $Target,

        [Parameter()]
        [string]
        $WorkingDirectory,

        [Parameter()]
        [string]
        $Arguments,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [string]
        $Icon,

        [Parameter()]
        [string]
        $HotKey,

        [ValidateSet("normal", "maximized", "minimized")]
        [string]
        $WindowStyle = [WindowStyle]::normal
    )

    if (-not $Path.EndsWith('.lnk')) {
        Write-Verbose ("File extension is not 'lnk'. Automatically add extension")
        $Path = $Path + '.lnk'
    }

    $Ensure = [Ensure]::Present

    # check file exists
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Verbose 'File not found.'
        $Ensure = [Ensure]::Absent
    }
    else {
        $shortcut = Get-Shortcut -Path $Path -ErrorAction Continue
    }
    $returnValue = @{
        Ensure           = $Ensure
        Path             = $Path
        Target           = $shortcut.TargetPath
        WorkingDirectory = $shortcut.WorkingDirectory
        Arguments        = $shortcut.Arguments
        Description      = $shortcut.Description
        Icon             = $shortcut.IconLocation
        HotKey           = $shortcut.Hotkey
        WindowStyle      = [WindowStyle]::undefined
    }

    if ($shortcut.WindowStyle -as [WindowStyle]) {
        $returnValue.WindowStyle = [WindowStyle]$shortcut.WindowStyle
    }

    $returnValue
} # end of Get-TargetResource


function Set-TargetResource {
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [string]
        $Ensure = [Ensure]::Present,

        [Parameter(Mandatory)]
        [string]
        $Path,

        [Parameter(Mandatory)]
        [string]
        $Target,

        [Parameter()]
        [string]
        $WorkingDirectory,

        [Parameter()]
        [string]
        $Arguments,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [string]
        $Icon,

        [Parameter()]
        [string]
        $HotKey,

        [ValidateSet("normal", "maximized", "minimized")]
        [string]
        $WindowStyle = [WindowStyle]::normal
    )

    if (-not $Path.EndsWith('.lnk')) {
        Write-Verbose ("File extension is not 'lnk'. Automatically add extension")
        $Path = $Path + '.lnk'
    }

    if ($Icon -and ($Icon -notmatch ',\d+$')) {
        $Icon = $Icon + ',0'
    }

    # Ensure = "Absent"
    if ($Ensure -eq [Ensure]::Absent) {
        Write-Verbose ('Remove shortcut file "{0}"' -f $Path)
        Remove-Item -LiteralPath $Path -Force
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
    [OutputType([bool])]
    param
    (
        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [string]
        $Ensure = [Ensure]::Present,

        [Parameter(Mandatory)]
        [string]
        $Path,

        [Parameter(Mandatory)]
        [string]
        $Target,

        [Parameter()]
        [string]
        $WorkingDirectory,

        [Parameter()]
        [string]
        $Arguments,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [string]
        $Icon = ',0',

        [Parameter()]
        [string]
        $HotKey,

        [ValidateSet("normal", "maximized", "minimized")]
        [string]
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
        Write-Verbose ("File extension is not 'lnk'. Automatically add extension")
        $Path = $Path + '.lnk'
    }

    if ($Icon -and ($Icon -notmatch ',\d+$')) {
        $Icon = $Icon + ',0'
    }

    # HotKey文字列組み立て
    if ($HotKey) {
        $HotKeyStr = Format-HotKeyString $HotKey
    }
    else {
        $HotKeyStr = [string]::Empty
    }

    $ReturnValue = $false
    switch ($Ensure) {
        'Absent' {
            # ファイルがなければ$true あれば$false
            $ReturnValue = (-not (Test-Path -LiteralPath $Path -PathType Leaf))
        }
        'Present' {
            $Info = Get-TargetResource -Ensure $Ensure -Path $Path -Target $Target
            if ($Info.Ensure -eq [Ensure]::Absent) {
                $ReturnValue = $false
            }
            else {
                # Tests whether the shortcut property is the same as the specified parameter.
                $NotMatched = @()
                if ($Info.Target -ne [System.Environment]::ExpandEnvironmentVariables($Target)) {
                    $NotMatched += 'Target'
                }

                if ($PSBoundParameters.ContainsKey('WorkingDirectory') -and ($Info.WorkingDirectory -ne $WorkingDirectory)) {
                    $NotMatched += 'WorkingDirectory'
                }

                if ($PSBoundParameters.ContainsKey('Arguments') -and ($Info.Arguments -ne $Arguments)) {
                    $NotMatched += 'Arguments'
                }

                if ($PSBoundParameters.ContainsKey('Description') -and ($Info.Description -ne $Description)) {
                    $NotMatched += 'Description'
                }

                if ($PSBoundParameters.ContainsKey('Icon') -and ($Info.Icon -ne $Icon)) {
                    $NotMatched += 'Icon'
                }

                if ($PSBoundParameters.ContainsKey('HotKey') -and ($Info.HotKey -ne $HotKey)) {
                    $NotMatched += 'HotKey'
                }

                if ($Info.WindowStyle -ne $WindowStyle) {
                    $NotMatched += 'WindowStyle'
                }

                $ReturnValue = ($NotMatched.Count -eq 0)
                if (-not $ReturnValue) {
                    $NotMatched | ForEach-Object {
                        Write-Verbose ('{0} property is not matched!' -f $_)
                    }
                }
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
        [Parameter(
            Position = 0,
            Mandatory,
            ValueFromPipelineByPropertyName)]
        [Alias('Target')]
        [string]$TargetPath,

        # set file path to create shortcut. If the path not ends with '.lnk', extension will be add automatically.
        [Parameter(
            Position = 1,
            Mandatory,
            ValueFromPipelineByPropertyName)]
        #[validateScript({Test-Path (Split-Path $_ -Parent)})]
        [string]$Path,

        # Set Description for shortcut.
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Comment')]
        [string]$Description,

        # Set Arguments for shortcut.
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$Arguments,

        # Set WorkingDirectory for shortcut.
        [Parameter(ValueFromPipelineByPropertyName)]
        # [validateScript({Test-Path $_})]
        [string]$WorkingDirectory,

        # Set IconLocation for shortcut.
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$Icon,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $HotKey,

        # Set WindowStyle for shortcut.
        [Parameter(ValueFromPipelineByPropertyName)]
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

        if ($HotKey) {
            $HotKeyStr = Format-HotKeyString $HotKey
        }

        if (-not (Test-Path -LiteralPath (Split-Path $Path -Parent))) {
            Write-Verbose ("Create parent folder")
            New-Item -Path (Split-Path $Path -Parent) -ItemType Directory -Force -ErrorAction Stop
        }

        $fileName = Split-Path $Path -Leaf  # Filename of shortcut
        $Directory = Resolve-Path -Path (Split-Path $Path -Parent) # Directory of shortcut
        $Path = Join-Path $Directory $fileName  # Fullpath of shortcut

        #Remove existing shortcut
        if (Test-Path -LiteralPath $path) {
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
            if ($PSBoundParameters.ContainsKey('Icon')) {
                $shortCut.IconLocation = $Icon
            }
            if ($HotKeyStr) {
                $shortCut.Hotkey = $HotKeyStr
            }
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
    [OutputType([ShellLink])]
    param
    (
        # Path of shortcut files
        [Parameter(Position = 0, Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [ValidateScript( { $_ | ForEach-Object { Test-Path -LiteralPath $_ } })]
        [string]$Path,

        [switch]$ReadOnly
    )

    Begin {
        if ($ReadOnly) {
            [int]$flag = 0x00000000 #STGM_READ
        }
        else {
            [int]$flag = 0x00000002 #STGM_READWRITE
        }
    }

    Process {
        try {
            $Shortcut = New-Object -TypeName ShellLink
            $Shortcut.Load($Path, $flag)
            return $Shortcut
        }
        catch {
            if ($Shortcut -is [IDisposable]) {
                $Shortcut.Dispose()
                $Shortcut = $null
            }

            Write-Error -Exception $_.Exception
            return $null
        }
    }
}


function Format-HotKeyString {
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [Parameter(Mandatory, Position = 0)]
        [string[]]$HotKeyArray
    )

    $HotKeyArray = $HotKey.split('+').Trim()
    if ($HotKeyArray.Count -notin (2..4)) {
        #最短で修飾+キーの2要素、最長でAlt+Ctrl+Shift+キーの4要素
        Write-Error ('HotKey is not valid format.')
    }
    elseif ($HotKeyArray[0] -notmatch '^(Ctrl|Alt|Shift)$') {
        #修飾キーから始まっていないとダメ
        Write-Error ('HotKey is not valid format.')
    }
    else {
        #優先順位付きソート
        $sort = $HotKeyArray | ForEach-Object {
            switch ($_) {
                'Alt' { 1 }
                'Ctrl' { 2 }
                'Shift' { 3 }
                Default { 4 }
            }
        }
        [Array]::Sort($sort, $HotKeyArray)
        $HotKeyArray -join '+'
    }
}


Export-ModuleMember -Function *-TargetResource
