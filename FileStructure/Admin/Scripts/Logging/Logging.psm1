#requires -Version 2

$LogFilePreference = $null

function ShouldPerformLogging
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Path,

        [System.String]
        $ActionPreference
    )

    if ([System.String]::IsNullOrEmpty($Path) -or $ActionPreference -eq 'SilentlyContinue' -or $ActionPreference -eq 'Ignore')
    {
        return $false
    }

    try
    {
        $folder = Split-Path $Path -Parent

        if ($folder -ne '' -and -not (Test-Path -Path $folder))
        {
            $null = New-Item -Path $folder -ItemType Directory -ErrorAction Stop
        }

        return $true
    }
    catch
    {
        try
        {
            $cmd = Get-Command -Name Write-Warning -CommandType Cmdlet
            & $cmd ($_ | Out-String)
        }
        catch { }

        return $false
    }
}

function PrependString
{
    [CmdletBinding()]
    param (
        [System.String]
        $Line,

        [System.String]
        $Flag
    )

    if ($null -eq $Line)
    {
        $Line = [System.String]::Empty
    }

    if ($null -eq $Flag)
    {
        $Flag = [System.String]::Empty
    }

    if ($Line.Trim() -ne '')
    {
        $prependString = "$(Get-Date -Format r) - "
        if (-not [System.String]::IsNullOrEmpty($Flag))
        {
            $prependString += "$Flag "
        }

        Write-Output $prependString
    }
}

function Write-DebugLog
{
    <#
    .Synopsis
       Proxy function for Write-Debug.  Optionally, also directs the debug output to a log file.
    .DESCRIPTION
       Has the same definition as Write-Debug, with the addition of a -LogFile parameter.  If this
       argument has a value, it is treated as a file path, and the function will attempt to write
       the debug output to that file as well (including creating the parent directory, if it doesn't
       already exist).  If the path is malformed or the user does not have permission to create or
       write to the file, New-Item and Add-Content will send errors back through the output stream.

       Non-blank lines in the log file are automatically prepended with a culture-invariant date
       and time, and with the text [D] to indicate this output came from the debug stream.
    .PARAMETER LogFile
       Specifies the full path to the log file.  If this value is not specified, it will default to
       the variable $LogFilePreference, which is provided for the user's convenience in redirecting
       output from all of the Write-*Log functions to the same file.
    .LINK
       Write-Debug
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position=0, ValueFromPipeline = $true)]
        [Alias('Msg')]
        [AllowEmptyString()]
        [System.String]
        $Message,

        [System.String]
        $LogFile = $null,

        [System.Management.Automation.ScriptBlock]
        $Prepend = { PrependString -Line $args[0] -Flag '[D]' }
    )

    begin
    {
        try 
        {
            if ($PSBoundParameters.ContainsKey('LogFile'))
            {
                $_logFile = $LogFile
                $null = $PSBoundParameters.Remove('LogFile')
            }
            else
            {
                $_logFile = $PSCmdlet.SessionState.PSVariable.GetValue("LogFilePreference")
            } 

            if ($PSBoundParameters.ContainsKey('Prepend'))
            {
                $null = $PSBoundParameters.Remove('Prepend')
            }

            $outBuffer = $null

            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
            {
                $PSBoundParameters['OutBuffer'] = 1
            }

            if (-not $PSBoundParameters.ContainsKey('Verbose'))
            {
                $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue("VerbosePreference")
            }

            if (-not $PSBoundParameters.ContainsKey('Debug'))
            {
                $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue("DebugPreference")
            }

            if (-not $PSBoundParameters.ContainsKey('WarningAction'))
            {
                $WarningPreference = $PSCmdlet.SessionState.PSVariable.GetValue("WarningPreference")
            }

            if (-not $PSBoundParameters.ContainsKey('ErrorAction'))
            {
                $ErrorActionPreference = $PSCmdlet.SessionState.PSVariable.GetValue("ErrorActionPreference")
            }

            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Write-Debug', [System.Management.Automation.CommandTypes]::Cmdlet)
            $scriptCmd = {& $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch
        {
            throw
        }
    }

    process
    {
        if (ShouldPerformLogging -Path $_logFile -ActionPreference $DebugPreference)
        {
            foreach ($line in $Message -split '\r?\n')
            {
                if ($null -ne $Prepend)
                {
                    $results = $Prepend.Invoke($line)
                    if ($results.Count -gt 0 -and ($prependString = $results[0]) -is [System.String])
                    {
                        $line = "${prependString}${line}"
                    }
                }

                Add-Content -Path $_logFile -Value $line
            }
        }

        try
        {
            $steppablePipeline.Process($_)
        }
        catch
        {
            throw
        }
    }

    end
    {
        try
        {
            $steppablePipeline.End()
        }
        catch
        {
            throw
        }
    }
}

function Write-VerboseLog {
    <#
    .Synopsis
       Proxy function for Write-Verbose.  Optionally, also directs the verbose output to a log file.
    .DESCRIPTION
       Has the same definition as Write-Verbose, with the addition of a -LogFile parameter.  If this
       argument has a value, it is treated as a file path, and the function will attempt to write
       the debug output to that file as well (including creating the parent directory, if it doesn't
       already exist).  If the path is malformed or the user does not have permission to create or
       write to the file, New-Item and Add-Content will send errors back through the output stream.

       Non-blank lines in the log file are automatically prepended with a culture-invariant date
       and time, and with the text [V] to indicate this output came from the verbose stream.
    .PARAMETER LogFile
       Specifies the full path to the log file.  If this value is not specified, it will default to
       the variable $LogFilePreference, which is provided for the user's convenience in redirecting
       output from all of the Write-*Log functions to the same file.
    .LINK
       Write-Verbose
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position=0, ValueFromPipeline = $true)]
        [Alias('Msg')]
        [AllowEmptyString()]
        [System.String]
        $Message,

        [System.String]
        $LogFile = $null,

        [System.Management.Automation.ScriptBlock]
        $Prepend = { PrependString -Line $args[0] -Flag '[V]' }
    )

    begin
    {
        try 
        {
            if ($PSBoundParameters.ContainsKey('LogFile'))
            {
                $_logFile = $LogFile
                $null = $PSBoundParameters.Remove('LogFile')
            }
            else
            {
                $_logFile = $PSCmdlet.SessionState.PSVariable.GetValue("LogFilePreference")
            } 

            if ($PSBoundParameters.ContainsKey('Prepend'))
            {
                $null = $PSBoundParameters.Remove('Prepend')
            }

            $outBuffer = $null

            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
            {
                $PSBoundParameters['OutBuffer'] = 1
            }

            if (-not $PSBoundParameters.ContainsKey('Verbose'))
            {
                $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue("VerbosePreference")
            }

            if (-not $PSBoundParameters.ContainsKey('Debug'))
            {
                $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue("DebugPreference")
            }

            if (-not $PSBoundParameters.ContainsKey('WarningAction'))
            {
                $WarningPreference = $PSCmdlet.SessionState.PSVariable.GetValue("WarningPreference")
            }

            if (-not $PSBoundParameters.ContainsKey('ErrorAction'))
            {
                $ErrorActionPreference = $PSCmdlet.SessionState.PSVariable.GetValue("ErrorActionPreference")
            }

            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Write-Verbose', [System.Management.Automation.CommandTypes]::Cmdlet)
            $scriptCmd = {& $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch
        {
            throw
        }
    }

    process
    {
        if (ShouldPerformLogging -Path $_logFile -ActionPreference $VerbosePreference)
        {
            foreach ($line in $Message -split '\r?\n')
            {
                if ($null -ne $Prepend)
                {
                    $results = $Prepend.Invoke($line)
                    if ($results.Count -gt 0 -and ($prependString = $results[0]) -is [System.String])
                    {
                        $line = "${prependString}${line}"
                    }
                }

                Add-Content -Path $_logFile -Value $line 
            }
        }

        try
        {
            $steppablePipeline.Process($_)
        }
        catch
        {
            throw
        }
    }

    end
    {
        try
        {
            $steppablePipeline.End()
        }
        catch
        {
            throw
        }
    }
}

function Write-WarningLog
{
    <#
    .Synopsis
       Proxy function for Write-Warning.  Optionally, also directs the warning output to a log file.
    .DESCRIPTION
       Has the same definition as Write-Warning, with the addition of a -LogFile parameter.  If this
       argument has a value, it is treated as a file path, and the function will attempt to write
       the debug output to that file as well (including creating the parent directory, if it doesn't
       already exist).  If the path is malformed or the user does not have permission to create or
       write to the file, New-Item and Add-Content will send errors back through the output stream.

       Non-blank lines in the log file are automatically prepended with a culture-invariant date
       and time, and with the text [W] to indicate this output came from the warning stream.
    .PARAMETER LogFile
       Specifies the full path to the log file.  If this value is not specified, it will default to
       the variable $LogFilePreference, which is provided for the user's convenience in redirecting
       output from all of the Write-*Log functions to the same file.
    .LINK
       Write-Warning
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Alias('Msg')]
        [AllowEmptyString()]
        [System.String]
        $Message,

        [System.String]
        $LogFile = $null,

        [System.Management.Automation.ScriptBlock]
        $Prepend = { PrependString -Line $args[0] -Flag '[W]' }
    )

    begin
    {
        try 
        {
            if ($PSBoundParameters.ContainsKey('LogFile'))
            {
                $_logFile = $LogFile
                $null = $PSBoundParameters.Remove('LogFile')
            }
            else
            {
                $_logFile = $PSCmdlet.SessionState.PSVariable.GetValue("LogFilePreference")
            } 

            if ($PSBoundParameters.ContainsKey('Prepend'))
            {
                $null = $PSBoundParameters.Remove('Prepend')
            }

            $outBuffer = $null

            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
            {
                $PSBoundParameters['OutBuffer'] = 1
            }

            if (-not $PSBoundParameters.ContainsKey('Verbose'))
            {
                $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue("VerbosePreference")
            }

            if (-not $PSBoundParameters.ContainsKey('Debug'))
            {
                $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue("DebugPreference")
            }

            if (-not $PSBoundParameters.ContainsKey('WarningAction'))
            {
                $WarningPreference = $PSCmdlet.SessionState.PSVariable.GetValue("WarningPreference")
            }

            if (-not $PSBoundParameters.ContainsKey('ErrorAction'))
            {
                $ErrorActionPreference = $PSCmdlet.SessionState.PSVariable.GetValue("ErrorActionPreference")
            }

            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Write-Warning', [System.Management.Automation.CommandTypes]::Cmdlet)
            $scriptCmd = {& $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch
        {
            throw
        }
    }

    process
    {
        if (ShouldPerformLogging -Path $_logFile -ActionPreference $WarningPreference)
        {
            foreach ($line in $Message -split '\r?\n')
            {
                if ($null -ne $Prepend)
                {
                    $results = $Prepend.Invoke($line)
                    if ($results.Count -gt 0 -and ($prependString = $results[0]) -is [System.String])
                    {
                        $line = "${prependString}${line}"
                    }
                }

                Add-Content -Path $_logFile -Value $line 
            }
        }

        try
        {
            $steppablePipeline.Process($_)
        }
        catch
        {
            throw
        }
    }

    end
    {
        try
        {
            $steppablePipeline.End()
        }
        catch
        {
            throw
        }
    }
}

function Write-ErrorLog
{
    <#
    .Synopsis
       Proxy function for Write-Error.  Optionally, also directs the error output to a log file.
    .DESCRIPTION
       Has the same definition as Write-Error, with the addition of a -LogFile parameter.  If this
       argument has a value, it is treated as a file path, and the function will attempt to write
       the debug output to that file as well (including creating the parent directory, if it doesn't
       already exist).  If the path is malformed or the user does not have permission to create or
       write to the file, New-Item and Add-Content will send errors back through the output stream.

       Non-blank lines in the log file are automatically prepended with a culture-invariant date
       and time, and with the text [E] to indicate this output came from the error stream.
    .PARAMETER LogFile
       Specifies the full path to the log file.  If this value is not specified, it will default to
       the variable $LogFilePreference, which is provided for the user's convenience in redirecting
       output from all of the Write-*Log functions to the same file.
    .LINK
       Write-Error
    #>

    [CmdletBinding(DefaultParameterSetName='NoException')]
    param(
        [Parameter(ParameterSetName = 'WithException', Mandatory = $true)]
        [System.Exception]
        $Exception,

        [Parameter(ParameterSetName = 'NoException', Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Parameter(ParameterSetName = 'WithException')]
        [Alias('Msg')]
        [AllowNull()]
        [AllowEmptyString()]
        [System.String]
        $Message,

        [Parameter(ParameterSetName = 'ErrorRecord', Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]
        $ErrorRecord,

        [Parameter(ParameterSetName = 'NoException')]
        [Parameter(ParameterSetName = 'WithException')]
        [System.Management.Automation.ErrorCategory]
        $Category,

        [Parameter(ParameterSetName = 'WithException')]
        [Parameter(ParameterSetName = 'NoException')]
        [System.String]
        $ErrorId,

        [Parameter(ParameterSetName = 'NoException')]
        [Parameter(ParameterSetName = 'WithException')]
        [System.Object]
        $TargetObject,

        [System.String]
        $RecommendedAction,

        [Alias('Activity')]
        [System.String]
        $CategoryActivity,

        [Alias('Reason')]
        [System.String]
        $CategoryReason,

        [Alias('TargetName')]
        [System.String]
        $CategoryTargetName,

        [Alias('TargetType')]
        [System.String]
        $CategoryTargetType,

        [System.String]
        $LogFile = $null,

        [System.Management.Automation.ScriptBlock]
        $Prepend = { PrependString -Line $args[0] -Flag '[E]' }
    )

    begin
    {
        try 
        {
            if ($PSBoundParameters.ContainsKey('LogFile'))
            {
                $_logFile = $LogFile
                $null = $PSBoundParameters.Remove('LogFile')
            }
            else
            {
                $_logFile = $PSCmdlet.SessionState.PSVariable.GetValue("LogFilePreference")
            } 

            if ($PSBoundParameters.ContainsKey('Prepend'))
            {
                $null = $PSBoundParameters.Remove('Prepend')
            }

            $outBuffer = $null

            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
            {
                $PSBoundParameters['OutBuffer'] = 1
            }

            if (-not $PSBoundParameters.ContainsKey('Verbose'))
            {
                $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue("VerbosePreference")
            }

            if (-not $PSBoundParameters.ContainsKey('Debug'))
            {
                $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue("DebugPreference")
            }

            if (-not $PSBoundParameters.ContainsKey('WarningAction'))
            {
                $WarningPreference = $PSCmdlet.SessionState.PSVariable.GetValue("WarningPreference")
            }

            if (-not $PSBoundParameters.ContainsKey('ErrorAction'))
            {
                $ErrorActionPreference = $PSCmdlet.SessionState.PSVariable.GetValue("ErrorActionPreference")
            }

            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Write-Error', [System.Management.Automation.CommandTypes]::Cmdlet)
            $scriptCmd = {& $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch
        {
            throw
        }
    }

    process
    {
        if (ShouldPerformLogging -Path $_logFile -ActionPreference $ErrorActionPreference)
        {
            $item = ''
            switch ($PSCmdlet.ParameterSetName)
            {
                'ErrorRecord'   { $item = $ErrorRecord }
                'WithException' { $item = $Exception   }
                default         { $item = $Message     }
            }

            foreach ($line in ($item | Out-String -Stream))
            {
                if ($null -ne $Prepend)
                {
                    $results = $Prepend.Invoke($line)
                    if ($results.Count -gt 0 -and ($prependString = $results[0]) -is [System.String])
                    {
                        $line = "${prependString}${line}"
                    }
                }

                Add-Content -Path $_logFile -Value $line
            }
        }

        try
        {
            $steppablePipeline.Process($_)
        }
        catch
        {
            throw
        }
    }

    end
    {
        try
        {
            $steppablePipeline.End()
        }
        catch
        {
            throw
        }
    }
}

function Write-HostLog
{
    <#
    .Synopsis
       Proxy function for Write-Host.  Optionally, also directs the output to a log file.
    .DESCRIPTION
       Has the same definition as Write-Host, with the addition of a -LogFile parameter.  If this
       argument has a value, it is treated as a file path, and the function will attempt to write
       the output to that file as well (including creating the parent directory, if it doesn't
       already exist).  If the path is malformed or the user does not have permission to create or
       write to the file, New-Item and Add-Content will send errors back through the output stream.

       Non-blank lines in the log file are automatically prepended with a culture-invariant date
       and time.
    .PARAMETER LogFile
       Specifies the full path to the log file.  If this value is not specified, it will default to
       the variable $LogFilePreference, which is provided for the user's convenience in redirecting
       output from all of the Write-*Log functions to the same file.
    .NOTES
       unlike Write-Host, this function defaults the value of the -Separator parameter to
       "`r`n".  This is to make the console output consistent with what is sent to the log file,
       where array elements are always written to separate lines (regardless of the value of the
       -Separator parameter;  if that argument is specified, it just gets passed on to Write-Host).
    .LINK
       Write-Host
    #>

    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [System.Object]
        $Object,

        [Switch]
        $NoNewline,
        
        [System.Object]
        $Separator = "`r`n",

        [System.ConsoleColor]
        $ForegroundColor,

        [System.ConsoleColor]
        $BackgroundColor,

        [System.String]
        $LogFile = $null,

        [System.Management.Automation.ScriptBlock]
        $Prepend = { PrependString -Line $args[0] }
    )

    begin
    {
        try 
        {
            if ($PSBoundParameters.ContainsKey('LogFile'))
            {
                $_logFile = $LogFile
                $null = $PSBoundParameters.Remove('LogFile')
            }
            else
            {
                $_logFile = $PSCmdlet.SessionState.PSVariable.GetValue("LogFilePreference")
            } 

            if ($PSBoundParameters.ContainsKey('Prepend'))
            {
                $null = $PSBoundParameters.Remove('Prepend')
            }

            $outBuffer = $null

            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
            {
                $PSBoundParameters['OutBuffer'] = 1
            }

            if (-not $PSBoundParameters.ContainsKey('Verbose'))
            {
                $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue("VerbosePreference")
            }

            if (-not $PSBoundParameters.ContainsKey('Debug'))
            {
                $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue("DebugPreference")
            }

            if (-not $PSBoundParameters.ContainsKey('WarningAction'))
            {
                $WarningPreference = $PSCmdlet.SessionState.PSVariable.GetValue("WarningPreference")
            }

            if (-not $PSBoundParameters.ContainsKey('ErrorAction'))
            {
                $ErrorActionPreference = $PSCmdlet.SessionState.PSVariable.GetValue("ErrorActionPreference")
            }

            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Write-Host', [System.Management.Automation.CommandTypes]::Cmdlet)
            $scriptCmd = {& $wrappedCmd @PSBoundParameters }
            $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch
        {
            throw
        }
    }

    process
    {
        if (ShouldPerformLogging -Path $_logFile)
        {
            foreach ($line in ($Object | Out-String -Stream))
            {
                if ($null -ne $Prepend)
                {
                    $results = $Prepend.Invoke($line)
                    if ($results.Count -gt 0 -and ($prependString = $results[0]) -is [System.String])
                    {
                        $line = "${prependString}${line}"
                    }
                }

                Add-Content -Path $_logFile -Value $line
            }
        }
            
        try
        {
            $steppablePipeline.Process($_)
        }
        catch
        {
            throw
        }
    }

    end
    {
        try
        {
            $steppablePipeline.End()
        }
        catch
        {
            throw
        }
    }
}

Set-Alias -Name Write-Host -Value Write-HostLog
Set-Alias -Name Write-Verbose -Value Write-VerboseLog
Set-Alias -Name Write-Debug -Value Write-DebugLog
Set-Alias -Name Write-Warning -Value Write-WarningLog
#Set-Alias -Name Write-Error -Value Write-ErrorLog

Export-ModuleMember -Function 'Write-DebugLog','Write-ErrorLog','Write-WarningLog','Write-VerboseLog','Write-HostLog'
Export-ModuleMember -Variable 'LogFilePreference'
Export-ModuleMember -Alias '*'
