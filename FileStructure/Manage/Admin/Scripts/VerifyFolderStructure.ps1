#requires -Version 2.0

[CmdletBinding()]
param (
    [ValidateNotNullOrEmpty()]
    [System.String]
    $ConfigFile = $null,

    [ValidateNotNullOrEmpty()]
    [System.String]
    $LogFile = $null,

    [switch]
    $SendMsg
)

#region Utility Functions

function Get-RelativePath
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Path,

        [Parameter(Mandatory = $true)]
        [System.String]
        $RelativeTo
    )

    $pattern = "^$([regex]::Escape($RelativeTo))\\?(.+)"

    if ($Path -notmatch $pattern)
    {
        throw "Path '$Path' is not a child of path '$RelativeTo'."
    }

    return $matches[1]
}

Add-Type -TypeDefinition @'
    public enum DifferenceType {
        OnlyOnDisk,
        OnlyInCSV
    }
'@

function ConfigureSubjectFolder
{
    [CmdletBinding()]
        param (
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]
        $Directory
    )

    try
    {
        # Copy template contents.
        Copy-Item -Path (Join-Path -Path $script:subjectTemplateFolder -ChildPath '*') -Destination $Directory.FullName -Force -Recurse -Container -ErrorAction Stop
        
        # Generate "New Project Here" shortcut.
        $shell = New-Object -ComObject WScript.Shell
        
        $shortcutPath = Join-Path -Path $Directory.FullName -ChildPath 'New Project Here.lnk'
        $shortcut = $shell.CreateShortcut($shortcutPath)

        $shortcut.TargetPath = "$PSHOME\powershell.exe"
        $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command & '$script:scriptFolder\CreateProject.ps1' -Path '$($Directory.FullName)'"
        $shortcut.Save()
        
        # Set permissions on initial contents of Subject folder.
        foreach ($item in $Directory.GetDirectories())
        {
            $acl = $item.GetAccessControl('Access')

            $ace = New-Object System.Security.AccessControl.FileSystemAccessRule(
                'Everyone',
                'FullControl',
                'ContainerInherit, ObjectInherit',
                'None',
                'Allow'
            )

            $acl.AddAccessRule($ace)

            $item.SetAccessControl($acl)
        }
    }
    catch
    {
        throw
    }
}

function CheckFolder
{
    # Recursive function to walk a folder tree, comparing the contents to a tree of PSObjects which list the intended folder structure.
    # Similar to Compare-Object, outputs PSObjects for differences containing the path to the folder that's different, and a value indicating
    # whether the folder is only on disk or only in the CSV file.

    # Also populates the ExistsOnDisk and Age fields of any nodes in the tree, for use in generating HTML code later.

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]
        $Directory,

        [Parameter(Mandatory = $true)]
        [psobject]
        $Node,

        [Parameter(Mandatory = $true)]
        [System.String]
        $RootPath,

        [Object]
        $IgnorePaths
    )

    $childFoldersChecked = @{}

    foreach ($dirInfo in $Directory.GetDirectories())
    {
        $relativePath = Get-RelativePath -Path $dirInfo.FullName -RelativeTo $RootPath
        if ($IgnorePaths -match "^\.?\\?$([regex]::Escape($relativePath))\\?$")
        {
            continue
        }

        $childFoldersChecked[$dirInfo.Name] = $true

        $childNode = $Node.Children[$dirInfo.Name]
        if ($null -eq $childNode)
        {
            # This is a folder that exists on disk, but isn't in the CSV file
            New-Object psobject -Property @{
                Path = $dirInfo.FullName
                DifferenceType = [DifferenceType]::OnlyOnDisk
            }

            continue
        }
        
        $childNode.ExistsOnDisk = $true
        $childNode.Age = (Get-Date) - $dirInfo.CreationTime

        # If this is not a "leaf node" of the CSV file (in other words, a subject folder), recursively check its contents
        if ($childNode.Children.Count -gt 0)
        {
            CheckFolder -Directory $dirInfo -Node $childNode -RootPath $RootPath -IgnorePaths $IgnorePaths
        }
        else
        {
            try
            {
                ConfigureSubjectFolder -Directory $dirInfo -ErrorAction Stop
            }
            catch
            {
                Write-Error -ErrorRecord $_
            }
        }

    } # end foreach ($dirInfo in $Directory.GetDirectories())

    # Now look for any nodes from the CSV that didn't exist on disk, and attempt to create them.
    foreach ($childNode in $node.Children.Values)
    {
        if ($childFoldersChecked.ContainsKey($childNode.Name))
        {
            continue
        }

        $path = Join-Path -Path $Directory.FullName -ChildPath $childNode.Name

        try
        {
            $dirInfo = New-Item -Path $path -ItemType Directory -ErrorAction Stop
        }
        catch
        {
            Write-Error -ErrorRecord $_                
                
            New-Object psobject -Property @{
                Path = Join-Path -Path $Directory.FullName -ChildPath $childNode.Name
                DifferenceType = [DifferenceType]::OnlyInCSV
            }

            continue
        }
        
        $childNode.ExistsOnDisk = $true
        $childNode.Age = (Get-Date) - $dirInfo.CreationTime

        if ($childNode.Children.Count -gt 0)
        {
            CheckFolder -Directory $dirInfo -Node $childNode -RootPath $RootPath -IgnorePaths $IgnorePaths
        }
        else
        {
            try
            {
                ConfigureSubjectFolder -Directory $dirInfo -ErrorAction Stop
            }
            catch
            {
                Write-Error -ErrorRecord $_
            }
        }

    } # end foreach ($childNode in $node.Children.Values)

}# end function CheckFolder

function Get-HtmlIndexCode
{
    # Walks the tree of objects starting with $Node, generating HTML code to be injected into the
    # appropriate places of a template file.

    # If the -TopLevelShortcuts switch is specified, the function is not recursive, and has
    # slightly different links (pointing to anchors on this page, rather than a file:/// URI
    # to the folders themselves.)

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Path,

        [Parameter(Mandatory = $true)]
        [psobject]
        $Node,

        [System.Text.StringBuilder]
        $Html = $null,

        [System.UInt32]
        $Indent = 0,

        [switch]
        $TopLevelShortcuts,

        [switch]
        $RecursiveCall
    )

    if ($null -ne $Html)
    {
        $stringBuilder = $Html
    }
    else
    {
        $stringBuilder = New-Object System.Text.StringBuilder
    }

    if ($TopLevelShortcuts)
    {
        foreach ($childNode in $Node.Children.Values)
        {
            if ($childNode.ExistsOnDisk)
            {
                $code = "<li><a href=""#$($childNode.Name)"">$($childNode.Name)</a></li>"
                $null = $stringBuilder.AppendLine(("{0,$Indent}{1}" -f ' ', $code))
            }
        }

        Write-Output $stringBuilder.ToString()
    }

    else
    {
        foreach ($childNode in $Node.Children.Values)
        {
            if ($childNode.ExistsOnDisk)
            {
                $listTag = '<li'
                
                if ($childNode.Age -is [System.Timespan] -and $childNode.Age.TotalDays -gt 30)
                {
                    $listTag += ' class="recent"'
                }

                if (-not $RecursiveCall)
                {
                    $listTag += " id=""$($childNode.Name)"""
                }

                $listTag += '>'

                $childPath = Join-Path -Path $Path -ChildPath $childNode.Name

                $uri = New-Object System.Uri($childPath)
                $code = "$listTag<span>$($childNode.Name)</span><a href=""$($uri.AbsoluteUri)""></a>"
                $null = $stringBuilder.Append(("{0,$Indent}{1}" -f ' ', $code))

                if ($childNode.Children.Count -gt 0)
                {
                    $null = $stringBuilder.AppendLine()

                    $Indent += 2
                    $null = $stringBuilder.AppendLine(("{0,$Indent}<ul>" -f ' '))

                    $null = Get-HtmlIndexCode -Path $childPath -Node $childNode -Html $stringBuilder -Indent ($Indent + 2) -RecursiveCall

                    $null = $stringBuilder.AppendLine("{0,$Indent}</ul>" -f ' ')
                    $Indent -= 2

                    $null = $stringBuilder.Append(("{0,$Indent}" -f ' '))
                }

                $null = $stringBuilder.AppendLine('</li>')
            }
        }

        if ($null -eq $Html)
        {
            Write-Output $stringBuilder.ToString()
        }
    }
}

function Import-IniFile
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Path
    )

    New-Variable -Name UnnamedSection -Value '\\Unnamed//' -Option ReadOnly

    if (-not (Test-Path -Path $Path -PathType Leaf))
    {
        throw New-Object System.IO.FileNotFoundException($Path)
    }

    $iniFile = @{}

    $currentSection = $null

    Get-Content -Path $Path -ErrorAction Stop |
    ForEach-Object {
        $line = $_

        switch -Regex ($line)
        {
            # Comments
            '^\s*;'
            { }

            # Sections
            '^\s*\[(.+?)\]\s*$'
            {
                $sectionName = $matches[1]

                if ($iniFile.ContainsKey($sectionName))
                {
                    $currentSection = $iniFile[$sectionName]
                }
                else
                {
                    $currentSection = @{}
                    $iniFile.Add($sectionName, $currentSection)
                }
            }

            # Key = Value pairs
            '^\s*(.+?)\s*=\s*(.+?)\s*$'
            {
                $name = $matches[1]
                $value = $matches[2]

                if ($null -eq $currentSection)
                {
                    $currentSection = @{}
                    $iniFile.Add($UnnamedSection, $currentSection)
                }

                $currentSection[$name] = $value
            }
        }
    }

    Write-Output $iniFile
}

#endregion

#region Main script

$scriptFolder = Split-Path -Path $MyInvocation.MyCommand.Path -Parent

# Set up log file, if requested.

if ($null -ne $LogFile)
{
    try
    {
        Import-Module $scriptFolder\Logging -ErrorAction Stop
        $LogFilePreference = $LogFile
    }
    catch
    {
        throw
    }
}

# Read configuration file

if (-not $PSBoundParameters.ContainsKey('ConfigFile'))
{
    $ConfigFile = Join-Path -Path $scriptFolder -ChildPath 'config.ini'
}

if (-not (Test-Path -Path $ConfigFile))
{
    Write-ErrorLog -Exception (New-Object System.IO.FileNotFoundException($ConfigFile))
    exit 1
}

Write-Verbose "Reading configuration file '$ConfigFile'..."

try
{
    $config = Import-IniFile -Path $ConfigFile -ErrorAction Stop
    
    if (-not $config.ContainsKey('Configuration'))
    {
        throw "Configuration file '$ConfigFile' does not contain the required [Configuration] section."
    }
}
catch
{
    Write-ErrorLog -ErrorRecord $_
    exit 1
}

# Fetch options from ini file, and make sure all required templates / directories exist.

$paths = @(
    New-Object psobject -Property @{
        VariableName = 'csvPath'
        Section = 'Configuration'
        Value = 'CsvFile'
        Default = Join-Path -Path $scriptFolder -ChildPath 'Structure.csv'
        Type = 'Leaf'
        Required = $true
    }

    New-Object psobject -Property @{
        VariableName = 'subjectTemplateFolder'
        Section = 'Configuration'
        Value = 'SubjectTemplate'
        Default = Join-Path -Path $scriptFolder -ChildPath '..\SubjectTemplate'
        Type = 'Container'
        Required = $false
    }

    New-Object psobject -Property @{
        VariableName = 'dataFolder'
        Section = 'Configuration'
        Value = 'DataFolder'
        Default = Join-Path -Path $scriptFolder -ChildPath '..\..\..'
        Type = 'Container'
        Required = $false
    }

    New-Object psobject -Property @{
        VariableName = 'indexTemplate'
        Section = 'Configuration'
        Value = 'HtmlTemplate'
        Default = Join-Path -Path $scriptFolder -ChildPath '..\index_template.html'
        Type = 'Leaf'
        Required = $true
    }

    New-Object psobject -Property @{
        VariableName = 'indexFile'
        Section = 'Configuration'
        Value = 'HtmlOutput'
        Default = Join-Path -Path $scriptFolder -ChildPath '..\index.html'
        Type = 'Leaf'
        Required = $false
    }
)

foreach ($var in $paths)
{
    $temp = $config[$var.Section][$var.Value]
    if ([string]::IsNullOrEmpty($temp))
    {
        $temp = $var.Default
    }
    elseif (-not [System.IO.Path]::IsPathRooted($temp))
    {
        $temp = Join-Path -Path $scriptFolder -ChildPath $temp
    }

    $temp = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($temp)

    Set-Variable -Name $var.VariableName -Value $temp -Force -Scope Script -Option ReadOnly

    if ($var.Required -and -not (Test-Path -Path $temp -PathType $var.Type))
    {
        if ($var.Type -eq 'Container')
        {
            $exception = New-Object System.IO.DirectoryNotFoundException($temp)
        }
        else
        {
            $exception = New-Object System.IO.FileNotFoundException($temp)
        }

        Write-ErrorLog -Exception $exception
        exit 1
    }
}

Write-Verbose 'Verifying Data folder and permissions...'

# Make sure Data folder exists, and has the proper permissions.
if (-not (Test-Path -Path $dataFolder -PathType Container))
{
    try
    {
        $null = New-Item -Path $dataFolder -ItemType Directory -ErrorAction Stop
    }
    catch
    {
        Write-ErrorLog -ErrorRecord $_
        exit 1
    }
}

try
{
    $acl = Get-Acl -Path $dataFolder -ErrorAction Stop
    $dirty = $false

    $entries = @(
        New-Object System.Security.AccessControl.FileSystemAccessRule(
            'Everyone',
            'ReadAndExecute',
            'ContainerInherit, ObjectInherit',
            'None',
            'Allow'
        )

        New-Object System.Security.AccessControl.FileSystemAccessRule(
            'BUILTIN\Administrators',
            'FullControl',
            'ContainerInherit, ObjectInherit',
            'None',
            'Allow'
        )
    )

    foreach ($entry in $entries)
    {
        $matchingAce = $acl.Access |
                       Where-Object {
                           $_.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]) -eq $entry.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]) -and
                           $_.FileSystemRights -eq $entry.FileSystemRights -and
                           $_.AccessControlType -eq $entry.AccessControlType -and
                           $_.IsInherited -eq $false -and
                           $_.InheritanceFlags -eq $entry.InheritanceFlags -and
                           $_.PropagationFlags -eq $entry.PropagationFlags
                       }

        if ($null -eq $matchingAce)
        {
            $acl.AddAccessRule($entry)
            $dirty = $true
        }
    }

    if ($dirty)
    {
        Set-Acl -Path $dataFolder -AclObject $acl -ErrorAction Stop
    }
}
catch
{
    Write-ErrorLog -ErrorRecord $_
    exit 1
}

# The idea is to compare what's in the CSV file with what's actually on disk under the root folder.  If any folders specified in the CSV are missing, or
# if any extra folders exist on disk, let the user know (through various reporting methods spelled out in the requirements document).

# First, convert the CSV file into a hierarchy of objects representing the folders

$rootNode = New-Object psobject -Property @{
    Name = Split-Path -Path $dataFolder -Leaf
    Children = @{}
    ExistsOnDisk = $true
    Age = $null
}

Import-Csv -Path $csvPath -ErrorAction Stop | 
ForEach-Object {
    $record = $_

    $node = $rootNode

    for ($i = 1; $i -le 5; $i++)
    {
        $childName = $record."Level$i"

        if ([string]::IsNullOrEmpty($childName))
        {
            continue
        }

        $child = $node.Children[$childName]

        if ($null -eq $child)
        {
            $child = New-Object psobject -Property @{
                Name = $childName
                Children = @{}
                ExistsOnDisk = $false
                Age = $null
            }

            $node.Children.Add($childName, $child)
        }

        $node = $child
    }
}

# Now recursively enumerate folders under $dataFolder, looking for differences.
$ignorePaths = $config['IgnorePaths']
if ($ignorePaths -is [hashtable])
{
    $ignorePaths = $ignorePaths.Values
}

$differences = @(CheckFolder -Directory $dataFolder -Node $rootNode -RootPath $dataFolder -IgnorePaths $ignorePaths -ErrorAction SilentlyContinue -ErrorVariable Err)

# Finish generating HTML and save index file.
$topLevelList = Get-HtmlIndexCode -Path $dataFolder -Node $rootNode -TopLevelShortcuts -Indent 8
$index = Get-HtmlIndexCode -Path $dataFolder -Node $rootNode -Indent 4

$ignore = $false

Get-Content -Path $indexTemplate -ErrorAction SilentlyContinue -ErrorVariable +Err |
ForEach-Object {
    $line = $_

    switch -Regex ($line) {
        '^\s*<!--\s*INDEX_START\s*-->\s*$'
        {
            $ignore = $true
            Write-Output $topLevelList
            break
        }
        
        '^\s*<!--\s*INDEX_END\s*-->\s*$'
        {
            $ignore = $false
            break
        }

        '^\s*<!--\s*LISTING_START\s*-->\s*$'
        {
            $ignore = $true
            Write-Output $index
            break
        }

        '^\s*<!--\s*LISTING_END\s*-->\s*$'
        {
            $ignore = $false
            break
        }

        '^\s*<!--\s*TIMESTAMP_START\s*-->\s*$'
        {
            $ignore = $true
            Write-Output "    <div id=""timestamp"">$(Get-Date -Format 'yyyyMMdd HH:mm')</div>"
            break
        }

        '^\s*<!--\s*TIMESTAMP_END\s*-->\s*$'
        {
            $ignore = $false
            break
        }

        '.*'
        {
            if (-not $ignore)
            {
                Write-Output $line
            }
        }
    }
} |
Set-Content -Path $indexFile -Force -ErrorAction SilentlyContinue -ErrorVariable +Err


$message = New-Object System.Text.StringBuilder

if ($differences.Count -gt 0)
{
    $null = $message.AppendLine("Differences between the disk's folder structure and the '$csvPath' file were detected:")

    $onlyOnDisk = @(
        $differences |
        Where-Object { $_.DifferenceType -eq [DifferenceType]::OnlyOnDisk } |
        Select-Object -ExpandProperty Path
    )

    $onlyInCsv = @(
        $differences |
        Where-Object { $_.DifferenceType -eq [DifferenceType]::OnlyInCSV } |
        Select-Object -ExpandProperty Path
    )

    if ($onlyOnDisk.Count -gt 0)
    {
        $null = $message.AppendLine("`r`nFolders that exist on disk, but not in the CSV file:`r`n`r`n$($onlyOnDisk | Out-String)")
    }

    if ($onlyInCsv.Count -gt 0)
    {
        $null = $message.AppendLine("`r`nFolders that are defined in the CSV file, but were not found on disk:`r`n`r`n$($onlyInCsv | Out-String)")
    }
}

if ($Err.Count -gt 0)
{
    $null = $message.AppendLine(
        "`r`nThe following errors were encountered:`r`n$($Err | Out-String)"
    )
}

if ($message.Length -gt 0)
{
    Write-Warning $message.ToString()
    
    if ($SendMsg)
    {
        $null = msg.exe * $message
    }
}

#endregion