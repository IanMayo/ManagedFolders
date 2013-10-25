<#
.Synopsis
   Verifies or creates a folder structure on disk, based on the contents of a CSV file and template directories.
.DESCRIPTION
   Each record of the CSV file defines a "Subject" folder.  The script attempts to create each of these subject
   directories (if they do not already exist), and copies the contents of a "Subject Folder" template into each.
   It also sets NTFS permissions of the directory structure according to hard-coded client requirements, and
   creates a "New Project Here" shortcut in each subject directory; this shortcut refers to a separate script.

   In addition to verifying the folder structure, this script creates an index HTML file with links to directories
   in this folder tree, based on a template HTML file.

   If any folders exist that are not part of the CSV-defined structure, or if any required folder could not be
   created, the script will output error messages according to which parameters are supplied on the command-line.
   It will always write Warning output to the console, but can also create a log file on disk and/or send a message
   using msg.exe.
.PARAMETER ConfigFile
   Optional path to an INI file.  If not specified, the value defaults to "config.ini" in the same directory as this script.

   The INI file must contain a [Configuration] section with the following values:  (If relative paths are specified for any
   values in this section, they are relative to the folder containing this script)

   CsvFile:  Path to the CSV file which defines the folder structure.
   DataFolder:  Path to the root folder where the folder structure is to be checked / created.
   HtmlTemplate:  Path to the template HTML file used for generating the index file.
   HtmlOutput:  Path where the index file should be saved.
   SubjectTemplate:  Path to the Subject Template folder which is copied in to each subject directory.

   Also, the configuration file may optionally include a section named [IgnorePaths].  Each value in this section is
   ignored when walking the folder tree; the script will not produce any errors or warnings due to the existence of these
   folders, or any folders contained within them.  The paths are all relative to the DataFolder specified in the [Configuration]
   section.
   
   The value names of this section are irrelevant; the script enumerates them all and only cares about the values.  For example:

   [IgnorePaths]
   Tom = SomeFolder
   Dick = SomeOtherFolder\ChildFolder
   Harry = PayNoAttentionTo\TheManBehind\TheCurtain\
.PARAMETER LogFile
   Optional path to a log file that the script should produce.  If this parameter is not specified, no log file is created.
   If it is specified, all console output will be copied to the log file (including prepended date-and-time information on
   each line.)
.PARAMETER SendMsg
   Optional switch parameter.  If set, the script will also send out any errors or warnings via msg.exe.
.EXAMPLE
   .\VerifyFolderStructure.ps1

   Uses the config.ini file in the same directory as the script, and does not produce a log file or call msg.exe.
.EXAMPLE
   .\VerifyFolderStructure.ps1 -ConfigFile .\SomeFile.ini -LogFile .\VerifyFolderStructure.log -SendMsg

   Uses the SomeFile.ini config file, creates a log file named VerifyFolderStructure.log, and calls msg.exe if any problems are encountered.
.INPUTS
   None.  This script does not accept pipeline input.
.OUTPUTS
   None.  This script writes all output directly to the host, and does not produce objects on the pipeline.
#>

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
    # Returns the "child" portion of a path, relative to the specified root folder.  For example:
    #
    # Get-RelativePath -Path C:\Folder\Subfolder\File.txt -RelativeTo C:\Folder
    # would return:  Subfolder\File.txt

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
    # For the specified Subject Folder (specified by a DirectoryInfo object), copies the contents of the subject template
    # folder to this location.  Grants Full Control permission to any child directories of the subject folder, and creates
    # a "New Project Here" shortcut which launches the CreateProject script with a Path parameter specifying this subject
    # folder, and the same ConfigFile parameter used in the call to VerifyFolderStructure.ps1.

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

        # Original PowerShell version.  Replaced with VBScript for performance tests.
        # Note:  VBScript performance was preferred, so this code can be deleted, if you never want
        # to worry about going back to a PowerShell version of CreateProject.

        # $shortcut.TargetPath = "$PSHOME\powershell.exe"
        # $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command & '$script:scriptFolder\CreateProject.ps1' -Path '$($Directory.FullName)' -ConfigFile '$script:ConfigFile'"

        # VBScript version
        $shortcut.TargetPath = "wscript.exe"
        $shortcut.Arguments = """$script:scriptFolder\CreateProject.vbs"" /Path:""$($Directory.FullName)"" /ConfigFile:""$script:ConfigFile"""

        $shortcut.Save()

        # Copy shortcut file to Projects directory.
        $projectsPath = Join-Path -Path $Directory.FullName -ChildPath 'Projects'
        if (Test-Path -Path $projectsPath -PathType Container)
        {
            Copy-Item -Path $shortcutPath -Destination $projectsPath -Force -ErrorAction Stop
        }
        
        # Set permissions on contents of Subject folder.
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

function EnumerateSubjectFolderContents
{
    # Recursive function to walk the contents of a subject folder, adding information about them to a tree structure.
    # Each node o the tree contains 4 properties:

    # Name:  Name of the directory (not including parent path)
    # Children:  Hashtable of child nodes, keyed by their Name properties.
    # ExistsOnDisk:  Boolean value which is always set to $true for nodes created by this function.
    # Age:  A Timespan value indicating the difference between now and the CreationTime of the directory.

    # This tree structure is later used in the creation of the index html file.

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]
        $Directory,

        [Parameter(Mandatory = $true)]
        [psobject]
        $Node,

        [System.Int32]
        $Levels = -1
    )

    if ($Levels -eq 0)
    {
        return
    }

    foreach ($dirInfo in $Directory.GetDirectories())
    {
        $child = New-Object psobject -Property @{
            Name = $dirInfo.Name
            Children = @{}
            ExistsOnDisk = $true
            Age = (Get-Date) - $dirInfo.CreationTime
        }

        $Node.Children[$dirInfo.Name] = $child

        EnumerateSubjectFolderContents -Directory $dirInfo -Node $child -Levels ($Levels - 1)
    }
}

function CheckFolder
{
    # Recursive function to walk a folder tree, comparing the contents to a tree of PSObjects which list the intended folder structure.
    # Similar to Compare-Object, outputs PSObjects for differences containing the path to the folder that's different, and a value indicating
    # whether the folder is only on disk or only in the CSV file.

    # Also populates the ExistsOnDisk and Age fields of any nodes in the tree, for use in generating HTML code later.

    # For each subject folder (leaf node of the initial structure, as defined by the CSV file), call EnumerateSubjectFolderContents and
    # ConfigureSubjectFolder; see comments on those functions in their definitions.

    # Note:  These functions change the tree structure passed in to the Node parameter of the initial call to CheckFolder.  There should
    # only ever be one call to CheckFolder for each node in the tree; if it is called a second time, the logic will fail due to the presence
    # of child nodes under each "Subject" folder.

    # This function produces pipeline output in the form of objects indicating differences between what's defined in the CSV file, and
    # what is on disk (not counting folders that are successfully created, if they didn't already exist.)  For example:

    # Path:  C:\ManagedFolders\Data\SomeFolder
    # DifferenceType:  [DifferenceType]::OnlyOnDisk

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

    foreach ($dirInfo in $Directory.GetDirectories())
    {
        # Check whether this directory should be ignored.

        $relativePath = Get-RelativePath -Path $dirInfo.FullName -RelativeTo $RootPath
        if ($IgnorePaths -match "^\.?\\?$([regex]::Escape($relativePath))\\?$")
        {
            continue
        }

        $childNode = $Node.Children[$dirInfo.Name]
        if ($null -eq $childNode)
        {
            # This is a folder that exists on disk, but isn't in the CSV file.  Output an object indicating this information.
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
            # This is a subject folder.  Copy the template contents into it, configure permissions, create the shortcut, and enumerate
            # its contents to two levels deep (for the index html file).

            ConfigureSubjectFolder -Directory $dirInfo
            EnumerateSubjectFolderContents -Directory $dirInfo -Node $childNode -Levels 2
        }

    } # end foreach ($dirInfo in $Directory.GetDirectories())

    # Now look for any nodes from the CSV that didn't exist on disk, and attempt to create them.
    foreach ($childNode in $node.Children.Values)
    {
        if ($childNode.ExistsOnDisk)
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

        [System.UInt32]
        $Level = 1
    )

    # If the caller passed in a StringBuilder object, append to that.  Otherwise, create a new StringBuilder local to this
    # function.  (Note:  When making recursive calls, pass the StringBuilder in with the -Html parameter, regardless of whether
    # it was passed in originally by the caller or not.)

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
                # Generate the full HTML tags for this item.  Each LI tag should have a class attribute with a value of
                # l<level number> <first word of folder name in lower case>, such as <li class="l1 foldername">.

                # Leaf nodes should be output on one line, ie:
                #
                # <li class="l1 foldername"><span>FolderName</span><a href="Folder URI"></a></li>
                #
                # (The empty text between the <a> and </a> tags is intentional; the CSS / JavaScript code
                # somehow turns that into a graphic link.)
                
                # For nodes that contain child folders, the <li> and </li> tags should be on separate lines, with
                # another unordered list contained between them, ie:
                #
                # <li class="l1 foldername"><span>FolderName</span><a href="Folder URI"></a>
                #   <ul>
                #     <li class="l2 childfoldername"><span>ChildFolderName</span><a href="Folder URI"></a></li>
                #   </ul>
                # </li>                $listTag = '<li'
                
                # Folders that have existed for less than 30 days are flagged as "recent"; the CSS code of the html file
                # causes them to display differently.

                $listTag = '<li'

                $classes = @(
                    "l$Level",
                    $childNode.Name.Split(' ')[0].ToLower()
                )

                if ($childNode.Age -is [System.Timespan] -and $childNode.Age.TotalDays -lt 30)
                {
                    $classes += 'recent'
                }

                # If we are dealing with the top level, the client's requirements state that we should add an
                # id="FolderName" attribute to the LI tag.

                if ($Level -eq 1)
                {
                    $listTag += " id=""$($childNode.Name)"""
                }

                $listTag += " class=""$($classes -join ' ')"">"

                $childPath = Join-Path -Path $Path -ChildPath $childNode.Name

                $uri = New-Object System.Uri($childPath)
                $code = "$listTag<span>$($childNode.Name)</span><a href=""$($uri.AbsoluteUri)""></a>"
                $null = $stringBuilder.Append(("{0,$Indent}{1}" -f ' ', $code))

                if ($childNode.Children.Count -gt 0)
                {
                    $null = $stringBuilder.AppendLine()

                    $Indent += 2
                    $null = $stringBuilder.AppendLine(("{0,$Indent}<ul>" -f ' '))

                    $null = Get-HtmlIndexCode -Path $childPath -Node $childNode -Html $stringBuilder -Indent ($Indent + 2) -Level ($Level + 1)

                    $null = $stringBuilder.AppendLine("{0,$Indent}</ul>" -f ' ')
                    $Indent -= 2

                    $null = $stringBuilder.Append(("{0,$Indent}" -f ' '))
                }

                $null = $stringBuilder.AppendLine('</li>')
            }
        }

        # For performance reasons, this function only writes output to the pipeline at the root level.  Recursive calls
        # just append directly to the same StringBuilder object, instead of making the caller do it (involving extra string
        # copies in memory.)

        if ($null -eq $Html)
        {
            Write-Output $stringBuilder.ToString()
        }
    }
}

function Import-IniFile
{
    # Reads an INI file from disk, importing it into a hashtable of hashtables.  Each key at the "root" level
    # is a section name, which contains another hashtable of the key=value pairs from that section of the file.

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
elseif (-not [System.IO.Path]::IsPathRooted($ConfigFile))
{
    $ConfigFile = Join-Path -Path $scriptFolder -ChildPath $ConfigFile
}

$ConfigFile = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($ConfigFile)

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