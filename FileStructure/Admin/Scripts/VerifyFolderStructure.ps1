#requires -Version 2.0

[CmdletBinding()]
param (
    [ValidateNotNullOrEmpty()]
    [System.String]
    $CsvPath
)

Add-Type -TypeDefinition @'
    public enum DifferenceType {
        OnlyOnDisk,
        OnlyInCSV
    }
'@

function CheckFolder
{
    # Recursive function to walk a folder tree, comparing the contents to a tree of PSObjects which list the intended folder structure.
    # Similar to Compare-Object, outputs PSObjects for differences containing the path to the folder that's different, and a value indicating
    # whether the folder is only on disk or only in the CSV file.

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]
        $Directory,

        [Parameter(Mandatory = $true)]
        [psobject]
        $Node
    )

    $childFoldersChecked = @{}

    foreach ($dirInfo in $Directory.GetDirectories())
    {
        $childFoldersChecked[$dirInfo.Name] = $true

        $childNode = $Node.Children[$dirInfo.Name]
        if ($null -eq $childNode)
        {
            # This is a folder that exists on disk, but isn't in the CSV file
            New-Object psobject -Property @{
                Path = $dirInfo.FullName
                DifferenceType = [DifferenceType]::OnlyOnDisk
            }
        }
        else
        {
            # If this is not a "leaf node" of the CSV file (in other words, a subject folder), recursively check its contents
            if ($childNode.Children.Count -gt 0)
            {
                CheckFolder -Directory $dirInfo -Node $childNode
            }
        }
    }

    # Now look for any nodes from the CSV that didn't exist on disk.
    foreach ($childNode in $node.Children.Values)
    {
        if (-not $childFoldersChecked.ContainsKey($childNode.Name))
        {
            New-Object psobject -Property @{
                Path = Join-Path -Path $Directory.FullName -ChildPath $childNode.Name
                DifferenceType = [DifferenceType]::OnlyInCSV
            }
        }
    }
}

# TODO:  For now, this is the same code used in CreateFolderStructure.ps1, but we're talking about changing this to a config file that contains data
# about the root folder path.  Sticking with the original code for the moment, to get the rest of the script development done.

# Also, there's talk of combining this script with CreateFolderStructure.ps1.

$scriptFolder = Split-Path -Path $MyInvocation.MyCommand.Path -Parent

if (-not $PSBoundParameters.ContainsKey('CsvPath'))
{
    $CsvPath = Join-Path -Path $scriptFolder -ChildPath 'Structure.csv'
}

if (-not (Test-Path -Path $CsvPath))
{
    throw New-Object System.IO.FileNotFoundException($CsvPath)
}

# Sanity check to make sure this script is being executed in the intended folder structure.  The Data folder doesn't have to exist, but if the
# script's folder doesn't end with "Admin\Scripts", throw an error.

if ($scriptFolder -notmatch '(.+)\\Admin\\Scripts\\?$')
{
    throw "$($MyInvocation.ScriptName) script is not located in the expected folder structure (which must end in \admin\scripts\)."
}

$rootFolder = $matches[1]

# Make sure Data folder exists
$dataFolderPath = Join-Path -Path $rootFolder -ChildPath 'Data'

if (-not (Test-Path -Path $dataFolderPath -PathType Container))
{
    throw New-Object System.IO.DirectoryNotFoundException($dataFolderPath)
}

try
{
    $dataFolder = Get-Item -Path $dataFolderPath -ErrorAction Stop
}
catch
{
    throw
}

# The idea is to compare what's in the CSV file with what's actually on disk under the root folder.  If any folders specified in the CSV are missing, or
# if any extra folders exist on disk, let the user know (through various reporting methods spelled out in the requirements document).

# The trick is to make the script not complain about valid subfolders of subject folders in the CSV file.  Will figure out options for that as the code is written.

# First, convert the CSV file into a hierarchy of objects representing the folders

$rootNode = New-Object psobject -Property @{
    Name = Split-Path -Path $dataFolder -Leaf
    Children = @{}
}

Import-Csv -Path $CsvPath -ErrorAction Stop | 
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
            }

            $node.Children.Add($childName, $child)
        }

        $node = $child
    }
}

# Now recursively enumerate folders under $dataFolder, looking for differences.

CheckFolder -Directory $dataFolder -Node $rootNode
