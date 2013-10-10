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

    # Also appends HTML code for the index file into the StringBuilder specified by the $Html parameter, if present.

    # TODO:  The HTML format in the sample index.html and in the pdf requirements document are different.  I drafted this code based on the sample file,
    # but it will probably need an update shortly to match the more detailed requirements.  It just means changing the current unordered lists at each
    # level of "subject" folder into <H1>, <H2>, etc, and adding an extra function to enumerate the contents of subject folders (two levels deep) to add
    # unordered lists.

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]
        $Directory,

        [Parameter(Mandatory = $true)]
        [psobject]
        $Node,

        [System.Text.StringBuilder]
        $Html = $null,

        [System.UInt32]
        $Indent = 0
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
            if ($null -ne $Html)
            {
                $uri = New-Object System.Uri($dirInfo.FullName)
                $code = "<li><a href=""$($uri.AbsoluteUri)"">$($dirInfo.Name)</a>"
                $null = $Html.Append(("{0,$Indent}{1}" -f ' ', $code))
            }

            # If this is not a "leaf node" of the CSV file (in other words, a subject folder), recursively check its contents
            if ($childNode.Children.Count -gt 0)
            {
                if ($null -ne $Html)
                {
                    $null = $Html.AppendLine()

                    $Indent += 2
                    $null = $Html.AppendLine(("{0,$Indent}<ul>" -f ' '))
                }
                
                CheckFolder -Directory $dirInfo -Node $childNode -Html $Html -Indent ($Indent + 2)
                
                if ($null -ne $Html)
                {
                    $null = $Html.AppendLine("{0,$Indent}</ul>" -f ' ')
                    $Indent -= 2

                    $null = $Html.Append(("{0,$Indent}" -f ' '))
                }
            }

            if ($null -ne $Html)
            {
                $null = $Html.AppendLine("</li>")
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

# Begin forming HTML index code.

foreach ($fileName in ('header.inc', 'footer.inc'))
{
    $filePath = Join-Path -Path $scriptFolder -ChildPath $fileName
    if (-not (Test-Path -Path $FilePath))
    {
        throw New-Object System.IO.FileNotFoundException($filePath)
    }
}

$html = New-Object System.Text.StringBuilder

try
{
    $null = $html.AppendLine([System.IO.File]::ReadAllText((Join-Path -Path $scriptFolder -ChildPath 'header.inc')))
}
catch
{
    throw
}

$null = $html.AppendLine('<h1>Directory listing</h1>').AppendLine()
$null = $html.AppendLine('<ul id="dirListing">')

# Now recursively enumerate folders under $dataFolder, looking for differences.  This function also generates HTML code as it goes, appending it to the
# $html StringBuilder.

$differences = CheckFolder -Directory $dataFolder -Node $rootNode -Html $html

# TODO:  Write various options for reporting the differences.  For now, just output to the screen.

if ($differences.Count -gt 0)
{
    Write-Warning "Differences between the disk's folder structure and the '$CsvPath' file were detected:"

    $onlyOnDisk = $differences |
    Where-Object { $_.DifferenceType -eq [DifferenceType]::OnlyOnDisk } |
    Select-Object -ExpandProperty Path

    $onlyInCsv = $differences |
    Where-Object { $_.DifferenceType -eq [DifferenceType]::OnlyInCSV } |
    Select-Object -ExpandProperty Path

    if ($onlyOnDisk.Count -gt 0)
    {
        Write-Warning "Folders that exist on disk, but not in the CSV file:`r`n$($onlyOnDisk | Out-String)"
    }

    if ($onlyInCsv.Count -gt 0)
    {
        Write-Warning "Folders that are defined in the CSV file, but were not found on disk:`r`n$($onlyInCsv | Out-String)"
    }
}

# Finish generating HTML and save index file.

$null = $html.AppendLine('</ul>')

try
{
    $null = $html.AppendLine([System.IO.File]::ReadAllText((Join-Path -Path $scriptFolder -ChildPath 'footer.inc')))
}
catch
{
    throw
}

Set-Content -Path (Join-Path -Path $rootFolder -ChildPath 'Admin\index.html') -Value $html.ToString() -Force
