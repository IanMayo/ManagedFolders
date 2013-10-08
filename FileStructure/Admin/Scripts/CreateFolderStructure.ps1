<#

Read CSV file (path can be specified via command-line parameter, or default to "Structure.csv" in the script's directory).

CSV file is assumed to have 5 fields named "Level1", "Level2", "Level3", "Level4", and "Level5".  Not all fields may be specified
in each record.  Each element must be combined to form a relative path.  The root of the path is ..\..\Data\ , starting from the script's
directory (which should be \Admin\scripts).

If the directory specified by that path already exists, do nothing?  Display an error?

If it doesn't exist, create it, then copy in the contents of ..\SubjectTemplate\ .

#>

#requires -Version 2.0

[CmdletBinding()]
param (
    [ValidateNotNullOrEmpty()]
    [System.String]
    $CsvPath
)

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

# Make sure the subject template folder exists
$subjectTemplateFolder = Join-Path -Path $rootFolder -ChildPath 'Admin\SubjectTemplate'

if (-not (Test-Path -Path $subjectTemplateFolder -PathType Container))
{
    throw New-Object System.IO.DirectoryNotFoundException($subjectTemplateFolder)
}

Write-Verbose 'Verifying Data folder and permissions...'

# Make sure Data folder exists, and has the proper permissions.
$dataFolder = Join-Path -Path $rootFolder -ChildPath 'Data'

if (-not (Test-Path -Path $dataFolder -PathType Container))
{
    try
    {
        $null = New-Item -Path $dataFolder -ItemType Directory -ErrorAction Stop
    }
    catch
    {
        throw
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
    throw
}

# Process the CSV file.
Write-Verbose "Reading data file '$CsvPath'...`r`n"

Import-Csv -Path $CsvPath -ErrorAction Stop |
ForEach-Object {
    $record = $_
    $subjectFolder = Join-Path -Path $rootFolder -ChildPath 'Data'

    for ($i = 1; $i -le 5; $i++)
    {
        if (-not [string]::IsNullOrEmpty($record."Level$i"))
        {
            $subjectFolder = Join-Path -Path $subjectFolder -ChildPath $record."Level$i"
        }
    }

    if (Test-Path -Path $subjectFolder -PathType Container)
    {
        Write-Verbose "Folder '$subjectFolder' already exists.  Skipping."
        return
    }

    Write-Verbose "Creating subject folder '$subjectFolder'..."

    try
    {
        # Create folder
        $null = New-Item -Path $subjectFolder -ItemType Directory -ErrorAction Stop

        # Set permissions
        $acl = Get-Acl -Path $subjectFolder -ErrorAction Stop

        $ace = New-Object System.Security.AccessControl.FileSystemAccessRule(
            'Everyone',
            'CreateDirectories',
            'None',
            'NoPropagateInherit',
            'Allow'
        )

        $acl.AddAccessRule($ace)

        Set-Acl -Path $subjectFolder -AclObject $acl -ErrorAction Stop

        # Copy template contents.
        Copy-Item -Path (Join-Path -Path $subjectTemplateFolder -ChildPath '*') -Destination $subjectFolder -Force -Recurse -Container -ErrorAction Stop
        
        # TODO:  Possibly generate "New Project Here.bat" file for this location, which will call the corresponding PowerShell script in the Admin\Scripts directory.

        Write-Verbose "Subject folder '$subjectFolder' created successfully.`r`n"
    }
    catch
    {
        Write-Error -ErrorRecord $_
        return
    }
}