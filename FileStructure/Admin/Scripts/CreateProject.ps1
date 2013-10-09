#requires -Version 2.0

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [System.String]
    $Path,

    [ValidateNotNullOrEmpty()]
    [System.String]
    $Name
)

$scriptFolder = Split-Path -Path $MyInvocation.MyCommand.Path -Parent

# Sanity check to make sure this script is being executed in the intended folder structure.  The Data folder doesn't have to exist, but if the
# script's folder doesn't end with "Admin\Scripts", throw an error.

if ($scriptFolder -notmatch '(.+)\\Admin\\Scripts\\?$')
{
    throw "$($MyInvocation.ScriptName) script is not located in the expected folder structure (which must end in \admin\scripts\)."
}

$rootFolder = $matches[1]

# Make sure the subject template folder exists
$projectTemplateFolder = Join-Path -Path $rootFolder -ChildPath 'Admin\ProjectTemplate'

if (-not (Test-Path -Path $projectTemplateFolder -PathType Container))
{
    throw New-Object System.IO.DirectoryNotFoundException($projectTemplateFolder)
}

# Validate user input
if (-not (Test-Path -Path $Path -PathType Container))
{
    throw New-Object System.IO.DirectoryNotFoundException($Path)
}

$invalidChars = [System.IO.Path]::GetInvalidFileNameChars()

if ($PSBoundParameters.ContainsKey('Name'))
{
    if ($Name.IndexOfAny($invalidChars) -ge 0)
    {
        throw "The Name argument ('$Name') contains invalid characters."
    }

    # TODO:  Requirements doc mentions project folder starts with today's year/month in YYYYMM format.
    # Is the user required to enter a folder with that name?  Should the script check for that and
    # throw an error / re-prompt the user if it's wrong?  Or should the script just prepend the date
    # information to whatever the user enters?
}
else
{
    while ($true)
    {
        $Name = Read-Host -Prompt 'Enter a new project name'

        if ($Name.IndexOfAny($invalidChars) -ge 0)
        {
            Write-Host "Value '$Name' contains invalid characters.  Please try again."
            continue
        }

        # TODO:  Again, should the script either verify or prepend the YYYYMM text at the start of the project name?

        break
    }
}

$projectPath = Join-Path -Path $Path -ChildPath $Name

try
{
    $null = New-Item -Path $projectPath -ItemType Directory -ErrorAction Stop

    $acl = Get-Acl -Path $projectPath -ErrorAction Stop

    $ace = New-Object System.Security.AccessControl.FileSystemAccessRule(
        'Everyone',
        'FullControl',
        'ContainerInherit, ObjectInherit',
        'None',
        'Allow'
    )

    $acl.AddAccessRule($ace)

    Set-Acl -Path $projectPath -AclObject $acl -ErrorAction Stop

    Copy-Item -Path (Join-Path -Path $projectTemplateFolder -ChildPath '*') -Destination $projectPath -Recurse -Force -Container -ErrorAction Stop
}
catch
{
    throw
}