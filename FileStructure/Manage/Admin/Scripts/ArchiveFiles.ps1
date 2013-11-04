<#
.Synopsis
   Archives folders based on age and/or file type, and creates symbolic
.DESCRIPTION
   Searches a directory tree for any files that are at least 3 years old (other than file extensions configurable by the user),
   or any file larger than 128MB which is at least 1 year old.  Any such files are moved to a user-configurable Archive directory,
   and an NTFS Symbolic Link is created in the original file location.
.PARAMETER ConfigFile
   Optional path to an INI file.  If not specified, the value defaults to "config.ini" in the same directory as this script.

   The INI file must contain a [Configuration] section with the following values:  (If relative paths are specified for any
   values in this section, they are relative to the folder containing this script)

   CsvFile:  Path to the CSV file which defines the folder structure.
   DataFolder:  Path to the root folder where the folder structure is to be checked / created.
   NO_ARCHIVE_TYPES:  Optional comma-separated list of file extensions which should be ignored, even if they are older than 3 years.
   ArchiveFolder:  Path to the folder where archived files should be moved.  The relative paths in DataFolder and ArchiveFolder will be identical.   
.PARAMETER LogFile
   Optional path to a log file that the script should produce.  If this parameter is not specified, no log file is created.
   If it is specified, all console output will be copied to the log file (including prepended date-and-time information on
   each line.)
.PARAMETER SendMsg
   Optional switch parameter.  If set, the script will also send out any errors or warnings via msg.exe.
.EXAMPLE
   .\ArchiveFiles.ps1

   Uses the config.ini file in the same directory as the script, and does not produce a log file or call msg.exe.
.EXAMPLE
   .\ArchiveFiles.ps1 -ConfigFile .\SomeFile.ini -LogFile .\VerifyFolderStructure.log -SendMsg

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
    $LogFile = $null
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

# Check for required privilege

if (-not (whoami /priv | Select-String -Pattern 'SeCreateSymbolicLinkPrivilege' -SimpleMatch))
{
    Write-ErrorLog 'Current process does not have the "Create Symbolic Links" privilege.  By default, this script must be executed with the "Run As Administrator" option to enable this privilege.'
    exit 1
}

# Process working directory changed to something local, to avoid potential complaints about UNC paths later from cmd.exe.
[System.Environment]::CurrentDirectory = $env:windir

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
        VariableName = 'dataFolder'
        Section = 'Configuration'
        Value = 'DataFolder'
        Default = Join-Path -Path $scriptFolder -ChildPath '..\..\..'
        Type = 'Container'
        Required = $false
    }

    New-Object psobject -Property @{
        VariableName = 'archiveFolder'
        Section = 'Configuration'
        Value = 'ArchiveFolder'
        Default = Join-Path -Path $scriptFolder -ChildPath '..\..\..\..\Archive'
        Type = 'Leaf'
        Required = $false
    }

    New-Object psobject -Property @{
        VariableName = 'subjectTemplateFolder'
        Section = 'Configuration'
        Value = 'SubjectTemplate'
        Default = Join-Path -Path $scriptFolder -ChildPath '..\SubjectTemplate'
        Type = 'Container'
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

$templateFiles = New-Object System.Collections.ArrayList
if (Test-Path -Path $subjectTemplateFolder -PathType Container)
{
    Get-ChildItem -Path $subjectTemplateFolder -Recurse -Force |
    Where-Object { -not $_.PSIsContainer } |
    ForEach-Object {
        $null = $templateFiles.Add((Get-RelativePath -Path $_.FullName -RelativeTo $subjectTemplateFolder))
    }
}

$noArchiveTypes = $config['Configuration']['NO_ARCHIVE_TYPES']

if ([string]::IsNullOrEmpty($noArchiveTypes))
{
    $noArchiveTypes = @()
}
else
{
    $noArchiveTypes = (($noArchiveTypes -split ',') -replace '^\s+|\s+$') -replace '^(?=[^\.])', '.'
}

# Make sure Data folder exists.
if (-not (Test-Path -Path $dataFolder -PathType Container))
{
    Write-ErrorLog -Exception (New-Object System.IO.DirectoryNotFoundException($dataFolder))
    exit 1
}

# Make sure Archive folder exists.
if (-not (Test-Path -Path $archiveFolder -PathType Container))
{
    try
    {
        $null = New-Item -Path $archiveFolder -ItemType Directory -ErrorAction Stop
    }
    catch
    {
        Write-ErrorLog -ErrorRecord $_
        exit 1
    }
}

Import-Csv -Path $csvPath -ErrorAction Stop | 
ForEach-Object {
    $record = $_
    $subjectPath = $dataFolder

    for ($i = 1; $i -le 5; $i++)
    {
        $childName = $record."Level$i"

        if (-not [string]::IsNullOrEmpty($childName))
        {
            $subjectPath = Join-Path -Path $subjectPath -ChildPath $childName
        }
    }

    if (-not (Test-Path -Path $subjectPath -PathType Container))
    {
        return
    }

    Get-ChildItem -Path $subjectPath -Recurse -Force |
    ForEach-Object {
        $item = $_

        # Ignore directories and existing symbolic links ("Reparse Points")

        if ($item.PSIsContainer) { return }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return }

        # Ignore files that came from the subject template folder.
        $relativePath = Get-RelativePath -Path $item.FullName -RelativeTo $subjectPath

        if ($templateFiles -contains $relativePath) { return }

        # Only archive files that are at least 1 year old and 128MB in size, OR
        # at least 3 years old with a file extension not on the "No Archive" list.
        #
        # Note:  Files >= 128MB will be archived regardless of their extension.

        $fileAge = (Get-Date) - $item.LastWriteTime
        
        if (($fileAge.Days -ge 365 -and $item.Length -ge 128MB) -or
            ($fileAge.Days -ge (365*3) -and $noArchiveTypes -notcontains $item.Extension))
        {
            $relativePath = Get-RelativePath -Path $item.FullName -RelativeTo $dataFolder

            $targetPath = Join-Path -Path $archiveFolder -ChildPath $relativePath
            $targetFolder = Split-Path -Path $targetPath -Parent

            if (-not (Test-Path -Path $targetFolder -PathType Container))
            {
                try
                {
                    $null = New-Item -Path $targetFolder -ItemType Directory -ErrorAction Stop
                }
                catch
                {
                    Write-ErrorLog -ErrorRecord $_
                    return
                }
            }

            try
            {
                Write-Debug "Archiving $($item.FullName)"

                Move-Item -Path $item.FullName -Destination $targetPath -Force -ErrorAction Stop
                $null = cmd /c mklink $item.FullName $targetPath
            }
            catch
            {
                Write-ErrorLog -ErrorRecord $_
                return
            }
        }
    }
}

#endregion