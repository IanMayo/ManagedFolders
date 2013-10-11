#requires -Version 2.0

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [System.String]
    $Path
)

function Get-ProjectName
{
    # Prompts the user to enter a project name via a pop-up window.

    $Form = New-Object System.Windows.Forms.Form 
    $Form.Text = "Enter Project Name"
    $Form.Size = New-Object System.Drawing.Size(300,200) 
    $Form.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({ $Form.Close() })
    $OKButton.DialogResult = 'OK'
    $Form.Controls.Add($OKButton)
    $Form.AcceptButton = $OKButton

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({ $Form.Close() })
    $CancelButton.DialogResult = 'Cancel'
    $Form.Controls.Add($CancelButton)
    $Form.CancelButton = $CancelButton

    $Label = New-Object System.Windows.Forms.Label
    $Label.Location = New-Object System.Drawing.Size(10,20) 
    $Label.Size = New-Object System.Drawing.Size(280,20) 
    $Label.Text = "Enter a project name (should begin with YYYYMM- )"
    $Form.Controls.Add($Label) 

    $TextBox = New-Object System.Windows.Forms.TextBox 
    $TextBox.Location = New-Object System.Drawing.Size(10,40) 
    $TextBox.Size = New-Object System.Drawing.Size(260,20)
    $TextBox.Text = Get-Date -Format 'yyyyMM-'
    $Form.Controls.Add($TextBox) 

    $Form.Topmost = $True

    $Form.Add_Shown({ $Form.Activate(); $null = $TextBox.Focus(); $TextBox.SelectionStart = 100 })
    $result = $Form.ShowDialog()

    if ($result -eq 'OK')
    {
        return $TextBox.Text
    }
}

try
{
    $null = [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    $null = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
}
catch
{
    throw
}

$scriptFolder = Split-Path -Path $MyInvocation.MyCommand.Path -Parent

# Sanity check to make sure this script is being executed in the intended folder structure.  The Data folder doesn't have to exist, but if the
# script's folder doesn't end with "Admin\Scripts", throw an error.

# TODO: Per discussions with client, folder structure may be changed around a bit.  Script will be updated to get the root data folder from a config
# file instead of making assumptions about paths relative to the script's location.

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

if ($Path -notmatch '\\Projects\\?$')
{
    $Path = Join-Path -Path $Path -ChildPath 'Projects'
}

# Validate user input
if (-not (Test-Path -Path $Path -PathType Container))
{
    throw New-Object System.IO.DirectoryNotFoundException($Path)
}

$invalidChars = [System.IO.Path]::GetInvalidFileNameChars()

while ($true)
{
    $projectName = Get-ProjectName

    if ($null -eq $projectName)
    {
        exit 0
    }

    if ($projectName.IndexOfAny($invalidChars) -ge 0)
    {
        $null = [System.Windows.Forms.MessageBox]::Show("Value '$projectName' contains invalid characters.  Please try again.", 'Invalid Project Name', 'OK')
        continue
    }

    break
}

$projectPath = Join-Path -Path $Path -ChildPath $projectName

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
