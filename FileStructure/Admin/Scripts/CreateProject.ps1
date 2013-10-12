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
    $null = [System.Windows.Forms.MessageBox]::Show(
        "Error: $($MyInvocation.ScriptName) script is not located in the expected folder structure (which must end in \admin\scripts\).",
        'Script Path Error',
        'OK'
    )

    exit 1
}

$rootFolder = $matches[1]

# Make sure the subject template folder exists
$projectTemplateFolder = Join-Path -Path $rootFolder -ChildPath 'Admin\ProjectTemplate'

if (-not (Test-Path -Path $projectTemplateFolder -PathType Container))
{
    $null = [System.Windows.Forms.MessageBox]::Show(
        "Error: Project template directory '$projectTemplateFolder' was not found.",
        'Project Template Folder Missing',
        'OK'
    )

    exit 1
}

if ($Path -notmatch '\\Projects\\?$')
{
    $Path = Join-Path -Path $Path -ChildPath 'Projects'
}

# Validate user input
if (-not (Test-Path -Path $Path -PathType Container))
{
    $null = [System.Windows.Forms.MessageBox]::Show(
        "Error: Projects directory '$Path' was not found.",
        'Projects Folder Missing',
        'OK'
    )

    exit 1
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
        $null = [System.Windows.Forms.MessageBox]::Show(
            "Value '$projectName' contains invalid characters.  Please try again.",
            'Invalid Project Name',
            'OK'
        )

        continue
    }
    elseif ($projectName -eq [string]::Empty)
    {
        $null = [System.Windows.Forms.MessageBox]::Show(
            "Project name cannot be blank.  Please try again.",
            'Invalid Project Name',
            'OK'
        )

        continue        
    }
    elseif (Test-Path -Path ($projectPath = Join-Path -Path $Path -ChildPath $projectName))
    {
        $null = [System.Windows.Forms.MessageBox]::Show(
            "A project named '$projectName' already exists in this location.  Please try again.",
            'Project Already Exists',
            'OK'
        )

        continue
    }

    break
}

try
{
    $null = New-Item -Path $projectPath -ItemType Directory -ErrorAction Stop
    Copy-Item -Path (Join-Path -Path $projectTemplateFolder -ChildPath '*') -Destination $projectPath -Recurse -Force -Container -ErrorAction Stop
}
catch
{
    $null = [System.Windows.Forms.MessageBox]::Show(
        "Error:  Could not create or populate new project folder:`r`n$($_.Exception.Message)",
        'Error Creating New Project',
        'OK'
    )

    exit 1
}

exit 0
