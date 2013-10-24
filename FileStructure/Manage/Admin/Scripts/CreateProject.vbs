' ***********************************************
' CreateProject.vbs
'
' Prompts the user to enter a new project name, and copies a template folder into
' a new folder of that name.
'
' Parameters:
'   Path:  Path to the folder where a new project is to be created.  This can either be
'          a folder named "Projects", or a folder that contains a folder named "Projects".
'   
'   ConfigFile:  Optional.  Path to an INI file which contains script options.  If this
'                parameter is not specified, it defaults to a file named "config.ini" in the
'                same folder as the script.
'
' Notes:
'   The INI file must contain a [Configuration] section.  In that section, there must be a
'   value named ProjectTemplate, which contains the path to the template folder.  If a relative
'   path is specified, it is relative in relation to the folder where this script is located.
' ***********************************************

Option Explicit

Const ForReading = 1

Dim strIniPath, strProjectsFolder, strTemplatePath, strProjectName, strProjectPath
Dim strScriptFolder, strScriptFile
Dim objFSO

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Determine the script's folder
strScriptFile = WScript.ScriptFullName

Dim objFile
Set objFile = objFSO.GetFile(strScriptFile)
strScriptFolder = objFSO.GetParentFolderName(objFile) 
Set objFile = Nothing

' Validate user input.  Make sure the values reference by the Path and ConfigFile parameters exist.

If (WScript.Arguments.Named.Exists("ConfigFile")) Then
    strIniPath = WScript.Arguments.Named("ConfigFile")
Else
    strIniPath = objFSO.BuildPath(strScriptFolder, "config.ini")
End If

If (WScript.Arguments.Named.Exists("Path")) Then
    strProjectsFolder = WScript.Arguments.Named("Path")
Else
    MsgBox "Error: Required Parameter 'Path' missing.", vbOkOnly, "Required parameter missing."
    WScript.Quit 1
End If

Dim reProjects
Set reProjects = New RegExp
reProjects.IgnoreCase = True
reProjects.Pattern = "\\Projects\\?$"

If (Not reProjects.Test(strProjectsFolder)) Then
    strProjectsFolder = objFSO.BuildPath(strProjectsFolder, "Projects")
End If

If (Not objFSO.FolderExists(strProjectsFolder)) Then
    MsgBox "Error: Projects directory '" & strProjectsFolder & "' does not exist.", vbOkOnly, "Projects directory not found."
    WScript.Quit 1
End If

If (Not objFSO.FileExists(strIniPath)) Then
    MsgBox "Error: Configuration file '" & strIniPath & "' does not exist.", vbOkOnly, "Config file not found."
    WScript.Quit 1
End If

' Read the INI file, and make sure the project template folder exists.

Dim objIniFile, objConfigSection
Set objIniFile = ReadIniFile(strIniPath)

If (objIniFile Is Nothing) Then
    MsgBox "Error: Configuration file '" & strIniPath & "' exists, but could not be read.", vbOkOnly, "Could not read config file."
    WScript.Quit 1    
End If

If (Not objIniFile.Exists("Configuration")) Then
    MsgBox "Error: Configuration file '" & strIniPath & "' does not contain the required [Configuration] section", vbOkOnly, "Invalid config file."
    WScript.Quit 1    
End If

Set objConfigSection = objIniFile("Configuration")

If (Not objConfigSection.Exists("ProjectTemplate")) Then
    MsgBox "Error: Configuration file '" & strIniPath & "' does not contain the required ProjectTemplate value.", vbOkOnly, "Invalid config file."
    WScript.Quit 1
End If

strTemplatePath = objConfigSection("ProjectTemplate")

Dim reRootedPath
Set reRootedPath = New RegExp
reRootedPath.IgnoreCase = True
reRootedPath.Pattern = "^(?:(?:[a-z]:)|(?:\\\\[^\\]+\\))"

If (Not reRootedPath.Test(strTemplatePath)) Then
    strTemplatePath = objFSO.BuildPath(strScriptFolder, strTemplatePath)
    strTemplatePath = objFSO.GetAbsolutePathName(strTemplatePath)
End If

If (Not objFSO.FolderExists(strTemplatePath)) Then
    MsgBox "Error: Project template directory '" & strTemplatePath & "' does not exist.", vbOkOnly, "Project template directory not found."
    WScript.Quit 1
End If

' Prompt the user to enter a project name.  If it is invalid or already exists, prompt the user to try again.
' If the user enters nothing (or clicks cancel), exit the script.

Dim reInvalidCharacters
Set reInvalidCharacters = New RegExp
reInvalidCharacters.Pattern = "[\\/:\*\?""\<\>\|]"

Do While (True)
    ' Per client requirements, the project name is suggested to begin with "YYYYMM-", but not required.
    ' The InputBox sets this value as the default; the user can keep or overwrite it, as desired.

    Dim intMonth, intYear
    Dim strPrefix
    
    intYear = DatePart("yyyy", Now)
    intMonth = DatePart("m", Now)
    
    If (intMonth < 10) Then
        strPrefix = intYear & "0" & intMonth & "-"
    Else
        strPrefix = intYear & intMonth & "-"
    End If
    
    strProjectName = InputBox("Enter a project name", "Enter New Project Name", strPrefix)

    If (strProjectName = "") Then
        WScript.Quit 0
    End If
    
    If (reInvalidCharacters.Test(strProjectName)) Then
        MsgBox "Value '" & strProjectName & "' contains invalid characters.  Please try again.", vbOkOnly, "Invalid Project Name"
    Else
        strProjectPath = objFSO.BuildPath(strProjectsFolder, strProjectName)

        If (objFSO.FolderExists(strProjectPath)) Then
            MsgBox "A project named '" & strProjectName & "' already exists in this location.  Please try again.", vbOkOnly, "Project Already Exists"
        Else
            Exit Do
        End If
    End If
Loop

' The user has entered a valid project name.  Attempt to copy the template directory into the new project folder.

On Error Resume Next

objFSO.CopyFolder strTemplatePath, strProjectPath

If (Err.Number <> 0) Then
    MsgBox "Error: Could not create or populate new project folder:" & vbCrLf & _
           "0x" & Hex(Err.Number) & ", " & Err.Description, _
           vbOkOnly, "Error Creating New Project"
End If

On Error Goto 0

'
' ReadIniFile
'
' Reads an INI file from disk, returning a Dictionary of Dictionaries.  Each section name is a key of the root dictionary, which maps
' to another dictionary containing the key=value pairs.
'
' If the file cannot be read, this function returns Nothing; the calling code should test for this as an indication of failure.
'
' Parameters:
'   strPath: Path to the INI file which is to be read.
'

Function ReadIniFile(strPath)
    Set ReadIniFile = Nothing
    
    Dim objDictionary, objFSO, objFile
    
    Set objDictionary = CreateObject("Scripting.Dictionary")
    objDictionary.CompareMode = vbTextCompare
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If (Not objFSO.FileExists(strPath)) Then
        Exit Function
    End If

    On Error Resume Next
    Set objFile = objFSO.OpenTextFile(strPath, ForReading)

    If (Err.Number <> 0) Then
        Exit Function
    End If
    On Error Goto 0
    
    Dim strLine
    Dim reSection, reValue
    Dim colMatches, objMatch
    Dim blnMatched
    
    Set reSection = New RegExp
    reSection.Pattern = "^\s*\[(.+)\]\s*$"
    
    Set reValue = New RegExp
    reValue.Pattern = "^\s*(.+?)\s*=\s*(.*?)\s*$"
    
    Dim strCurrentSection, strKey, strValue
    strCurrentSection = ""
    
    Do While (Not objFile.AtEndOfStream)
        strLine = objFile.ReadLine()
        
        blnMatched = False
        
        Set colMatches = reSection.Execute(strLine)
        For Each objMatch In colMatches
            blnMatched = True
            strCurrentSection = objMatch.SubMatches(0)
            
            If (Not objDictionary.Exists(strCurrentSection)) Then
                Dim objTemp
                Set objTemp = CreateObject("Scripting.Dictionary")
                objTemp.CompareMode = vbTextCompare
                
                Set objDictionary(strCurrentSection) = objTemp
            End If
        Next
        
        If ((Not blnMatched) And strCurrentSection <> "") Then
            Set colMatches = reValue.Execute(strLine)
            For Each objMatch In colMatches
                strKey = objMatch.SubMatches(0)
                strValue = objMatch.SubMatches(1)

                objDictionary(strCurrentSection)(strKey) = strValue
            Next
        End If
    Loop

    objFile.Close()
    Set objFile = Nothing

    Set ReadIniFile = objDictionary
End Function

