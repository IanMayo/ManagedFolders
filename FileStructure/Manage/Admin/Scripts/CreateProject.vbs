Option Explicit

Const ForReading = 1

Dim strIniPath, strProjectsFolder, strTemplatePath, strProjectName, strProjectPath
Dim strScriptFolder, strScriptFile
Dim objFSO, objFile

strScriptFile = WScript.ScriptFullName

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.GetFile(strScriptFile)
strScriptFolder = objFSO.GetParentFolderName(objFile) 
Set objFile = Nothing

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

Dim reInvalidCharacters
Set reInvalidCharacters = New RegExp
reInvalidCharacters.Pattern = "[\\/:\*\?""\<\>\|]"

Do While (True)
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

On Error Resume Next

objFSO.CopyFolder strTemplatePath, strProjectPath

If (Err.Number <> 0) Then
    MsgBox "Error: Could not create or populate new project folder:" & vbCrLf & _
           "0x" & Hex(Err.Number) & ", " & Err.Description, _
           vbOkOnly, "Error Creating New Project"
End If

On Error Goto 0

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

    Set ReadIniFile = objDictionary
End Function

