Attribute VB_Name = "modBuilderUtils"
Option Explicit

' =====================================================================================
' UTILITY FUNCTIONS FOR VBA BUILD SYSTEM
' =====================================================================================

' Read the full contents of a text file into a string
Public Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    ReadTextFile = Input$(LOF(fileNum), #fileNum)
    Close #fileNum
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    ReadTextFile = ""
End Function

' Get the VBProject that hosts modBuildSystem (fallback to ActiveVBProject)
Public Function GetVBProject() As Object
    Dim vbProj As Object, vbComp As Object
    For Each vbProj In Application.VBE.VBProjects
        For Each vbComp In vbProj.VBComponents
            If vbComp.Name = "modBuildSystem" Then
                Set GetVBProject = vbProj
                Exit Function
            End If
        Next vbComp
    Next vbProj
    Set GetVBProject = Application.VBE.ActiveVBProject
End Function

' Remove a component if it exists in the project
Public Sub RemoveComponent(vbProj As Object, componentName As String)
    On Error Resume Next
    Dim comp As Object: Set comp = vbProj.VBComponents(componentName)
    If Not comp Is Nothing Then vbProj.VBComponents.Remove comp
End Sub

' Force save of the host document so recent VBE changes persist
Public Sub ForceProjectStateSave()
    On Error Resume Next
    
    Dim hostDoc As Object
    Set hostDoc = GetVBProject().Parent
    If Not hostDoc Is Nothing And hostDoc.Path <> "" And Not hostDoc.Saved Then
        hostDoc.Save
    End If
End Sub

' Confirm that the “Trust access to VBA project object model” option is enabled
Public Function ValidateTrustSettings() As Boolean
    On Error GoTo ErrHandler
    Dim test As Object: Set test = Application.VBE.ActiveVBProject
    ValidateTrustSettings = True
    Exit Function
ErrHandler:
    MsgBox "VBA project access disabled. Enable 'Trust access to VBA project object model' in Trust Center.", vbCritical
    ValidateTrustSettings = False
End Function

' Registry functions for source path persistence
Public Function GetSourcePath() As String
    GetSourcePath = GetSetting("VBABuilder", "Config", "SourcePath", "")
End Function

Public Sub SaveSourcePath(path As String)
    SaveSetting "VBABuilder", "Config", "SourcePath", path
End Sub

' Prompt the user to select a folder when SourcePath is not yet stored
Public Function PromptForSourcePath() As String
    On Error GoTo UseInputBox
    Dim fd As Object: Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker
    With fd
        .Title = "Select VBA Source Folder"
        .AllowMultiSelect = False
        If .Show = -1 Then PromptForSourcePath = .SelectedItems(1)
    End With
    Exit Function
UseInputBox:
    PromptForSourcePath = InputBox("Enter source folder path:", "VBA Builder", "C:\YourProject\src")
End Function

' Return a list of sub-folders that contain a manifest.json
Public Function GetAvailableApps() As Collection
    Dim apps As New Collection, fso As Object, fld As Object, subFld As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If GetSourcePath() = "" Or Not fso.FolderExists(GetSourcePath()) Then Set GetAvailableApps = apps: Exit Function
    Set fld = fso.GetFolder(GetSourcePath())
    For Each subFld In fld.SubFolders
        If Dir(subFld.Path & "\manifest.json") <> "" Then apps.Add subFld.Name
    Next subFld
    Set GetAvailableApps = apps
End Function

' Show system status and available commands
Public Sub ShowSystemStatus()
    Dim msg As String
    msg = "=== VBA Builder Status ===" & vbCrLf & vbCrLf
    msg = msg & "Version: 1.0.0 (Refactored)" & vbCrLf
    msg = msg & "Source Path: " & IIf(GetSourcePath() = "", "(not set)", GetSourcePath()) & vbCrLf
    
    Dim apps As Collection
    Set apps = GetAvailableApps()
    msg = msg & "Available Apps: " & apps.Count & vbCrLf & vbCrLf
    
    msg = msg & "Commands:" & vbCrLf
    msg = msg & "• Initialize() - Setup/change source folder" & vbCrLf
    msg = msg & "• Build() - Show menu and build an app" & vbCrLf
    msg = msg & "• BuildApplication(""AppName"") - Build specific app"
    
    MsgBox msg, vbInformation, "VBA Builder Status"
End Sub

' Set the source path to a new location (internal use)
Private Sub SetSourcePath(newPath As String)
    ' Remove trailing backslash if present
    If Right(newPath, 1) = "\" Then newPath = Left(newPath, Len(newPath) - 1)
    
    ' Validate the path exists
    If Dir(newPath, vbDirectory) = "" Then
        MsgBox "Error: Path does not exist: " & newPath, vbCritical, "Invalid Path"
        Exit Sub
    End If
    
    ' Save the new path
    SaveSourcePath newPath
    
    ' Show confirmation with available apps
    Dim apps As Collection
    Set apps = GetAvailableApps()
    
    Dim msg As String
    msg = "✅ Source path updated successfully!" & vbCrLf & vbCrLf
    msg = msg & "New Path: " & newPath & vbCrLf
    msg = msg & "Available Apps: " & apps.Count
    
    If apps.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Found applications:"
        Dim i As Integer
        For i = 1 To apps.Count
            msg = msg & vbCrLf & "• " & apps(i)
        Next i
    End If
    
    MsgBox msg, vbInformation, "Source Path Updated"
End Sub
