Attribute VB_Name = "modBuilderUtils"
Option Explicit

' =====================================================================================
' UTILITY FUNCTIONS FOR VBA BUILD SYSTEM
' =====================================================================================

' Read the full contents of a text file into a string
Public Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2 ' adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath
        ReadTextFile = .ReadText(-1) ' adReadAll
        .Close
    End With
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
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
    If Not hostDoc Is Nothing And hostDoc.Path <> "" Then
        If hostDoc.Saved Then hostDoc.Saved = False
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
    msg = msg & "Version: 1.0.1 (Refactored)" & vbCrLf
    msg = msg & "Source Path: " & IIf(GetSourcePath() = "", "(not set)", GetSourcePath()) & vbCrLf
    
    Dim apps As Collection
    Set apps = GetAvailableApps()
    msg = msg & "Available Apps: " & apps.Count & vbCrLf & vbCrLf
    
    msg = msg & "Commands:" & vbCrLf
    msg = msg & "• Build() - Main entry point (auto-initializes if needed)" & vbCrLf
    msg = msg & "• BuildApplication(""AppName"") - Build specific app directly"
    
    MsgBox msg, vbInformation, "VBA Builder Status"
End Sub
