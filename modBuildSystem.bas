Attribute VB_Name = "modBuildSystem"
' =====================================================================================
' VBA APPLICATION BUILDER - SIMPLIFIED BUILD SYSTEM
' =====================================================================================
' Version: 0.0.5 - Refactored and simplified
'
' QUICK START:
' 1. Call Initialize() to setup
' 2. Call BuildApplication("AppName") to build
'
' FEATURES:
' • Direct form creation via VBA object model
' • Lightweight JSON parsing for manifests and designs
' • Automatic module and form importing
' • Minimal dependencies for maximum compatibility
' =====================================================================================

Option Explicit

' Win32 API for waiting
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' =====================================================================================
' CONSTANTS
' =====================================================================================
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_StdModule As Long = 1
Private Const SLEEP_DELAY As Long = 500
Private Const CONTROL_DELAY As Long = 10

' =====================================================================================
' MODULE VARIABLES
' =====================================================================================
Private sourcePath As String

' =====================================================================================
' PUBLIC API
' =====================================================================================

' Initialize the build system - prompts for source folder
Public Sub Initialize()
    On Error GoTo ErrorHandler
    
    If Not ValidateTrustSettings() Then Exit Sub
    
    ' Always prompt for source path
    Dim newPath As String
    newPath = PromptForSourcePath()
    If newPath <> "" Then
        SaveSourcePath newPath
        MsgBox "✅ VBA Builder initialized!" & vbCrLf & "Source: " & newPath, vbInformation
    Else
        MsgBox "Initialization cancelled.", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error initializing: " & Err.Description, vbCritical
End Sub

' Build application from source
Public Sub BuildApplication(appName As String)
    On Error GoTo ErrorHandler
    
    ' Auto-initialize if no source path set
    sourcePath = GetSourcePath()
    If sourcePath = "" Then
        Call Initialize
        sourcePath = GetSourcePath()
        If sourcePath = "" Then Exit Sub
    End If
    
    Dim appPath As String
    appPath = sourcePath & "\" & appName
    
    If Dir(appPath & "\manifest.json") = "" Then
        MsgBox "Application not found: " & appPath, vbExclamation
        Exit Sub
    End If
    
    MsgBox "Building: " & appName, vbInformation
    
    Dim manifest As Object
    Set manifest = LoadJSON(appPath & "\manifest.json")
    If manifest Is Nothing Then Exit Sub
    
    ' Build components
    Dim modulesSuccess As Boolean, formsSuccess As Boolean
    
    modulesSuccess = ProcessModules(manifest, appPath)
    formsSuccess = ProcessForms(manifest, appPath)
    
    If modulesSuccess And formsSuccess Then
        ' Create user-friendly success message
        Dim successMsg As String
        successMsg = "✅ Build completed successfully!" & vbCrLf & vbCrLf & _
                    "Application: " & appName & vbCrLf & vbCrLf & _
                    "🚀 To test your form:" & vbCrLf & _
                    "   • Type: UserForm1.Show" & vbCrLf & _
                    "   • Press Enter in Immediate window" & vbCrLf & vbCrLf & _
                    "💡 Check Immediate window for detailed logs"
        
        MsgBox successMsg, vbInformation, "VBA Builder - Success!"
    Else
        MsgBox "❌ Build failed. Check Immediate window for details.", vbCritical
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Build error: " & Err.Description, vbCritical
End Sub

' Build - shows available apps and lets you choose
Public Sub Build()
    On Error GoTo ErrorHandler
    
    ' Auto-initialize if needed
    sourcePath = GetSourcePath()
    If sourcePath = "" Then
        Call Initialize
        sourcePath = GetSourcePath()
        If sourcePath = "" Then Exit Sub
    End If
    
    Dim apps As Collection
    Set apps = GetAvailableApps()
    
    If apps.Count = 0 Then
        MsgBox "No applications found in: " & sourcePath & vbCrLf & vbCrLf & "Run Initialize() to change the source folder.", vbExclamation
        Exit Sub
    End If
    
    Dim msg As String, i As Integer
    msg = "Select Application to Build:" & vbCrLf & vbCrLf
    For i = 1 To apps.Count
        msg = msg & i & ". " & apps(i) & vbCrLf
    Next i
    
    Dim choice As String
    choice = InputBox(msg, "VBA Builder", "1")
    If choice = "" Then Exit Sub
    
    If IsNumeric(choice) Then
        Dim sel As Integer
        sel = CInt(choice)
        If sel >= 1 And sel <= apps.Count Then
            Call BuildApplication(apps(sel))
        End If
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Build error: " & Err.Description, vbCritical
End Sub

' =====================================================================================
' JSON PARSING - SIMPLIFIED
' =====================================================================================

' Load and parse JSON file
Private Function LoadJSON(filePath As String) As Object
    On Error GoTo ErrorHandler
    
    
    Dim content As String
    content = ReadTextFile(filePath)
    If content = "" Then Set LoadJSON = Nothing: Exit Function
    
    Set LoadJSON = ParseJSON(content)
    Exit Function
    
ErrorHandler:
    Debug.Print "Error loading JSON: " & Err.Description
    Set LoadJSON = Nothing
End Function

' Simple JSON parser for manifest and design files
Private Function ParseJSON(jsonText As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Check if this is a design file with form section
    If InStr(jsonText, """form""") > 0 Then
        ' Extract form properties from form section
        Dim formSection As String
        formSection = ExtractFormSection(jsonText)
        
        dict("name") = GetJsonString(formSection, "name")
        dict("caption") = GetJsonString(formSection, "caption")
        dict("width") = CStr(GetJsonNumber(formSection, "width"))
        dict("height") = CStr(GetJsonNumber(formSection, "height"))
        dict("startUpPosition") = CStr(GetJsonNumber(formSection, "startUpPosition"))
    Else
        ' Legacy format or manifest file - extract from root level
        dict("name") = GetJsonString(jsonText, "name")
        dict("version") = GetJsonString(jsonText, "version")
        dict("caption") = GetJsonString(jsonText, "caption")
        dict("width") = CStr(GetJsonNumber(jsonText, "width"))
        dict("height") = CStr(GetJsonNumber(jsonText, "height"))
        dict("startUpPosition") = CStr(GetJsonNumber(jsonText, "startUpPosition"))
    End If
    
    ' Extract arrays (convert to comma-separated strings)
    Dim arrVar As Variant
    dict("modules") = GetJsonString(jsonText, "modules")
    If dict("modules") = "" Then
        arrVar = GetJsonStringArray(jsonText, "modules")
        If IsArray(arrVar) Then dict("modules") = Join(arrVar, ",")
    End If
    
    dict("forms") = GetJsonString(jsonText, "forms")
    If dict("forms") = "" Then
        arrVar = GetJsonStringArray(jsonText, "forms")
        If IsArray(arrVar) Then dict("forms") = Join(arrVar, ",")
    End If
    
    ' Extract controls array as raw text for processing
    dict("controls") = ExtractJsonArrayText(jsonText, "controls")
    
    Set ParseJSON = dict
End Function

' ===== JSON helper functions (schema specific, no error handling for speed) =====
Private Function GetJsonString(json As String, key As String) As String
    Dim p As Long, q As Long, colonPos As Long, quote As String
    quote = Chr$(34)
    p = InStr(json, quote & key & quote) ' key position
    If p = 0 Then Exit Function
    colonPos = InStr(p, json, ":")
    If colonPos = 0 Then Exit Function
    ' find first quote after the colon
    p = InStr(colonPos, json, quote) + 1
    q = InStr(p, json, quote)
    GetJsonString = Mid$(json, p, q - p)
End Function

Private Function GetJsonNumber(json As String, key As String) As Long
    Dim p As Long, colonPos As Long, quote As String: quote = Chr$(34)
    p = InStr(json, quote & key & quote)
    If p = 0 Then Exit Function
    colonPos = InStr(p, json, ":")
    If colonPos = 0 Then Exit Function
    
    ' Find the start of the number
    p = colonPos + 1
    Dim ch As String
    Do While p <= Len(json)
        ch = Mid$(json, p, 1)
        If ch Like "[0-9-]" Then Exit Do
        p = p + 1
    Loop
    
    ' Find the end of the number
    Dim endPos As Long
    endPos = p
    Do While endPos <= Len(json)
        ch = Mid$(json, endPos, 1)
        If Not ch Like "[0-9]" Then Exit Do
        endPos = endPos + 1
    Loop
    
    ' Extract just the number part
    Dim numberStr As String
    numberStr = Mid$(json, p, endPos - p)
    
    
    GetJsonNumber = CLng(numberStr)
End Function

' Extract the form section from design JSON
Private Function ExtractFormSection(jsonText As String) As String
    Dim startPos As Long, endPos As Long, braceCount As Long
    Dim i As Long, char As String, inString As Boolean
    
    ' Find "form": {
    startPos = InStr(jsonText, """form"":")
    If startPos = 0 Then Exit Function
    
    ' Find the opening brace
    startPos = InStr(startPos, jsonText, "{")
    If startPos = 0 Then Exit Function
    
    ' Find the matching closing brace
    braceCount = 1
    inString = False
    
    For i = startPos + 1 To Len(jsonText)
        char = Mid(jsonText, i, 1)
        
        If char = """" Then inString = Not inString
        
        If Not inString Then
            If char = "{" Then
                braceCount = braceCount + 1
            ElseIf char = "}" Then
                braceCount = braceCount - 1
                If braceCount = 0 Then
                    endPos = i
                    Exit For
                End If
            End If
        End If
    Next i
    
    If endPos > startPos Then
        ExtractFormSection = Mid(jsonText, startPos, endPos - startPos + 1)
    End If
End Function

Private Function GetJsonStringArray(json As String, key As String) As Variant
    Dim arrText As String, items() As String, i As Long
    arrText = ExtractJsonArrayText(json, key)
    If arrText = "" Then Exit Function
    arrText = Replace(arrText, """", "")
    items = Split(arrText, ",")
    For i = 0 To UBound(items)
        items(i) = Trim$(items(i))
    Next i
    GetJsonStringArray = items
End Function

Private Function ExtractJsonArrayText(json As String, key As String) As String
    Dim p As Long, startPos As Long, endPos As Long, quote As String
    quote = Chr$(34)
    p = InStr(json, quote & key & quote & ":")
    If p = 0 Then Exit Function
    startPos = InStr(p, json, "[") + 1
    endPos = InStr(startPos, json, "]")
    If startPos = 0 Or endPos = 0 Then Exit Function
    ExtractJsonArrayText = Mid$(json, startPos, endPos - startPos)
End Function


' Parse string array
Private Function ParseStringArray(arrayText As String) As Collection
    Dim coll As New Collection
    Dim items As Variant
    items = Split(arrayText, ",")
    
    Dim i As Integer
    For i = 0 To UBound(items)
        Dim item As String
        item = Trim(items(i))
        item = Replace(item, """", "")
        item = Replace(item, "'", "")
        item = Trim(item)
        If item <> "" Then coll.Add item
    Next i
    
    Set ParseStringArray = coll
End Function

' Parse controls array from JSON text
Private Function ParseControlsArray(arrayText As String) As Collection
    Dim coll As New Collection
    
    Dim controlObjects As Collection
    Set controlObjects = SplitJSONControlObjects(arrayText)
    
    Dim i As Integer
    For i = 1 To controlObjects.Count
        Dim controlDict As Object
        Set controlDict = ParseControlObject(controlObjects(i))
        If Not controlDict Is Nothing Then
            coll.Add controlDict
        End If
    Next i
    
    Set ParseControlsArray = coll
End Function

' Split JSON control objects more robustly
Private Function SplitJSONControlObjects(arrayText As String) As Collection
    Dim objects As New Collection
    Dim i As Integer
    Dim braceCount As Integer
    Dim startPos As Integer
    Dim inString As Boolean
    Dim char As String
    
    startPos = 1
    braceCount = 0
    inString = False
    
    ' Find first opening brace
    For i = 1 To Len(arrayText)
        char = Mid(arrayText, i, 1)
        If char = "{" Then
            startPos = i
            Exit For
        End If
    Next i
    
    Dim objStart As Integer
    objStart = startPos
    
    For i = startPos To Len(arrayText)
        char = Mid(arrayText, i, 1)
        
        ' Track string boundaries
        If char = """" Then inString = Not inString
        
        If Not inString Then
            If char = "{" Then
                braceCount = braceCount + 1
            ElseIf char = "}" Then
                braceCount = braceCount - 1
                
                ' Complete object found
                If braceCount = 0 Then
                    Dim objText As String
                    objText = Mid(arrayText, objStart, i - objStart + 1)
                    objects.Add objText
                    
                    ' Find next object start
                    For objStart = i + 1 To Len(arrayText)
                        If Mid(arrayText, objStart, 1) = "{" Then Exit For
                    Next objStart
                    
                    If objStart <= Len(arrayText) Then
                        i = objStart - 1 ' Will be incremented by For loop
                    End If
                End If
            End If
        End If
    Next i
    
    Set SplitJSONControlObjects = objects
End Function

' Parse single control object
Private Function ParseControlObject(controlText As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Extract basic properties
    dict("name") = GetJsonString(controlText, "name")
    dict("type") = GetJsonString(controlText, "type")
    dict("caption") = GetJsonString(controlText, "caption")
    
    ' Extract numeric position values
    Dim leftVal As String, topVal As String, widthVal As String, heightVal As String
    leftVal = CStr(GetJsonNumber(controlText, "left"))
    topVal = CStr(GetJsonNumber(controlText, "top"))
    widthVal = CStr(GetJsonNumber(controlText, "width"))
    heightVal = CStr(GetJsonNumber(controlText, "height"))
    
    If IsNumeric(leftVal) Then dict("left") = CLng(leftVal)
    If IsNumeric(topVal) Then dict("top") = CLng(topVal)
    If IsNumeric(widthVal) Then dict("width") = CLng(widthVal)
    If IsNumeric(heightVal) Then dict("height") = CLng(heightVal)
    
    Set ParseControlObject = dict
End Function

' =====================================================================================
' FORM CREATION - SIMPLIFIED DIRECT APPROACH
' =====================================================================================

' Process forms from manifest
Private Function ProcessForms(manifest As Object, appPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not manifest.Exists("forms") Or manifest("forms") = "" Then
        ProcessForms = True
        Exit Function
    End If
    
    Dim forms As Variant
    forms = Split(manifest("forms"), ",")
    
    Dim i As Integer
    For i = 0 To UBound(forms)
        Dim formName As String
        formName = Trim(forms(i))
        If formName = "" Then GoTo NextForm
        
        If Not CreateFormDirect(formName, appPath) Then
            ProcessForms = False
            Exit Function
        End If
        
NextForm:
    Next i
    
    ProcessForms = True
    Exit Function
    
ErrorHandler:
    ProcessForms = False
End Function

' Create form directly via code
Private Function CreateFormDirect(formName As String, appPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim vbProj As Object
    Set vbProj = GetVBProject()
    
    ' Remove existing form if it exists
    Call RemoveComponent(vbProj, formName)
    
    ' Create new form
    Dim formComp As Object
    Set formComp = vbProj.VBComponents.Add(vbext_ct_MSForm)
    
    ' Apply design first (makes form "dirty" for easier renaming)
    Dim designPath As String
    designPath = appPath & "\forms\" & formName & "\design.json"
    
    If Dir(designPath) <> "" Then
        Dim design As Object
        Set design = LoadJSON(designPath)
        If Not design Is Nothing Then
            Call ApplyDesign(formComp, design)
        End If
    End If
    
    ' Allow VBE to process changes
    DoEvents
    Sleep SLEEP_DELAY
    
    ' Attempt to rename form
    On Error Resume Next
    formComp.Properties("Name").Value = formName
    If Err.Number <> 0 Then
        Err.Clear
        formComp.Name = formName ' Fallback
    End If
    On Error GoTo ErrorHandler
    
    ' Add code-behind if available
    Dim codePath As String
    codePath = appPath & "\forms\" & formName & "\code-behind.vba"
    
    If Dir(codePath) <> "" Then
        Dim codeContent As String
        codeContent = ReadTextFile(codePath)
        If codeContent <> "" Then
            formComp.CodeModule.AddFromString codeContent
        End If
    End If
    
    ' Save changes
    Call ForceProjectStateSave
    
    Debug.Print "Form created: " & formComp.Name & " (use " & formComp.Name & ".Show)"
    CreateFormDirect = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error creating form " & formName & ": " & Err.Description
    CreateFormDirect = False
End Function

' Apply design to form
Private Sub ApplyDesign(formComp As Object, design As Object)
    On Error Resume Next ' Don't fail build for design issues
    
    ' Set basic properties
    If design.Exists("caption") And design("caption") <> "" Then
        formComp.Properties("Caption").Value = design("caption")
    End If
    
    ' Set dimensions
    If design.Exists("width") And IsNumeric(design("width")) And CLng(design("width")) > 0 Then
        formComp.Properties("Width").Value = CSng(design("width"))
    End If
    If design.Exists("height") And IsNumeric(design("height")) And CLng(design("height")) > 0 Then
        formComp.Properties("Height").Value = CSng(design("height"))
    End If
    
    ' Set startup position
    If design.Exists("startUpPosition") And IsNumeric(design("startUpPosition")) Then
        formComp.Properties("StartUpPosition").Value = CInt(design("startUpPosition"))
    End If
    
    On Error GoTo ControlsError
    
    ' Create controls
    If design.Exists("controls") And design("controls") <> "" Then
        If TypeName(design("controls")) = "String" Then
            Set design("controls") = ParseControlsArray(design("controls"))
        End If
        Call CreateControls(formComp.Designer, design("controls"))
    End If
    Exit Sub
    
ControlsError:
    Debug.Print "Warning: Controls creation failed: " & Err.Description
End Sub

' Create controls on form
Private Sub CreateControls(formObj As Object, controlsData As Object)
    On Error Resume Next
    
    If TypeName(controlsData) = "Collection" Then
        Dim i As Integer
        For i = 1 To controlsData.Count
            Dim controlDict As Object
            Set controlDict = controlsData(i)
            If Not controlDict Is Nothing Then
                Call CreateSingleControl(formObj, controlDict)
            End If
        Next i
    End If
End Sub

' Create individual control from dictionary
Private Sub CreateSingleControl(formObj As Object, controlDict As Object)
    On Error Resume Next
    
    Dim ctrlName As String, ctrlType As String, caption As String
    ctrlName = "Control1"
    ctrlType = "CommandButton"
    
    If controlDict.Exists("name") Then ctrlName = controlDict("name")
    If controlDict.Exists("type") Then ctrlType = controlDict("type")
    If controlDict.Exists("caption") Then caption = controlDict("caption")
    
    ' Create the control
    Dim ctrl As Object
    Set ctrl = formObj.Controls.Add(GetControlType(ctrlType), ctrlName)
    
    ' Apply properties
    If caption <> "" Then ctrl.Caption = caption
    If controlDict.Exists("left") Then ctrl.Left = CLng(controlDict("left"))
    If controlDict.Exists("top") Then ctrl.Top = CLng(controlDict("top"))
    If controlDict.Exists("width") Then ctrl.Width = CLng(controlDict("width"))
    If controlDict.Exists("height") Then ctrl.Height = CLng(controlDict("height"))

    ' Allow VBE to process control
    DoEvents
    Sleep CONTROL_DELAY
End Sub

' Get control type string for VBA
Private Function GetControlType(ctrlType As String) As String
    Select Case ctrlType
        Case "CommandButton": GetControlType = "Forms.CommandButton.1"
        Case "Label": GetControlType = "Forms.Label.1"
        Case "TextBox": GetControlType = "Forms.TextBox.1"
        Case "ListBox": GetControlType = "Forms.ListBox.1"
        Case "ComboBox": GetControlType = "Forms.ComboBox.1"
        Case "CheckBox": GetControlType = "Forms.CheckBox.1"
        Case "OptionButton": GetControlType = "Forms.OptionButton.1"
        Case "Frame": GetControlType = "Forms.Frame.1"
        Case "Image": GetControlType = "Forms.Image.1"
        Case Else: GetControlType = "Forms.CommandButton.1" ' Default fallback
    End Select
End Function

' =====================================================================================
' MODULE PROCESSING
' =====================================================================================

' Process modules from manifest
Private Function ProcessModules(manifest As Object, appPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not manifest.Exists("modules") Or manifest("modules") = "" Then
        ProcessModules = True
        Exit Function
    End If
    
    Dim modules As Variant
    modules = Split(manifest("modules"), ",")
    
    Dim i As Integer
    For i = 0 To UBound(modules)
        Dim moduleName As String
        moduleName = Trim(modules(i))
        If moduleName = "" Then GoTo NextModule
        
        Dim modulePath As String
        modulePath = appPath & "\modules\" & moduleName & ".vba"
        
        If Dir(modulePath) = "" Then
            MsgBox "Module not found: " & modulePath, vbCritical
            ProcessModules = False
            Exit Function
        End If
        
        If Not ImportModule(moduleName, modulePath) Then
            ProcessModules = False
            Exit Function
        End If
        
NextModule:
    Next i
    
    ProcessModules = True
    Exit Function
    
ErrorHandler:
    ProcessModules = False
End Function

' Import module from file
Private Function ImportModule(moduleName As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim vbProj As Object
    Set vbProj = GetVBProject()
    
    ' Remove existing
    Call RemoveComponent(vbProj, moduleName)
    
    ' Import new
    Dim comp As Object
    Set comp = vbProj.VBComponents.Import(filePath)
    If comp.Name <> moduleName Then comp.Name = moduleName
    
    ImportModule = True
    Exit Function
    
ErrorHandler:
    ImportModule = False
End Function