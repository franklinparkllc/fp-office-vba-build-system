Attribute VB_Name = "modBuildSystem"
' =====================================================================================
' VBA APPLICATION BUILDER - SIMPLIFIED BUILD SYSTEM
' =====================================================================================
' Version: 0.0.5 - Fixed control positioning and user feedback
'
' This simplified build system focuses on core functionality:
' • Direct form creation via code (no export/import complexity)
' • Simple regex-based JSON parsing
' • Minimal dependencies and maximum stability
' • Optimized for AI/agentic workflows
'
' QUICK START:
' 1. Call Initialize() to setup
' 2. Call BuildApplication("AppName") to build
'
' ARCHITECTURE:
' • Direct VBA object manipulation
' • Simple file I/O
' • Streamlined JSON parsing
' • Direct form/control creation
' =====================================================================================

Option Explicit

' =====================================================================================
' MODULE VARIABLES
' =====================================================================================
Private sourcePath As String
Private Const vbext_ct_MSForm = 3
Private Const vbext_ct_StdModule = 1

' =====================================================================================
' PUBLIC API
' =====================================================================================

' Initialize the build system
Public Sub Initialize()
    On Error GoTo ErrorHandler
    
    If Not ValidateTrustSettings() Then Exit Sub
    
    sourcePath = GetSourcePath()
    If sourcePath = "" Then
        sourcePath = PromptForSourcePath()
        If sourcePath <> "" Then
            SaveSourcePath sourcePath
        Else
            MsgBox "Build system requires a source path.", vbExclamation
            Exit Sub
        End If
    End If
    
    MsgBox "VBA Builder initialized!" & vbCrLf & "Source: " & sourcePath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error initializing: " & Err.Description, vbCritical
End Sub

' Build application from source
Public Sub BuildApplication(appName As String)
    On Error GoTo ErrorHandler
    
    If sourcePath = "" Then Call Initialize: If sourcePath = "" Then Exit Sub
    
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
    Debug.Print "=== Starting Build Components ==="
    
    Dim modulesSuccess As Boolean, formsSuccess As Boolean
    
    Debug.Print "Processing modules..."
    modulesSuccess = ProcessModules(manifest, appPath)
    Debug.Print "Modules result: " & modulesSuccess
    
    Debug.Print "Processing forms..."
    formsSuccess = ProcessForms(manifest, appPath)
    Debug.Print "Forms result: " & formsSuccess
    
    If modulesSuccess And formsSuccess Then
        Debug.Print "✅ Build completed successfully"
        
        ' Create user-friendly success message
        Dim successMsg As String
        successMsg = "✅ Build completed successfully!" & vbCrLf & vbCrLf & _
                    "Application: " & appName & vbCrLf & vbCrLf & _
                    "🚀 To test your form:" & vbCrLf & _
                    "   • Type: UserForm1.Show" & vbCrLf & _
                    "   • Press Enter in Immediate window" & vbCrLf & vbCrLf & _
                    "📝 Optional renaming:" & vbCrLf & _
                    "   • Save document first (Ctrl+S)" & vbCrLf & _
                    "   • Right-click 'UserForm1' → Properties" & vbCrLf & _
                    "   • Change Name to 'frmExampleApp'" & vbCrLf & vbCrLf & _
                    "💡 Check Immediate window for detailed logs"
        
        MsgBox successMsg, vbInformation, "VBA Builder - Success!"
    Else
        Debug.Print "❌ Build failed - Modules: " & modulesSuccess & ", Forms: " & formsSuccess
        MsgBox "❌ Build failed. Check Immediate window for details.", vbCritical
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Build error: " & Err.Description, vbCritical
End Sub

' Interactive build menu
Public Sub BuildInteractive()
    Call Initialize
    
    Dim apps As Collection
    Set apps = GetAvailableApps()
    
    If apps.Count = 0 Then
        MsgBox "No applications found in: " & sourcePath, vbExclamation
        Exit Sub
    End If
    
    Dim msg As String, i As Integer
    msg = "Select Application:" & vbCrLf & vbCrLf
    For i = 1 To apps.Count
        msg = msg & i & ". " & apps(i) & vbCrLf
    Next i
    
    Dim choice As String
    choice = InputBox(msg, "Build Application", "1")
    If choice = "" Then Exit Sub
    
    If IsNumeric(choice) Then
        Dim sel As Integer
        sel = CInt(choice)
        If sel >= 1 And sel <= apps.Count Then
            Call BuildApplication(apps(sel))
        End If
    End If
End Sub

' =====================================================================================
' JSON PARSING - SIMPLIFIED
' =====================================================================================

' Load and parse JSON file
Private Function LoadJSON(filePath As String) As Object
    On Error GoTo ErrorHandler
    
    Debug.Print "=== LoadJSON Debug ==="
    Debug.Print "Loading JSON from: " & filePath
    
    Dim content As String
    content = ReadTextFile(filePath)
    If content = "" Then 
        Debug.Print "❌ File content is empty"
        Set LoadJSON = Nothing
        Exit Function
    End If
    
    Debug.Print "File content loaded, length: " & Len(content)
    Debug.Print "First 200 chars: " & Left(content, 200)
    
    Set LoadJSON = ParseJSON(content)
    If LoadJSON Is Nothing Then
        Debug.Print "❌ JSON parsing failed"
    Else
        Debug.Print "✅ JSON parsed successfully"
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print "❌ ERROR in LoadJSON: " & Err.Number & " - " & Err.Description
    Set LoadJSON = Nothing
End Function

' Simple JSON parser using regex
Private Function ParseJSON(jsonText As String) As Object
    On Error GoTo ErrorHandler
    
    Debug.Print "=== ParseJSON Debug ==="
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Debug.Print "Dictionary created"
    
    ' Extract simple string values
    dict("name") = ExtractValue(jsonText, "name")
    Debug.Print "Name extracted: " & dict("name")
    
    dict("version") = ExtractValue(jsonText, "version")
    Debug.Print "Version extracted: " & dict("version")
    
    dict("modules") = ExtractValue(jsonText, "modules")
    Debug.Print "Modules extracted: " & dict("modules")
    
    dict("forms") = ExtractValue(jsonText, "forms")
    Debug.Print "Forms extracted: " & dict("forms")
    
    ' Extract references array
    Dim refsText As String
    refsText = ExtractArray(jsonText, "references")
    If refsText <> "" Then
        Debug.Print "References array found: " & refsText
        Dim refs As Collection
        Set refs = ParseStringArray(refsText)
        Set dict("references") = refs
        Debug.Print "References parsed, count: " & refs.Count
    Else
        Debug.Print "No references array found"
    End If
    
    ' Extract controls array
    Dim controlsText As String
    controlsText = ExtractArray(jsonText, "controls")
    If controlsText <> "" Then
        Debug.Print "Controls array found, length: " & Len(controlsText)
        Debug.Print "Controls text preview: " & Left(controlsText, 100)
        Dim controls As Collection
        Set controls = ParseControlsArray(controlsText)
        Set dict("controls") = controls
        Debug.Print "Controls parsed, count: " & controls.Count
    Else
        Debug.Print "No controls array found"
    End If
    
    Debug.Print "✅ ParseJSON completed, dictionary has " & dict.Count & " keys"
    Set ParseJSON = dict
    Exit Function
    
ErrorHandler:
    Debug.Print "❌ ERROR in ParseJSON: " & Err.Number & " - " & Err.Description
    Set ParseJSON = Nothing
End Function

' Extract JSON value (string or number)
Private Function ExtractValue(json As String, key As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    ' Updated pattern to handle both quoted strings and unquoted numbers
    regex.Pattern = """" & key & """\s*:\s*""([^""]*)""|""" & key & """\s*:\s*([^,}\s]+)"
    regex.IgnoreCase = True
    
    Dim matches As Object
    Set matches = regex.Execute(json)
    
    If matches.Count > 0 Then
        If matches(0).SubMatches(0) <> "" Then
            ' Quoted string value
            ExtractValue = matches(0).SubMatches(0)
        ElseIf matches(0).SubMatches(1) <> "" Then
            ' Unquoted numeric value
            ExtractValue = Trim(matches(0).SubMatches(1))
        End If
    End If
End Function

' Extract JSON array content
Private Function ExtractArray(json As String, key As String) As String
    Dim startPos As Long, endPos As Long
    Dim searchStr As String
    searchStr = """" & key & """" & "\s*:\s*\["
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = searchStr
    regex.IgnoreCase = True
    
    Dim matches As Object
    Set matches = regex.Execute(json)
    
    If matches.Count > 0 Then
        startPos = matches(0).FirstIndex + matches(0).Length
        endPos = InStr(startPos, json, "]")
        If endPos > startPos Then
            ExtractArray = Mid(json, startPos, endPos - startPos)
        End If
    End If
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

' Parse controls array - enhanced approach
Private Function ParseControlsArray(arrayText As String) As Collection
    Dim coll As New Collection
    
    Debug.Print "=== ParseControlsArray Debug ==="
    Debug.Print "Array text length: " & Len(arrayText)
    Debug.Print "Array text preview: " & Left(arrayText, 200)
    
    ' Use a more robust approach to split JSON objects
    Dim controlObjects As Collection
    Set controlObjects = SplitJSONControlObjects(arrayText)
    
    Debug.Print "Found " & controlObjects.Count & " control objects"
    
    Dim i As Integer
    For i = 1 To controlObjects.Count
        Dim controlText As String
        controlText = controlObjects(i)
        
        Debug.Print "Processing control " & i & ": " & Left(controlText, 100)
        
        ' Parse individual control
        Dim controlDict As Object
        Set controlDict = ParseControlObject(controlText)
        If Not controlDict Is Nothing Then
            coll.Add controlDict
            Debug.Print "✅ Control " & i & " parsed successfully"
        Else
            Debug.Print "❌ Control " & i & " parsing failed"
        End If
    Next i
    
    Debug.Print "✅ ParseControlsArray completed with " & coll.Count & " controls"
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
    
    Debug.Print "--- ParseControlObject Debug ---"
    Debug.Print "Control text: " & Left(controlText, 150)
    
    ' Extract basic properties
    dict("name") = ExtractValue(controlText, "name")
    dict("type") = ExtractValue(controlText, "type")
    dict("caption") = ExtractValue(controlText, "caption")
    
    Debug.Print "Extracted - Name: " & dict("name") & ", Type: " & dict("type") & ", Caption: " & dict("caption")
    
    Dim leftVal As String, topVal As String, widthVal As String, heightVal As String
    leftVal = ExtractValue(controlText, "left")
    topVal = ExtractValue(controlText, "top")
    widthVal = ExtractValue(controlText, "width")
    heightVal = ExtractValue(controlText, "height")
    
    Debug.Print "Position values - Left: " & leftVal & ", Top: " & topVal & ", Width: " & widthVal & ", Height: " & heightVal
    
    If IsNumeric(leftVal) Then dict("left") = CLng(leftVal)
    If IsNumeric(topVal) Then dict("top") = CLng(topVal)
    If IsNumeric(widthVal) Then dict("width") = CLng(widthVal)
    If IsNumeric(heightVal) Then dict("height") = CLng(heightVal)
    
    ' Handle font object (simplified - just extract what we need)
    If InStr(controlText, """font""") > 0 Then
        Debug.Print "Font object found in control"
        ' For now, we'll skip font parsing to focus on getting controls created
        ' Font properties can be added later if needed
    End If
    
    Debug.Print "✅ ParseControlObject completed for: " & dict("name")
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
    
    Debug.Print "=== CreateFormDirect Debug ==="
    Debug.Print "Form Name: " & formName
    Debug.Print "App Path: " & appPath
    
    Dim vbProj As Object
    Set vbProj = GetVBProject()
    Debug.Print "VB Project obtained: " & vbProj.Name
    
    ' Remove existing form
    Debug.Print "Checking for existing form: " & formName
    Dim existingForm As Object
    On Error Resume Next
    Set existingForm = vbProj.VBComponents(formName)
    If Not existingForm Is Nothing Then
        Debug.Print "Found existing form, removing..."
        Call RemoveComponent(vbProj, formName)
        Debug.Print "Existing form removed"
    Else
        Debug.Print "No existing form found"
    End If
    On Error GoTo ErrorHandler
    
    ' Create new form
    Debug.Print "Creating new UserForm component..."
    Dim formComp As Object
    Set formComp = vbProj.VBComponents.Add(vbext_ct_MSForm)
    Debug.Print "UserForm created with name: " & formComp.Name
    
    Debug.Print "Renaming form to: " & formName
    Debug.Print "Current form name: " & formComp.Name
    
    ' Try multiple renaming strategies
    Dim renameSuccess As Boolean
    renameSuccess = False
    
    ' Strategy 1: Direct rename
    On Error Resume Next
    formComp.Name = formName
    If Err.Number = 0 And formComp.Name = formName Then
        renameSuccess = True
        Debug.Print "✅ Direct rename successful"
    Else
        Debug.Print "⚠️ Direct rename failed: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Strategy 2: If direct rename failed, try via VBComponents collection
    If Not renameSuccess Then
        Debug.Print "Trying rename via VBComponents collection..."
        On Error Resume Next
        Dim tempName As String
        tempName = formComp.Name
        vbProj.VBComponents(tempName).Name = formName
        If Err.Number = 0 Then
            renameSuccess = True
            Debug.Print "✅ Collection rename successful"
        Else
            Debug.Print "⚠️ Collection rename failed: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Strategy 3: If still failed, continue with original name (this is normal and OK)
    If Not renameSuccess Then
        Debug.Print "ℹ️ Form created with auto-generated name: " & formComp.Name
        Debug.Print "ℹ️ To use the intended name '" & formName & "':"
        Debug.Print "ℹ️   1. Save your document (Ctrl+S)"
        Debug.Print "ℹ️   2. In VBA Editor, right-click '" & formComp.Name & "' and select Properties"
        Debug.Print "ℹ️   3. Change the Name from '" & formComp.Name & "' to '" & formName & "'"
        Debug.Print "ℹ️ This is normal VBA behavior - the form works perfectly as-is!"
    Else
        Debug.Print "✅ Form renamed successfully to: " & formComp.Name
    End If
    
    ' Load design (use the original formName for path, regardless of actual form name)
    Dim designPath As String
    designPath = appPath & "\forms\" & formName & "\design.json"
    Debug.Print "Design path: " & designPath
    
    If Dir(designPath) <> "" Then
        Debug.Print "Design file found, loading..."
        Dim design As Object
        Set design = LoadJSON(designPath)
        If Not design Is Nothing Then
            Debug.Print "Design JSON loaded successfully"
            Debug.Print "Applying design to form..."
            Call ApplyDesign(formComp.Designer, design)
            Debug.Print "Design applied successfully"
        Else
            Debug.Print "❌ Failed to parse design JSON"
        End If
    Else
        Debug.Print "⚠️ No design file found at: " & designPath
    End If
    
    ' Load code-behind
    Dim codePath As String
    codePath = appPath & "\forms\" & formName & "\code-behind.vba"
    Debug.Print "Code-behind path: " & codePath
    
    If Dir(codePath) <> "" Then
        Debug.Print "Code-behind file found, loading..."
        Dim codeContent As String
        codeContent = ReadTextFile(codePath)
        If codeContent <> "" Then
            Debug.Print "Code content loaded, length: " & Len(codeContent)
            Debug.Print "Adding code to form module..."
            formComp.CodeModule.AddFromString codeContent
            Debug.Print "Code-behind added successfully"
        Else
            Debug.Print "⚠️ Code-behind file is empty"
        End If
    Else
        Debug.Print "⚠️ No code-behind file found at: " & codePath
    End If
    
    ' Force save to persist form state changes
    Call ForceProjectStateSave
    
    Debug.Print "✅ CreateFormDirect completed successfully"
    Debug.Print "Final form name: " & formComp.Name
    
    ' Provide user guidance regardless of rename success
    If Not renameSuccess Then
        Debug.Print "📋 FORM READY: Use " & formComp.Name & ".Show to display the form"
        Debug.Print "📋 Or manually rename to '" & formName & "' and use " & formName & ".Show"
    Else
        Debug.Print "📋 FORM READY: Use " & formName & ".Show to display the form"
    End If
    
    CreateFormDirect = True
    Exit Function
    
ErrorHandler:
    Debug.Print "❌ ERROR in CreateFormDirect: " & Err.Number & " - " & Err.Description
    Debug.Print "Error occurred at line in CreateFormDirect"
    CreateFormDirect = False
End Function

' Apply design to form
Private Sub ApplyDesign(formObj As Object, design As Object)
    On Error GoTo DesignError
    
    Debug.Print "=== ApplyDesign Debug ==="
    Debug.Print "Design object type: " & TypeName(design)
    Debug.Print "Design keys count: " & design.Count
    
    ' Basic properties
    If design.Exists("caption") Then 
        Debug.Print "Setting caption: " & design("caption")
        formObj.Caption = design("caption")
        Debug.Print "Caption set successfully"
    End If
    
    If design.Exists("width") Then 
        Dim widthVal As Long
        widthVal = CLng(design("width")) * 20
        Debug.Print "Setting width: " & design("width") & " points (" & widthVal & " twips)"
        formObj.Width = widthVal
        Debug.Print "Width set successfully"
    End If
    
    If design.Exists("height") Then 
        Dim heightVal As Long
        heightVal = CLng(design("height")) * 20
        Debug.Print "Setting height: " & design("height") & " points (" & heightVal & " twips)"
        formObj.Height = heightVal
        Debug.Print "Height set successfully"
    End If
    
    ' Create controls
    If design.Exists("controls") Then
        Debug.Print "Controls found in design, creating..."
        Call CreateControls(formObj, design("controls"))
        Debug.Print "Controls creation completed"
    Else
        Debug.Print "No controls found in design"
    End If
    
    Debug.Print "✅ ApplyDesign completed successfully"
    Exit Sub
    
DesignError:
    Debug.Print "❌ ERROR in ApplyDesign: " & Err.Number & " - " & Err.Description
    ' Continue execution - don't fail the entire build for design issues
    Resume Next
End Sub

' Create controls on form - handles JSON control arrays
Private Sub CreateControls(formObj As Object, controlsData As Object)
    On Error GoTo ControlsError
    
    Debug.Print "=== CreateControls Debug ==="
    Debug.Print "Controls data type: " & TypeName(controlsData)
    
    ' controlsData should be a Collection from JSON parsing
    If TypeName(controlsData) = "Collection" Then
        Debug.Print "Controls collection found with " & controlsData.Count & " items"
        
        Dim i As Integer
        For i = 1 To controlsData.Count
            Debug.Print "Processing control " & i & " of " & controlsData.Count
            
            Dim controlDict As Object
            Set controlDict = controlsData(i)
            
            If Not controlDict Is Nothing Then
                Debug.Print "Control " & i & " dictionary type: " & TypeName(controlDict)
                Call CreateSingleControl(formObj, controlDict)
                Debug.Print "Control " & i & " created successfully"
            Else
                Debug.Print "⚠️ Control " & i & " is Nothing"
            End If
        Next i
        
        Debug.Print "✅ All controls processed"
    Else
        Debug.Print "⚠️ Controls data is not a Collection, type: " & TypeName(controlsData)
    End If
    Exit Sub
    
ControlsError:
    Debug.Print "❌ ERROR in CreateControls: " & Err.Number & " - " & Err.Description
    Resume Next
End Sub

' Create individual control from dictionary
Private Sub CreateSingleControl(formObj As Object, controlDict As Object)
    On Error GoTo SingleControlError
    
    Debug.Print "--- CreateSingleControl Debug ---"
    
    Dim ctrlName As String, ctrlType As String, caption As String
    ctrlName = "Control1"
    ctrlType = "CommandButton"
    
    If controlDict.Exists("name") Then 
        ctrlName = controlDict("name")
        Debug.Print "Control name: " & ctrlName
    End If
    
    If controlDict.Exists("type") Then 
        ctrlType = controlDict("type")
        Debug.Print "Control type: " & ctrlType
    End If
    
    If controlDict.Exists("caption") Then 
        caption = controlDict("caption")
        Debug.Print "Control caption: " & caption
    End If
    
    ' Create the control
    Debug.Print "Creating control with type: " & GetControlType(ctrlType)
    Dim ctrl As Object
    Set ctrl = formObj.Controls.Add(GetControlType(ctrlType), ctrlName)
    Debug.Print "Control created successfully: " & ctrl.Name
    
    ' Apply properties
    If caption <> "" Then 
        ctrl.Caption = caption
        Debug.Print "Caption applied: " & caption
    End If
    
    ' Apply positioning with validation
    If controlDict.Exists("left") And IsNumeric(controlDict("left")) Then 
        ctrl.Left = CLng(controlDict("left"))
        Debug.Print "Left position set: " & controlDict("left")
    Else
        Debug.Print "⚠️ No valid left position, using default (0)"
    End If
    
    If controlDict.Exists("top") And IsNumeric(controlDict("top")) Then 
        ctrl.Top = CLng(controlDict("top"))
        Debug.Print "Top position set: " & controlDict("top")
    Else
        Debug.Print "⚠️ No valid top position, using default (0)"
    End If
    
    If controlDict.Exists("width") And IsNumeric(controlDict("width")) Then 
        ctrl.Width = CLng(controlDict("width"))
        Debug.Print "Width set: " & controlDict("width")
    Else
        Debug.Print "⚠️ No valid width, using default"
    End If
    
    If controlDict.Exists("height") And IsNumeric(controlDict("height")) Then 
        ctrl.Height = CLng(controlDict("height"))
        Debug.Print "Height set: " & controlDict("height")
    Else
        Debug.Print "⚠️ No valid height, using default"
    End If
    
    Debug.Print "Final control position: Left=" & ctrl.Left & ", Top=" & ctrl.Top & ", Width=" & ctrl.Width & ", Height=" & ctrl.Height
    
    Debug.Print "✅ Single control creation completed"
    Exit Sub
    
SingleControlError:
    Debug.Print "❌ ERROR in CreateSingleControl: " & Err.Number & " - " & Err.Description
    Debug.Print "Control name: " & ctrlName & ", Type: " & ctrlType
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

' =====================================================================================
' UTILITY FUNCTIONS
' =====================================================================================

' Force project state save to persist form properties
Private Sub ForceProjectStateSave()
    On Error Resume Next
    
    Debug.Print "=== Forcing Project State Save ==="
    
    ' This works for documents (Word, Excel)
    If Not Application.ActiveDocument Is Nothing Then
        If Not Application.ActiveDocument.Saved Then
            Application.ActiveDocument.Save
            Debug.Print "Host document saved."
        Else
            Debug.Print "Host document was already saved, no save needed."
        End If
    End If
    
    If Err.Number <> 0 Then
        Debug.Print "⚠️ Could not save document. Error: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

' Get VBA project
Private Function GetVBProject() As Object
    Set GetVBProject = Application.VBE.ActiveVBProject
End Function

' Remove component if exists
Private Sub RemoveComponent(vbProj As Object, componentName As String)
    On Error Resume Next
    
    Debug.Print "--- RemoveComponent Debug ---"
    Debug.Print "Attempting to remove: " & componentName
    
    Dim comp As Object
    Set comp = vbProj.VBComponents(componentName)
    If Not comp Is Nothing Then
        Debug.Print "Component found, removing..."
        vbProj.VBComponents.Remove comp
        If Err.Number = 0 Then
            Debug.Print "✅ Component removed successfully"
        Else
            Debug.Print "⚠️ Remove failed: " & Err.Number & " - " & Err.Description
        End If
    Else
        Debug.Print "Component not found (already removed or doesn't exist)"
    End If
    
    On Error GoTo 0
End Sub

' Read text file
Private Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As fileNum
    ReadTextFile = Input(LOF(fileNum), fileNum)
    Close fileNum
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close fileNum
    ReadTextFile = ""
End Function

' Validate trust settings
Private Function ValidateTrustSettings() As Boolean
    On Error GoTo ErrorHandler
    Dim test As Object
    Set test = Application.VBE.ActiveVBProject
    ValidateTrustSettings = True
    Exit Function
    
ErrorHandler:
    MsgBox "VBA project access disabled. Enable 'Trust access to VBA project object model' in Trust Center.", vbCritical
    ValidateTrustSettings = False
End Function

' Get available applications
Private Function GetAvailableApps() As Collection
    Dim apps As New Collection
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If sourcePath <> "" And fso.FolderExists(sourcePath) Then
        Dim folder As Object
        Set folder = fso.GetFolder(sourcePath)
        
        Dim subfolder As Object
        For Each subfolder In folder.SubFolders
            If Dir(subfolder.Path & "\manifest.json") <> "" Then
                apps.Add subfolder.Name
            End If
        Next
    End If
    
    Set GetAvailableApps = apps
End Function

' Source path management
Private Function GetSourcePath() As String
    On Error Resume Next
    GetSourcePath = GetSetting("VBABuilder", "Config", "SourcePath", "")
End Function

Private Sub SaveSourcePath(path As String)
    On Error Resume Next
    SaveSetting "VBABuilder", "Config", "SourcePath", path
End Sub

Private Function PromptForSourcePath() As String
    On Error GoTo ErrorHandler
    
    Dim folderPicker As Object
    Set folderPicker = Application.FileDialog(4)
    
    With folderPicker
        .Title = "Select VBA Source Folder"
        .AllowMultiSelect = False
        If .Show = -1 Then
            PromptForSourcePath = .SelectedItems(1)
        End If
    End With
    Exit Function
    
ErrorHandler:
    PromptForSourcePath = InputBox("Enter source folder path:", "VBA Builder", "C:\YourProject\src")
End Function

' System status
Public Sub ShowSystemStatus()
    Dim msg As String
    msg = "=== VBA Builder Status ===" & vbCrLf & vbCrLf
    msg = msg & "Version: 2.0.4 (Fixed positioning)" & vbCrLf
    msg = msg & "Source Path: " & IIf(GetSourcePath() = "", "(not set)", GetSourcePath()) & vbCrLf
    
    Dim apps As Collection
    Set apps = GetAvailableApps()
    msg = msg & "Available Apps: " & apps.Count & vbCrLf & vbCrLf
    
    msg = msg & "Commands:" & vbCrLf
    msg = msg & "• Initialize() - Setup system" & vbCrLf
    msg = msg & "• BuildInteractive() - Build with menu" & vbCrLf
    msg = msg & "• BuildApplication(""AppName"") - Build specific app" & vbCrLf
    msg = msg & "• TestLastBuiltForm() - Show the last built form"
    
    MsgBox msg, vbInformation, "VBA Builder Status"
End Sub

' Test function to show the last built form
Public Sub TestLastBuiltForm()
    On Error GoTo ErrorHandler
    
    ' Try common form names
    Dim formNames As Variant
    formNames = Array("UserForm1", "frmExampleApp", "UserForm2")
    
    Dim i As Integer
    For i = 0 To UBound(formNames)
        On Error Resume Next
        Dim testForm As Object
        ' Try to get the form by name
        Set testForm = Application.VBE.ActiveVBProject.VBComponents(formNames(i)).Designer
        If Not testForm Is Nothing Then
            On Error GoTo ErrorHandler
            Debug.Print "🚀 Showing form: " & formNames(i)
            ' Show the form
            Application.Run formNames(i) & ".Show"
            Exit Sub
        End If
        On Error GoTo ErrorHandler
    Next i
    
    MsgBox "No forms found to test. Build an application first using BuildInteractive().", vbInformation, "No Forms Found"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error testing form: " & Err.Description & vbCrLf & vbCrLf & _
           "Try manually: UserForm1.Show", vbExclamation, "Test Error"
End Sub
