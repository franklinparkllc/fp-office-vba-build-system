Attribute VB_Name = "modAppBuilder"
' =====================================================================================
' VBA APPLICATION BUILDER - UNIFIED BUILD SYSTEM
' =====================================================================================
' Version: 0.2.0 - Direct app folder selection
'
' QUICK START:
' 1. Import this module into your VBA project
' 2. Call Build() - prompts for app folder and builds it
' 3. Call BuildApplication("C:\Path\To\App") - builds specific app folder
'
' FEATURES:
' • Direct form creation via VBA object model
' • Lightweight JSON parsing with validation and comment support
' • Automatic module and form importing
' • Enhanced error reporting and progress tracking
' • Auto-save and quality of life improvements
' • Direct app folder selection - no parent folder confusion
' =====================================================================================

Option Explicit

' Win32 API for waiting
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
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
Private buildProgress As String
Private totalSteps As Integer
Private currentStep As Integer

' =====================================================================================
' PUBLIC API
' =====================================================================================

' Build application from source
Public Sub BuildApplication(Optional appPath As String = "")
    On Error GoTo ErrorHandler
    
    ' Initialize progress tracking
    ResetProgress
    
    ' Always prompt for app folder if not provided
    If appPath = "" Then
        appPath = PromptForAppFolder()
        If appPath = "" Then
            MsgBox "Build cancelled - no app folder selected.", vbInformation
            Exit Sub
        End If
    End If
    
    UpdateProgress "Validating application path..."
    
    ' Check if manifest.json exists in the selected folder
    If Dir(appPath & "\manifest.json") = "" Then
        ShowError "BuildApplication", 0, "No manifest.json found in: " & appPath, _
                  "Please select a folder containing manifest.json"
        Exit Sub
    End If
    
    UpdateProgress "Loading manifest..."
    
    Dim manifest As Object
    Set manifest = LoadJSON(appPath & "\manifest.json")
    If manifest Is Nothing Then
        ShowError "BuildApplication", 0, "Failed to load manifest.json", _
                  "Check that manifest.json is valid JSON format."
        Exit Sub
    End If
    
    ' Calculate total steps for progress
    CalculateBuildSteps manifest
    
    ' Get app name from manifest
    Dim appName As String
    appName = manifest("name")
    If appName = "" Then appName = "Application"
    
    MsgBox "Building: " & appName & vbCrLf & _
           "Source: " & appPath & vbCrLf & vbCrLf & _
           "Check Immediate window for progress...", vbInformation
    
    ' Process references from manifest (if any)
    UpdateProgress "Processing references..."
    Call ProcessReferences(manifest)

    ' Build components
    Dim modulesSuccess As Boolean, formsSuccess As Boolean
    
    modulesSuccess = ProcessModules(manifest, appPath)
    formsSuccess = ProcessForms(manifest, appPath)
    
    If modulesSuccess And formsSuccess Then
        ' Auto-save if enabled
        If GetAutoSaveEnabled() Then
            UpdateProgress "Auto-saving project..."
            Call ForceProjectStateSave
        End If
        
        ' Create user-friendly success message
        Dim successMsg As String
        Dim formLaunchHints As String
        Dim builtForms As Variant
        formLaunchHints = ""
        If manifest.Exists("forms") And manifest("forms") <> "" Then
            builtForms = Split(manifest("forms"), ",")
            Dim f As Integer
            For f = 0 To UBound(builtForms)
                If Trim(builtForms(f)) <> "" Then
                    formLaunchHints = formLaunchHints & "   - Type: " & Trim(builtForms(f)) & ".Show" & vbCrLf
                End If
            Next f
        End If
        If formLaunchHints = "" Then formLaunchHints = "   - Type: <YourFormName>.Show" & vbCrLf
        
        successMsg = "Build completed successfully!" & vbCrLf & vbCrLf & _
                    "Application: " & appName & vbCrLf & vbCrLf & _
                    "To test your form:" & vbCrLf & _
                    formLaunchHints & _
                    "   - Press Enter in Immediate window" & vbCrLf & vbCrLf & _
                    "Steps completed: " & currentStep & " of " & totalSteps
        
        MsgBox successMsg, vbInformation, "VBA App Builder - Success!"
        UpdateProgress "Build completed successfully!"
    Else
        ShowError "BuildApplication", 0, "Build failed", _
                  "Check Immediate window for detailed error information."
    End If
    Exit Sub
    
ErrorHandler:
    ShowError "BuildApplication", Err.Number, Err.Description, _
              "Unexpected error during build process."
End Sub

' Build - prompts for app folder and builds it
Public Sub Build()
    On Error GoTo ErrorHandler
    
    ' Simply call BuildApplication which will prompt for folder
    Call BuildApplication
    Exit Sub
    
    ErrorHandler:
    ShowError "Build", Err.Number, Err.Description
End Sub

' Configure auto-save preference
Public Sub ConfigureAutoSave()
    Dim current As Boolean
    current = GetAutoSaveEnabled()
    
    Dim msg As String
    msg = "Auto-save is currently " & IIf(current, "ENABLED", "DISABLED") & vbCrLf & vbCrLf & _
          "When enabled, the project will automatically save after successful builds." & vbCrLf & vbCrLf & _
          "Enable auto-save?"
    
    Dim result As VbMsgBoxResult
    result = MsgBox(msg, vbYesNoCancel + vbQuestion, "Configure Auto-Save")
    
    If result <> vbCancel Then
        SetAutoSaveEnabled (result = vbYes)
        MsgBox "Auto-save is now " & IIf(result = vbYes, "ENABLED", "DISABLED"), vbInformation
    End If
End Sub

' =====================================================================================
' ERROR HANDLING & USER FEEDBACK
' =====================================================================================

' Enhanced error display with context and recovery suggestions
Private Sub ShowError(context As String, errNum As Long, errDesc As String, Optional suggestion As String = "")
    Dim msg As String
    msg = "Error in " & context
    If errNum <> 0 Then msg = msg & " (#" & errNum & ")"
    msg = msg & vbCrLf & vbCrLf & "Details: " & errDesc
    
    If suggestion <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Suggestion: " & suggestion
    End If
    
    MsgBox msg, vbCritical, "VBA App Builder Error"
    Debug.Print "ERROR [" & context & "]: " & errDesc
End Sub

' =====================================================================================
' PROGRESS TRACKING
' =====================================================================================

' Reset progress tracking
Private Sub ResetProgress()
    buildProgress = ""
    totalSteps = 0
    currentStep = 0
End Sub

' Calculate total build steps from manifest
Private Sub CalculateBuildSteps(manifest As Object)
    totalSteps = 1 ' Initial validation
    
    If manifest.Exists("references") And manifest("references") <> "" Then
        totalSteps = totalSteps + 1
    End If
    
    If manifest.Exists("modules") And manifest("modules") <> "" Then
        Dim modules As Variant
        modules = Split(manifest("modules"), ",")
        totalSteps = totalSteps + UBound(modules) + 1
    End If
    
    If manifest.Exists("forms") And manifest("forms") <> "" Then
        Dim forms As Variant
        forms = Split(manifest("forms"), ",")
        totalSteps = totalSteps + UBound(forms) + 1
    End If
    
    If GetAutoSaveEnabled() Then totalSteps = totalSteps + 1
End Sub

' Update progress with current operation
Private Sub UpdateProgress(operation As String)
    currentStep = currentStep + 1
    Dim progressBar As String
    progressBar = String(10, "=")
    
    If totalSteps > 0 Then
        Dim pct As Integer
        pct = Int((currentStep / totalSteps) * 10)
        If pct > 10 Then pct = 10
        If pct < 0 Then pct = 0
        progressBar = String(pct, ChrW(&H2588)) & String(10 - pct, ChrW(&H2591))
    End If
    
    buildProgress = "[" & progressBar & "] " & operation
    Debug.Print buildProgress & " (" & currentStep & "/" & totalSteps & ")"
    DoEvents
End Sub

' =====================================================================================
' JSON PARSING WITH VALIDATION
' =====================================================================================

' Load and parse JSON file with enhanced error reporting
Private Function LoadJSON(filePath As String) As Object
    On Error GoTo ErrorHandler
    
    Debug.Print "Loading JSON from: " & filePath
    
    Dim content As String
    content = ReadTextFile(filePath)
    If content = "" Then
        Debug.Print "JSON Error: File is empty or couldn't be read: " & filePath
        Set LoadJSON = Nothing
        Exit Function
    End If
    
    Debug.Print "File content length: " & Len(content)
    Debug.Print "First 50 chars: " & Left(content, 50)
    
    ' Try parsing without comment stripping first
    On Error Resume Next
    Set LoadJSON = ParseJSON(content)
    If Not LoadJSON Is Nothing Then
        Debug.Print "Successfully parsed JSON without comment stripping"
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' If that failed, try with comment stripping
    Debug.Print "Trying with comment stripping..."
    content = StripJSONComments(content)
    
    ' Validate JSON structure
    Dim validationError As String
    validationError = ValidateJSONStructure(content)
    If validationError <> "" Then
        Debug.Print "JSON Validation Error in " & filePath & ":"
        Debug.Print validationError
        Set LoadJSON = Nothing
        Exit Function
    End If
    
    Set LoadJSON = ParseJSON(content)
    Exit Function
    
ErrorHandler:
    Debug.Print "Error loading JSON from " & filePath & ": " & Err.Description
    Set LoadJSON = Nothing
End Function

' Strip single-line comments from JSON
Private Function StripJSONComments(jsonText As String) As String
    On Error GoTo ErrorHandler
    
    Dim lines() As String
    Dim cleanLines() As String
    Dim i As Long
    Dim line As String
    Dim inString As Boolean
    Dim j As Long
    Dim char As String
    
    Debug.Print "StripJSONComments: Input length = " & Len(jsonText)
    
    ' Handle different line endings
    Dim normalizedText As String
    normalizedText = Replace(jsonText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    
    Debug.Print "StripJSONComments: After normalization length = " & Len(normalizedText)
    
    lines = Split(normalizedText, vbLf)
    Debug.Print "StripJSONComments: Number of lines = " & (UBound(lines) + 1)
    
    If UBound(lines) < 0 Then
        StripJSONComments = jsonText
        Exit Function
    End If
    ReDim cleanLines(UBound(lines))
    
    For i = 0 To UBound(lines)
        line = lines(i)
        inString = False
        
        ' Check each character to find comment start
        For j = 1 To Len(line)
            char = Mid(line, j, 1)
            
            ' Toggle string state
            If char = """" And (j = 1 Or Mid(line, j - 1, 1) <> "\") Then
                inString = Not inString
            End If
            
            ' Check for comment start
            If Not inString And j < Len(line) Then
                If Mid(line, j, 2) = "//" Then
                    line = Left(line, j - 1)
                    Exit For
                End If
            End If
        Next j
        
        cleanLines(i) = RTrim(line)
    Next i
    
    StripJSONComments = Join(cleanLines, vbLf)
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in StripJSONComments: " & Err.Description & " (Line: " & Erl & ")"
    StripJSONComments = jsonText ' Return original text if error
End Function

' Validate JSON structure and report specific errors
Private Function ValidateJSONStructure(jsonText As String) As String
    On Error GoTo ErrorHandler
    
    Dim braceCount As Long, bracketCount As Long
    Dim i As Long, lineNum As Long
    Dim char As String, prevChar As String
    Dim inString As Boolean
    Dim lines() As String
    
    Debug.Print "ValidateJSONStructure: Input length = " & Len(jsonText)
    
    ' Normalize line endings first
    Dim normalizedText As String
    normalizedText = Replace(jsonText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    
    lines = Split(normalizedText, vbLf)
    lineNum = 1
    
    For i = 1 To Len(normalizedText)
        char = Mid(normalizedText, i, 1)
        
        ' Track line numbers
        If char = vbLf Then lineNum = lineNum + 1
        
        ' Toggle string state
        If char = """" And prevChar <> "\" Then
            inString = Not inString
        End If
        
        If Not inString Then
            Select Case char
                Case "{"
                    braceCount = braceCount + 1
                Case "}"
                    braceCount = braceCount - 1
                    If braceCount < 0 Then
                        ValidateJSONStructure = "Line " & lineNum & ": Unexpected closing brace '}'"
                        Exit Function
                    End If
                Case "["
                    bracketCount = bracketCount + 1
                Case "]"
                    bracketCount = bracketCount - 1
                    If bracketCount < 0 Then
                        ValidateJSONStructure = "Line " & lineNum & ": Unexpected closing bracket ']'"
                        Exit Function
                    End If
            End Select
        End If
        
        prevChar = char
    Next i
    
    If braceCount <> 0 Then
        ValidateJSONStructure = "Mismatched braces: " & IIf(braceCount > 0, braceCount & " unclosed", Abs(braceCount) & " extra closing") & " brace(s)"
    ElseIf bracketCount <> 0 Then
        ValidateJSONStructure = "Mismatched brackets: " & IIf(bracketCount > 0, bracketCount & " unclosed", Abs(bracketCount) & " extra closing") & " bracket(s)"
    ElseIf inString Then
        ValidateJSONStructure = "Unterminated string literal"
    Else
        ValidateJSONStructure = ""
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in ValidateJSONStructure: " & Err.Description
    ValidateJSONStructure = "Validation error: " & Err.Description
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
    
    ' Add dependency parsing
    Dim depSection As String
    depSection = ExtractJsonObjectSection(jsonText, "dependencies")
    If depSection <> "" Then
        Dim refsArr As Variant
        refsArr = GetJsonStringArray(depSection, "references")
        If IsArray(refsArr) Then
            dict("references") = Join(refsArr, ",")
        Else
            Dim refsText As String
            refsText = ExtractJsonArrayText(depSection, "references")
            If refsText <> "" Then dict("references") = refsText
        End If
    End If
    
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

Private Function GetJsonBoolean(json As String, key As String) As Boolean
    Dim p As Long, colonPos As Long, quote As String: quote = Chr$(34)
    p = InStr(json, quote & key & quote)
    If p = 0 Then Exit Function
    colonPos = InStr(p, json, ":")
    If colonPos = 0 Then Exit Function
    
    Dim i As Long, ch As String
    i = colonPos + 1
    Do While i <= Len(json)
        ch = Mid$(json, i, 1)
        If ch Like "[A-Za-z]" Then Exit Do
        i = i + 1
    Loop
    GetJsonBoolean = (LCase$(Mid$(json, i, 4)) = "true")
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

Private Function ExtractJsonObjectSection(jsonText As String, key As String) As String
    Dim startPos As Long, endPos As Long, braceCount As Long
    Dim i As Long, char As String, inString As Boolean
    
    startPos = InStr(jsonText, Chr$(34) & key & Chr$(34) & ":")
    If startPos = 0 Then Exit Function
    
    startPos = InStr(startPos, jsonText, "{")
    If startPos = 0 Then Exit Function
    
    braceCount = 1
    inString = False
    
    For i = startPos + 1 To Len(jsonText)
        char = Mid$(jsonText, i, 1)
        If char = Chr$(34) Then inString = Not inString
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
        ExtractJsonObjectSection = Mid$(jsonText, startPos, endPos - startPos + 1)
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
    
    ' Extend control parsing to support fonts
    Dim fontObjText As String
    fontObjText = ExtractJsonObjectSection(controlText, "font")
    If fontObjText <> "" Then
        Dim fontName As String
        Dim fontSize As String
        Dim fontBold As Boolean
        Dim fontItalic As Boolean
        fontName = GetJsonString(fontObjText, "name")
        fontSize = CStr(GetJsonNumber(fontObjText, "size"))
        fontBold = GetJsonBoolean(fontObjText, "bold")
        fontItalic = GetJsonBoolean(fontObjText, "italic")
        If fontName <> "" Then dict("font.name") = fontName
        If IsNumeric(fontSize) Then dict("font.size") = CLng(fontSize)
        dict("font.bold") = fontBold
        dict("font.italic") = fontItalic
    End If
    
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
        
        UpdateProgress "Creating form: " & formName & "..."
        
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
    
    ' Warn on design form name mismatch in CreateFormDirect before ApplyDesign:
    If design.Exists("name") Then
        If Len(design("name")) > 0 And UCase$(design("name")) <> UCase$(formName) Then
            Debug.Print "Warning: design.json form name (" & design("name") & ") differs from manifest form ('" & formName & "'). Using manifest name."
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
    
    Debug.Print "✓ Form created: " & formComp.Name & " (use " & formComp.Name & ".Show)"
    CreateFormDirect = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error creating form " & formName & ": " & Err.Description
    ShowError "CreateFormDirect", Err.Number, "Failed to create form: " & formName, _
              "Check that form name is valid and design.json is properly formatted."
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

    ' Apply font in CreateSingleControl after size/height assignments:
    If controlDict.Exists("font.name") Then ctrl.Font.Name = CStr(controlDict("font.name"))
    If controlDict.Exists("font.size") Then ctrl.Font.Size = CLng(controlDict("font.size"))
    If controlDict.Exists("font.bold") Then ctrl.Font.Bold = CBool(controlDict("font.bold"))
    If controlDict.Exists("font.italic") Then ctrl.Font.Italic = CBool(controlDict("font.italic"))

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
        
        UpdateProgress "Importing module: " & moduleName & "..."
        
        Dim modulePath As String
        modulePath = appPath & "\modules\" & moduleName & ".vba"
        
        If Dir(modulePath) = "" Then
            ShowError "ProcessModules", 0, "Module not found: " & modulePath, _
                      "Ensure the module file exists at the specified location."
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
    
    Debug.Print "✓ Module imported: " & moduleName
    ImportModule = True
    Exit Function
    
ErrorHandler:
    ShowError "ImportModule", Err.Number, "Failed to import module: " & moduleName, _
              "Check file format and VBA project permissions."
    ImportModule = False
End Function

' Add reference processing
Private Sub ProcessReferences(manifest As Object)
    On Error Resume Next
    If manifest Is Nothing Then Exit Sub
    If Not manifest.Exists("references") Then Exit Sub
    If manifest("references") = "" Then Exit Sub
    
    Dim refs As Variant
    refs = Split(manifest("references"), ",")
    Dim i As Integer, refName As String
    For i = 0 To UBound(refs)
        refName = Trim$(refs(i))
        If refName <> "" Then EnsureReferenceByName refName
    Next i
End Sub

Private Sub EnsureReferenceByName(referenceName As String)
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = GetVBProject()
    If vbProj Is Nothing Then Exit Sub
    
    Dim r As Object
    For Each r In vbProj.References
        If StrComp(r.Description, referenceName, vbTextCompare) = 0 _
           Or StrComp(r.Name, referenceName, vbTextCompare) = 0 Then Exit Sub
        If InStr(1, referenceName, r.Description, vbTextCompare) > 0 Then Exit Sub
    Next r
    
    Select Case LCase$(referenceName)
        Case LCase$("Microsoft Forms 2.0 Object Library"), LCase$("MSForms"), LCase$("Forms 2.0")
            vbProj.References.AddFromGuid "{0D452EE1-E08F-101A-852E-02608C4D0BB4}", 2, 0
        Case Else
            Debug.Print "Reference not auto-added (unknown mapping): " & referenceName
    End Select
End Sub

' =====================================================================================
' UTILITY FUNCTIONS (Previously in modBuilderUtils)
' =====================================================================================

' Read the full contents of a text file into a string
Public Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    ' Try ADODB.Stream first (supports UTF-8)
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
    
    ' Fallback to standard VBA file reading
    On Error GoTo FileError
    Dim fileNum As Integer
    Dim fileContent As String
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    ReadTextFile = fileContent
    Exit Function
    
FileError:
    Debug.Print "Error reading file " & filePath & ": " & Err.Description
    ReadTextFile = ""
End Function

' Get the VBProject that hosts modAppBuilder (fallback to ActiveVBProject)
Public Function GetVBProject() As Object
    Dim vbProj As Object, vbComp As Object
    For Each vbProj In Application.VBE.VBProjects
        For Each vbComp In vbProj.VBComponents
            If vbComp.Name = "modAppBuilder" Then
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

' Confirm that the "Trust access to VBA project object model" option is enabled
Public Function ValidateTrustSettings() As Boolean
    On Error GoTo ErrHandler
    Dim test As Object: Set test = Application.VBE.ActiveVBProject
    ValidateTrustSettings = True
    Exit Function
ErrHandler:
    MsgBox "VBA project access disabled. Enable 'Trust access to VBA project object model' in Trust Center.", vbCritical
    ValidateTrustSettings = False
End Function


' Get auto-save preference
Private Function GetAutoSaveEnabled() As Boolean
    GetAutoSaveEnabled = (GetSetting("VBAAppBuilder", "Config", "AutoSave", "1") = "1")
End Function

' Set auto-save preference
Private Sub SetAutoSaveEnabled(value As Boolean)
    SaveSetting "VBAAppBuilder", "Config", "AutoSave", IIf(value, "1", "0")
End Sub

' Prompt the user to select an app folder
Public Function PromptForAppFolder() As String
    On Error GoTo UseInputBox
    Dim fd As Object: Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker
    With fd
        .Title = "Select App Folder (containing manifest.json)"
        .AllowMultiSelect = False
        If .Show = -1 Then PromptForAppFolder = .SelectedItems(1)
    End With
    Exit Function
UseInputBox:
    PromptForAppFolder = InputBox("Enter app folder path:", "VBA App Builder", "C:\YourProject\ExampleApp")
End Function


' Show system status and available commands
Public Sub ShowSystemStatus()
    Dim msg As String
    msg = "=== VBA App Builder Status ===" & vbCrLf & vbCrLf
    msg = msg & "Version: 2.2.0 (Direct App Folder Selection)" & vbCrLf
    msg = msg & "Auto-Save: " & IIf(GetAutoSaveEnabled(), "Enabled", "Disabled") & vbCrLf & vbCrLf
    
    msg = msg & "Commands:" & vbCrLf
    msg = msg & "• Build() - Browse for app folder and build it" & vbCrLf
    msg = msg & "• BuildApplication() - Same as Build()" & vbCrLf
    msg = msg & "• BuildApplication(""C:\Path\To\App"") - Build specific app folder" & vbCrLf
    msg = msg & "• ConfigureAutoSave() - Toggle auto-save preference" & vbCrLf
    msg = msg & "• ShowSystemStatus() - Display this information" & vbCrLf & vbCrLf
    
    msg = msg & "💡 Tip: Select the folder containing manifest.json (e.g., ExampleApp folder)."
    
    MsgBox msg, vbInformation, "VBA App Builder Status"
End Sub