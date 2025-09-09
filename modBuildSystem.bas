Attribute VB_Name = "modBuildSystem"
' =====================================================================================
' VBA APPLICATION BUILDER - COMPLETE BUILD SYSTEM
' =====================================================================================
' Version: 1.0.7 - Streamlined single-module build system with improved form sizing
' 
' This is a complete, self-contained VBA build system that transforms VBA development
' from a legacy workflow into a modern development experience with:
' • TRUE Code Injection - Direct VBA project manipulation
' • Modern Development - Code in external text files with version control
' • Automated Builds - One-command deployment
' • Professional Quality - Enterprise-grade form building
'
' ARCHITECTURE:
' This single module contains all necessary functionality including:
' • Build orchestration and user interface
' • VBA project manipulation and module imports
' • Form creation with JSON-driven design
' • Comprehensive JSON parsing
' • File I/O and manifest processing
' • Reference management
' • Error handling and diagnostics
'
' QUICK START USAGE:
' 1. Call Initialize() to setup the system
' 2. Call BuildInteractive() for guided building
' 3. Or call BuildApplication("AppName") for direct builds
'
' EXAMPLE:
'   Sub QuickStart()
'       Call modBuildSystem.Initialize()
'       Call modBuildSystem.BuildInteractive()
'   End Sub
'
' PREREQUISITES:
' \u2022 Trust Center: "Trust access to the VBA project object model" must be enabled
' \u2022 Source Files: Organized in the expected directory structure
' \u2022 Manifest Files: Valid JSON configuration for each application
'
' ARCHITECTURE BENEFITS:
' \u2022 Single-file deployment (just copy this .bas file)
' \u2022 No external dependencies or references required
' \u2022 Works across Excel, Word, and PowerPoint
' \u2022 Self-contained with all necessary JSON parsing and form building
' =====================================================================================

Option Explicit

' =====================================================================================
' MODULE-LEVEL VARIABLES
' =====================================================================================

' Core build system state
Private sourcePath As String          ' Path to source code root directory
Private currentProject As Object      ' Dictionary holding current project configuration

' Event callback system - allows external code to hook into build process
Public BeforeBuildCallback As String  ' Function name to call before build starts
Public AfterBuildCallback As String   ' Function name to call after build completes
Public FormCreatedCallback As String  ' Function name to call when each form is created

' VBA Extensibility API constants - used for component type identification
Private Const vbext_ct_MSForm = 3      ' UserForm component type
Private Const vbext_ct_StdModule = 1   ' Standard module component type  
Private Const vbext_ct_ClassModule = 2 ' Class module component type
Private Const vbext_ct_Document = 100  ' Document module component type (worksheets, etc.)

' =====================================================================================
' PUBLIC API - MAIN ENTRY POINTS
' =====================================================================================
' These are the primary functions users call to interact with the build system

' Initialize the VBA Build System
' 
' This is the primary setup function that must be called before using any other
' build system functionality. It performs the following tasks:
' • Validates Trust Center settings for VBA project access
' • Prompts for source path if not already configured
' • Saves configuration to Windows registry
' • Displays confirmation message
'
' USAGE: Call modBuildSystem.Initialize() once per session
' PREREQUISITES: Trust Center must allow VBA project object model access
Public Sub Initialize()
    On Error GoTo ErrorHandler
    
    ' Validate Trust Center settings first (using internal function)
    If Not ValidateTrustSettings() Then
        Exit Sub
    End If
    
    ' Get or set source path
    sourcePath = GetSourcePath()
    If sourcePath = "" Then
        sourcePath = PromptForSourcePath()
        If sourcePath <> "" Then
            SaveSourcePath sourcePath
        Else
            MsgBox "Build system requires a source path to function.", vbExclamation, "VBA Builder"
            Exit Sub
        End If
    End If
    
    ' Configuration loaded successfully
    
    MsgBox "VBA Builder initialized!" & vbCrLf & "Source Path: " & sourcePath, vbInformation, "VBA Builder"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error initializing build system: " & Err.Description, vbCritical, "VBA Builder Error"
End Sub


' Build a VBA Application from Source Files
'
' This is the core build function that transforms source files into a working
' VBA application. The build process includes:
' 1. Loading and validating the application manifest
' 2. Configuring VBA project references
' 3. Importing VBA modules from .vba files
' 4. Creating UserForms from JSON design + code-behind files
' 5. Executing build callbacks if configured
'
' PARAMETERS:
'   appName - Name of the application folder in the source directory
'
' EXPECTED STRUCTURE:
'   sourcePath\appName\manifest.json
'   sourcePath\appName\modules\*.vba
'   sourcePath\appName\forms\formName\design.json
'   sourcePath\appName\forms\formName\code-behind.vba
Public Sub BuildApplication(appName As String)
    On Error GoTo ErrorHandler
    
    ' Call before build callback if set
    If BeforeBuildCallback <> "" Then
        On Error Resume Next
        Application.Run BeforeBuildCallback
        On Error GoTo ErrorHandler
    End If
    
    ' Initialize if needed
    If sourcePath = "" Then
        Call Initialize
        If sourcePath = "" Then Exit Sub
    End If
    
    ' Set application path
    Dim appPath As String
    appPath = sourcePath & "\" & appName
    
    If Dir(appPath & "\manifest.json") = "" Then
        MsgBox "Application '" & appName & "' not found or missing manifest.json" & vbCrLf & vbCrLf & _
               "Expected location: " & appPath & "\manifest.json", vbExclamation, "App Not Found"
        Exit Sub
    End If
    
    MsgBox "Building application: " & appName & vbCrLf & "From: " & appPath, vbInformation, "VBA Builder"
    
    ' Load and process manifest
    Debug.Print "=== BUILD PROCESS DEBUG ==="
    Debug.Print "About to load manifest from: " & appPath & "\manifest.json"
    
    Dim manifest As Object
    Set manifest = LoadManifest(appPath & "\manifest.json")
    
    Debug.Print "LoadManifest returned, checking result..."
    If manifest Is Nothing Then
        Debug.Print "❌ LoadManifest returned Nothing"
    Else
        Debug.Print "✅ LoadManifest returned valid object"
        Debug.Print "Manifest name: " & manifest("name")
    End If
    
    If Not manifest Is Nothing Then
        ' Store current project for reference
        Set currentProject = manifest
        
        ' Process manifest with TRUE code injection
        Debug.Print "About to call ProcessBuild..."
        If ProcessBuild(manifest, appPath) Then
            ' Call after build callback if set
            If AfterBuildCallback <> "" Then
                On Error Resume Next
                Application.Run AfterBuildCallback, True
                On Error GoTo ErrorHandler
            End If
            MsgBox "✅ Build completed successfully!" & vbCrLf & vbCrLf & _
                   "Application '" & appName & "' is now ready to use.", vbInformation, "Build Complete"
        Else
            MsgBox "❌ Build failed during processing.", vbCritical, "Build Failed"
        End If
    Else
        MsgBox "Failed to load manifest from: " & appPath, vbCritical, "Manifest Error"
    End If
    Exit Sub
    
ErrorHandler:
    ' Call after build callback if set
    If AfterBuildCallback <> "" Then
        On Error Resume Next
        Application.Run AfterBuildCallback, False
        On Error Resume Next
    End If
    MsgBox "Error during build: " & Err.Description, vbCritical, "VBA Builder Error"
End Sub

' Interactive Build System with Application Selection Menu
'
' Provides a user-friendly interface for building applications. This function:
' • Automatically initializes the build system if needed
' • Scans the source directory for available applications
' • Presents a numbered menu of applications to build
' • Handles user input validation
' • Calls BuildApplication() with the selected app
'
' This is the recommended entry point for end users.
Public Sub BuildInteractive()
    Call Initialize
    
    Dim apps As Collection
    Dim i As Integer
    Dim msg As String
    Dim userChoice As String
    
    Set apps = GetAvailableApplications()
    
    If apps.Count = 0 Then
        msg = "No VBA applications found." & vbCrLf & vbCrLf & _
              "Please create application folders with manifest.json files in:" & vbCrLf & _
              sourcePath & vbCrLf & vbCrLf & _
              "Would you like to change the source path?"
        
        If MsgBox(msg, vbYesNo + vbQuestion, "No Applications Found") = vbYes Then
            Call ChangeSourcePath
            ' Try again after path change
            Set apps = GetAvailableApplications()
            If apps.Count = 0 Then Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    ' Build selection message
    msg = "Select VBA Application to Build:" & vbCrLf & vbCrLf
    
    For i = 1 To apps.Count
        Dim appInfo As Variant
        appInfo = apps(i)
        msg = msg & i & ". " & appInfo(0) & vbCrLf
    Next i
    
    msg = msg & vbCrLf & "Enter the number (1-" & apps.Count & "):"
    
    ' Get user selection
    userChoice = InputBox(msg, "Build VBA Application", "1")
    
    If userChoice = "" Then Exit Sub ' User cancelled
    
    ' Validate selection
    If IsNumeric(userChoice) Then
        Dim selection As Integer
        selection = CInt(userChoice)
        
        If selection >= 1 And selection <= apps.Count Then
            Dim selectedApp As Variant
            selectedApp = apps(selection)
            Call BuildApplication(CStr(selectedApp(0)))
        Else
            MsgBox "Invalid selection. Please enter a number between 1 and " & apps.Count, vbExclamation, "Invalid Selection"
        End If
    Else
        MsgBox "Invalid input. Please enter a number.", vbExclamation, "Invalid Input"
    End If
End Sub

' =====================================================================================
' SOURCE PATH MANAGEMENT
' =====================================================================================
' Functions for managing the source code directory path where applications are stored

Public Sub ChangeSourcePath()
    Dim newPath As String
    newPath = PromptForSourcePath()
    
    If newPath <> "" Then
        sourcePath = newPath
        SaveSourcePath newPath
        MsgBox "Source path updated to: " & newPath, vbInformation, "Path Updated"
    End If
End Sub

Public Function GetSourcePath() As String
    ' Try to get saved source path from registry
    On Error Resume Next
    GetSourcePath = GetSetting("VBABuilder", "Config", "SourcePath", "")
    On Error GoTo 0
End Function

Public Sub ListAvailableApplications()
    Call Initialize
    
    Dim apps As Collection
    Dim i As Integer
    Dim msg As String
    
    Set apps = GetAvailableApplications()
    
    msg = "Available VBA Applications:" & vbCrLf & vbCrLf
    
    If apps.Count = 0 Then
        msg = msg & "No applications found." & vbCrLf & _
              "Create subfolders with manifest.json files in:" & vbCrLf & _
              sourcePath & vbCrLf & vbCrLf & _
              "Expected structure:" & vbCrLf & _
              sourcePath & "\YourApp1\manifest.json" & vbCrLf & _
              sourcePath & "\YourApp2\manifest.json" & vbCrLf & vbCrLf & _
              "Current source path: " & sourcePath
    Else
        For i = 1 To apps.Count
            Dim appInfo As Variant
            appInfo = apps(i)
            msg = msg & "• " & appInfo(0) & vbCrLf
        Next i
        
        msg = msg & vbCrLf & "To build an app, use: BuildInteractive()"
    End If
    
    MsgBox msg, vbInformation, "Available Applications"
End Sub

' =====================================================================================
' SYSTEM INFORMATION & DIAGNOSTICS
' =====================================================================================
' Functions for displaying system status and diagnostic information

Public Function GetCurrentProjectInfo() As Object
    Set GetCurrentProjectInfo = currentProject
End Function

Public Sub ShowSystemStatus()
    Dim msg As String
    Dim currentSourcePath As String
    
    currentSourcePath = GetSourcePath()
    
    msg = "=== VBA Builder System Status ===" & vbCrLf & vbCrLf
    msg = msg & "Current Source Path: " & IIf(currentSourcePath = "", "(not set)", currentSourcePath) & vbCrLf
    
    If currentSourcePath <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        msg = msg & "Path Exists: " & IIf(fso.FolderExists(currentSourcePath), "✅ Yes", "❌ No") & vbCrLf & vbCrLf
        
        If fso.FolderExists(currentSourcePath) Then
            Dim apps As Collection
            Set apps = GetAvailableApplications()
            msg = msg & "Available Applications: " & apps.Count & vbCrLf
        End If
    Else
        msg = msg & vbCrLf & "❌ Source path not configured!" & vbCrLf
    End If
    
    msg = msg & vbCrLf & "Available Actions:" & vbCrLf
    msg = msg & "• Initialize() - Initialize build system" & vbCrLf
    msg = msg & "• ChangeSourcePath() - Pick new folder" & vbCrLf
    msg = msg & "• BuildInteractive() - Build with menu" & vbCrLf
    msg = msg & "• ListAvailableApplications() - Show apps"
    
    MsgBox msg, vbInformation, "VBA Builder System Status"
End Sub

' =====================================================================================
' PRIVATE HELPER FUNCTIONS - BUILD ORCHESTRATION
' =====================================================================================
' Internal functions that coordinate the build process

Private Sub SaveSourcePath(path As String)
    On Error Resume Next
    SaveSetting "VBABuilder", "Config", "SourcePath", path
    On Error GoTo 0
End Sub

Private Function PromptForSourcePath() As String
    Dim folderPicker As Object
    
    On Error GoTo ErrorHandler
    
    ' Use FileDialog for folder picking
    Set folderPicker = Application.FileDialog(4) ' msoFileDialogFolderPicker = 4
    
    With folderPicker
        .Title = "Select VBA Source Root Folder (containing app subfolders)"
        .AllowMultiSelect = False
        .InitialFileName = "C:\Users\" & Environ("USERNAME") & "\OneDrive"
        
        If .Show = -1 Then ' User clicked OK
            PromptForSourcePath = .SelectedItems(1)
        Else
            PromptForSourcePath = ""
        End If
    End With
    
    Exit Function
    
ErrorHandler:
    ' Fallback to input box if FileDialog fails
    MsgBox "Unable to show folder picker. Using text input instead.", vbInformation, "VBA Builder"
    PromptForSourcePath = PromptForSourcePathFallback()
End Function

Private Function PromptForSourcePathFallback() As String
    Dim fso As Object
    Dim selectedPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    selectedPath = InputBox("Enter the path to your VBA source root folder:" & vbCrLf & vbCrLf & _
                           "This should contain app subfolders with manifest.json files." & vbCrLf & vbCrLf & _
                           "Example: C:\MyProject\src", _
                           "VBA Builder - Source Path", _
                           "C:\YourProject\src\")
    
    If selectedPath <> "" Then
        If fso.FolderExists(selectedPath) Then
            PromptForSourcePathFallback = selectedPath
        Else
            MsgBox "Folder does not exist: " & selectedPath, vbExclamation, "Invalid Path"
            PromptForSourcePathFallback = ""
        End If
    End If
End Function

Private Function GetAvailableApplications() As Collection
    Dim apps As New Collection
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If sourcePath <> "" And fso.FolderExists(sourcePath) Then
        Set folder = fso.GetFolder(sourcePath)
        
        For Each subfolder In folder.SubFolders
            If Dir(subfolder.Path & "\manifest.json") <> "" Then
                apps.Add Array(subfolder.Name, subfolder.Path)
            End If
        Next subfolder
    End If
    
    Set GetAvailableApplications = apps
End Function

Private Function ProcessBuild(manifest As Object, appPath As String) As Boolean
    On Error GoTo ProcessBuildError
    
    Dim success As Boolean
    success = True
    
    Debug.Print "=== Starting Build Process ==="
    Debug.Print "Source Path: " & appPath
    Debug.Print "Form Build Mode: Export+Import with Pre-calculated Sizing"
    
    ' Check if manifest has required fields before accessing them
    If manifest.Exists("name") Then
        Debug.Print "Project Name: " & manifest("name")
    Else
        Debug.Print "❌ Manifest missing 'name' field"
    End If
    
    If manifest.Exists("version") Then
        Debug.Print "Project Version: " & manifest("version")
    Else
        Debug.Print "⚠️ Manifest missing 'version' field"
    End If
    
    ' 1. Configure references first
    Debug.Print "Step 1: Configuring References..."
    Dim dependencies As Object
    Set dependencies = GetManifestDependencies(manifest)
    
    If Not ConfigureReferences(dependencies) Then
        Debug.Print "❌ References configuration failed"
        success = False
    Else
        Debug.Print "✅ References configured successfully"
    End If
    
    ' 2. Import modules
    If success Then
        Debug.Print "Step 2: Importing Modules..."
        If Not ProcessModules(manifest, appPath) Then
            Debug.Print "❌ Module import failed"
            success = False
        Else
            Debug.Print "✅ Modules imported successfully"
        End If
    End If
    
    ' 3. Create forms
    If success Then
        Debug.Print "Step 3: Creating Forms..."
        If Not ProcessForms(manifest, appPath) Then
            Debug.Print "❌ Form creation failed"
            success = False
        Else
            Debug.Print "✅ Forms created successfully"
        End If
    End If
    
    ProcessBuild = success
    Exit Function
    
ProcessBuildError:
    Debug.Print "❌ PROCESSBUILD ERROR: " & Err.Number & " - " & Err.Description
    Debug.Print "Error occurred in ProcessBuild function"
    ProcessBuild = False
End Function

Private Function ProcessModules(manifest As Object, sourcePath As String) As Boolean
    Dim modules As Variant
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' Check if modules field exists and is not empty
    If Not manifest.Exists("modules") Or manifest("modules") = "" Then
        Debug.Print "No modules specified in manifest"
        ProcessModules = True ' Not an error if no modules
        Exit Function
    End If
    
    modules = Split(manifest("modules"), ",")
    
    For i = 0 To UBound(modules)
        Dim moduleName As String
        Dim modulePath As String
        
        moduleName = Trim(modules(i))
        If moduleName = "" Then GoTo NextModule
        
        modulePath = sourcePath & "\modules\" & moduleName & ".vba"
        
        ' Check if module file exists
        If Dir(modulePath) = "" Then
            MsgBox "Module file not found: " & modulePath, vbCritical, "Module File Not Found"
            ProcessModules = False
            Exit Function
        End If
        
        ' Import the module
        If Not ImportModuleFromFile(moduleName, modulePath) Then
            MsgBox "Failed to import module: " & moduleName & vbCrLf & _
                   "File: " & modulePath, vbCritical, "Module Import Error"
            ProcessModules = False
            Exit Function
        End If
        
        Debug.Print "Successfully imported module: " & moduleName
        
NextModule:
    Next i
    
    ProcessModules = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error in ProcessModules: " & Err.Description, vbCritical, "Module Import Error"
    ProcessModules = False
End Function

Private Function ProcessForms(manifest As Object, sourcePath As String) As Boolean
    Dim forms As Variant
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' Check if forms field exists and is not empty
    If Not manifest.Exists("forms") Or manifest("forms") = "" Then
        Debug.Print "No forms specified in manifest"
        ProcessForms = True ' Not an error if no forms
        Exit Function
    End If
    
    forms = Split(manifest("forms"), ",")
    
    For i = 0 To UBound(forms)
        Dim formName As String
        Dim designFile As String
        Dim codeFile As String
        
        formName = Trim(forms(i))
        If formName = "" Then GoTo NextForm
        
        designFile = sourcePath & "\forms\" & formName & "\design.json"
        codeFile = sourcePath & "\forms\" & formName & "\code-behind.vba"
        
        ' Check if form files exist
        If Dir(designFile) = "" Then
            MsgBox "Form design file not found: " & designFile, vbCritical, "Form Design File Not Found"
            ProcessForms = False
            Exit Function
        End If
        
        If Dir(codeFile) = "" Then
            MsgBox "Form code file not found: " & codeFile, vbCritical, "Form Code File Not Found"
            ProcessForms = False
            Exit Function
        End If
        
        ' Use Export+Import method with pre-calculated sizing
        Dim formObj As Object
        Set formObj = BuildAndImportForm(formName, designFile, codeFile, sourcePath)
        If formObj Is Nothing Then
            MsgBox "Failed to build form: " & formName, vbCritical, "Form Build Error"
            ProcessForms = False
            Exit Function
        End If
        
        ' Call form created callback if set
        If FormCreatedCallback <> "" Then
            On Error Resume Next
            Application.Run FormCreatedCallback, formName, formObj
            On Error GoTo ErrorHandler
        End If
        
        Debug.Print "Successfully built+imported form: " & formName
        
NextForm:
    Next i
    
    ProcessForms = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error in ProcessForms: " & Err.Description, vbCritical, "Form Creation Error"
    ProcessForms = False
End Function

Private Function GetManifestDependencies(manifest As Object) As Object
    On Error GoTo GetDepsError
    
    Debug.Print "=== GET MANIFEST DEPENDENCIES DEBUG ==="
    
    Dim deps As Object
    Set deps = CreateObject("Scripting.Dictionary")
    Debug.Print "Dependencies dictionary created"
    
    If Not manifest Is Nothing Then
        Debug.Print "Manifest is not Nothing, checking for dependencies..."
        If manifest.Exists("dependencies") Then
            Debug.Print "✅ Dependencies section found in manifest"
            Set deps = manifest("dependencies")
            Debug.Print "Dependencies object retrieved successfully"
        Else
            Debug.Print "⚠️ No dependencies section in manifest"
        End If
    Else
        Debug.Print "❌ Manifest is Nothing!"
    End If
    
    ' Ensure Forms reference is always present
    Debug.Print "Checking references in dependencies..."
    If Not deps.Exists("references") Then
        Debug.Print "No references found, adding default Forms reference"
        deps("references") = Array("Microsoft Forms 2.0 Object Library")
    Else
        Debug.Print "References found, checking for Forms library..."
        Dim refs As Object, hasForms As Boolean, i As Integer
        Dim refItem As Variant
        Set refs = deps("references")
        hasForms = False
        
        Debug.Print "Processing " & refs.Count & " references..."
        For i = 1 To refs.Count
            refItem = refs(i)
            Debug.Print "Reference " & i & ": " & refItem
            If InStr(refItem, "Forms") > 0 Then
                hasForms = True
                Debug.Print "✅ Forms library found"
                Exit For
            End If
        Next i
        
        If Not hasForms Then
            Debug.Print "Forms library not found, adding it..."
            refs.Add "Microsoft Forms 2.0 Object Library"
            Debug.Print "✅ Forms library added"
        End If
    End If
    
    Debug.Print "✅ Dependencies processed successfully"
    Set GetManifestDependencies = deps
    Exit Function
    
GetDepsError:
    Debug.Print "❌ GET DEPENDENCIES ERROR: " & Err.Number & " - " & Err.Description
    Debug.Print "Error occurred in GetManifestDependencies function"
    Set GetManifestDependencies = Nothing
End Function

' =====================================================================================
' VBA PROJECT MANIPULATION - MODULE OPERATIONS
' =====================================================================================
' Functions for importing and managing VBA modules

' Import a VBA Module from External File
'
' Imports a .vba file into the current VBA project, replacing any existing
' module with the same name. This function handles:
' • Removing existing modules to prevent conflicts
' • Using VBA Extensibility API for true code injection
' • Renaming imported modules to match expected names
'
' PARAMETERS:
'   moduleName - Target name for the module in VBA project
'   filePath   - Full path to the .vba source file
'
' RETURNS: True if import successful, False otherwise
Public Function ImportModuleFromFile(moduleName As String, filePath As String) As Boolean
    Dim vbProj As Object
    Dim vbComp As Object
    
    On Error GoTo ErrorHandler
    
    ' Get VBA project reference
    Set vbProj = GetHostVBProject()
    
    ' Check if module already exists and remove if needed
    If ModuleExists(moduleName) Then
        vbProj.VBComponents.Remove vbProj.VBComponents(moduleName)
    End If
    
    ' Import the module file
    Set vbComp = vbProj.VBComponents.Import(filePath)
    
    ' Rename if necessary (VBA might change the name)
    If vbComp.Name <> moduleName Then
        vbComp.Name = moduleName
    End If
    
    ImportModuleFromFile = True
    Exit Function
    
ErrorHandler:
    ImportModuleFromFile = False
    Debug.Print "Error importing module " & moduleName & ": " & Err.Description
End Function

Public Function ModuleExists(moduleName As String) As Boolean
    Dim vbProj As Object
    Dim vbComp As Object
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set vbProj = GetHostVBProject()
    
    For i = 1 To vbProj.VBComponents.Count
        Set vbComp = vbProj.VBComponents(i)
        If vbComp.Name = moduleName Then
            ModuleExists = True
            Exit Function
        End If
    Next i
    
    ModuleExists = False
    Exit Function
    
ErrorHandler:
    ModuleExists = False
End Function

' =====================================================================================
' VBA PROJECT MANIPULATION - REFERENCE MANAGEMENT  
' =====================================================================================
' Functions for configuring VBA project references (libraries)

' Configure VBA Project References from Dependencies
'
' Adds required library references to the VBA project based on the
' dependencies object from the application manifest. Automatically
' ensures Microsoft Forms 2.0 Object Library is always included.
'
' PARAMETERS:
'   dependencies - Dictionary object containing reference configurations
'                 Expected structure: {"references": ["Library Name 1", ...]}
'
' RETURNS: True if all references configured successfully, False otherwise
'
' SUPPORTED LIBRARIES:
' • Microsoft Forms 2.0 Object Library (always included)
' • Microsoft ActiveX Data Objects 6.1 Library
Public Function ConfigureReferences(dependencies As Object) As Boolean
    Dim vbProj As Object
    Dim refs As Object
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== CONFIGURE REFERENCES DEBUG ==="
    
    Set vbProj = GetHostVBProject()
    Set refs = vbProj.References
    
    ' Get required references from dependencies
    If Not dependencies Is Nothing Then
        Debug.Print "Dependencies object is not Nothing"
        If dependencies.Exists("references") Then
            Debug.Print "References section found in dependencies"
            Dim refsCollection As Object
            Set refsCollection = dependencies("references")
            
            Debug.Print "References object type: " & TypeName(refsCollection)
            
            ' Handle Collection object (from JSON parser)
            If TypeName(refsCollection) = "Collection" Then
                Debug.Print "Processing " & refsCollection.Count & " references from Collection"
                For i = 1 To refsCollection.Count
                    Dim refName As String
                    refName = refsCollection(i)
                    Debug.Print "Processing reference: '" & refName & "'"
                    
                    ' Check if reference already exists
                    If Not ReferenceExists(refName) Then
                        Debug.Print "Reference not found, attempting to add: " & refName
                        If Not AddReferenceByName(refName) Then
                            Debug.Print "❌ Failed to add reference: " & refName
                        Else
                            Debug.Print "✅ Successfully added reference: " & refName
                        End If
                    Else
                        Debug.Print "✅ Reference already exists: " & refName
                    End If
                Next i
            ' Handle Array (legacy support)
            ElseIf IsArray(refsCollection) Then
                Debug.Print "Processing references from Array"
                Dim refsArray As Variant
                refsArray = refsCollection
                For i = LBound(refsArray) To UBound(refsArray)
                    refName = refsArray(i)
                    Debug.Print "Processing reference: '" & refName & "'"
                    
                    If Not ReferenceExists(refName) Then
                        If Not AddReferenceByName(refName) Then
                            Debug.Print "❌ Failed to add reference: " & refName
                        End If
                    End If
                Next i
            Else
                Debug.Print "❌ Unknown references type: " & TypeName(refsCollection)
            End If
        Else
            Debug.Print "No references section found in dependencies"
        End If
    Else
        Debug.Print "Dependencies object is Nothing"
    End If
    
    Debug.Print "✅ ConfigureReferences completed successfully"
    ConfigureReferences = True
    Exit Function
    
ErrorHandler:
    Debug.Print "❌ ConfigureReferences ERROR: " & Err.Number & " - " & Err.Description
    ConfigureReferences = False
End Function

Private Function ReferenceExists(refName As String) As Boolean
    Dim vbProj As Object
    Dim refs As Object
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set vbProj = GetHostVBProject()
    Set refs = vbProj.References
    
    For i = 1 To refs.Count
        ' Try both the full description and the short name
        If StrComp(refs(i).Description, refName, vbTextCompare) = 0 Then
            ReferenceExists = True
            Exit Function
        End If
        If StrComp(refs(i).Name, refName, vbTextCompare) = 0 Then
            ReferenceExists = True
            Exit Function
        End If
    Next i
    
    ReferenceExists = False
    Exit Function
    
ErrorHandler:
    ReferenceExists = False
End Function

Private Function AddReferenceByName(refName As String) As Boolean
    Dim vbProj As Object
    
    On Error GoTo ErrorHandler
    
    Set vbProj = GetHostVBProject()
    
    Select Case refName
        Case "Microsoft ActiveX Data Objects 6.1 Library"
            vbProj.References.AddFromGuid "{2A75196C-D9EB-4129-B803-931327F72D5C}", 6, 1
        Case "Microsoft Forms 2.0 Object Library"
            vbProj.References.AddFromGuid "{0D452EE1-E08F-101A-852E-02608C4D0BB4}", 2, 0
        Case Else
            Debug.Print "Unknown reference: " & refName
            AddReferenceByName = False
            Exit Function
    End Select
    
    AddReferenceByName = True
    Exit Function
    
ErrorHandler:
    AddReferenceByName = False
End Function

' =====================================================================================
' BUILD VALIDATION & INTEGRITY CHECKS
' =====================================================================================
' Functions for validating build results and system integrity

Public Function ValidateBuildIntegrity() As Boolean
    ' Basic validation that required components exist
    Dim vbProj As Object
    Dim i As Integer
    Dim foundModules As Integer
    Dim foundForms As Integer
    
    On Error GoTo ErrorHandler
    
    Set vbProj = GetHostVBProject()
    foundModules = 0
    foundForms = 0
    
    ' Count modules and forms
    For i = 1 To vbProj.VBComponents.Count
        Select Case vbProj.VBComponents(i).Type
            Case vbext_ct_StdModule
                foundModules = foundModules + 1
            Case vbext_ct_MSForm
                foundForms = foundForms + 1
        End Select
    Next i
    
    ' Basic validation
    If foundModules > 0 And foundForms > 0 Then
        ValidateBuildIntegrity = True
    Else
        ValidateBuildIntegrity = False
    End If
    
    Exit Function
    
ErrorHandler:
    ValidateBuildIntegrity = False
End Function

' =====================================================================================
' TRUST CENTER VALIDATION
' =====================================================================================
' Functions for validating VBA project access permissions

' Validate VBA Project Access Trust Settings
'
' Checks if the Trust Center is configured to allow programmatic access
' to VBA project objects. This is essential for the build system to function.
' If access is denied, displays detailed instructions for enabling it.
'
' RETURNS: True if access is allowed, False if blocked
'
' REQUIRED SETTING:
' File → Options → Trust Center → Trust Center Settings → Macro Settings
' → "Trust access to the VBA project object model" must be checked
Public Function ValidateTrustSettings() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test VBA project access
    Dim testAccess As Object
    Set testAccess = GetHostVBProject()
    
    ValidateTrustSettings = True
    Exit Function
    
ErrorHandler:
    MsgBox "VBA Project access is disabled. Please enable 'Trust access to the VBA project object model' in Trust Center settings." & vbCrLf & vbCrLf & _
           "Steps:" & vbCrLf & _
           "1. File → Options → Trust Center" & vbCrLf & _
           "2. Trust Center Settings → Macro Settings" & vbCrLf & _
           "3. Check 'Trust access to the VBA project object model'" & vbCrLf & _
           "4. Restart Office application", vbCritical, "Trust Center Settings Required"
    ValidateTrustSettings = False
End Function 

' Get the Active VBA Project Object
'
' Returns a reference to the VBA project object for the current Office application.
' This function is host-agnostic and works across Excel, Word, and PowerPoint.
' It includes logic to ensure forms are created in the correct project (not templates).
'
' DETECTION LOGIC:
' 1. Identifies the host application (Excel/Word/PowerPoint)
' 2. Gets the appropriate document/workbook VBA project
' 3. Validates the project contains builder modules
' 4. Falls back to active VBA project if needed
'
' RETURNS: VBProject object for code injection
Public Function GetHostVBProject() As Object
    ' Host-agnostic VBProject accessor without compile-time references
    Dim hostType As String
    hostType = TypeName(Application)
    
    On Error Resume Next
    Select Case hostType
        Case "Excel.Application"
            Dim wb As Object
            Set wb = CallByName(Application, "ThisWorkbook", vbGet)
            If wb Is Nothing Then Set wb = CallByName(Application, "ActiveWorkbook", vbGet)
            If Not wb Is Nothing Then Set GetHostVBProject = wb.VBProject
        Case "Word.Application"
            Dim doc As Object
            ' ThisDocument is a project-level object; try ActiveDocument if unavailable
            Set doc = CallByName(Application, "ThisDocument", vbGet)
            If doc Is Nothing Then Set doc = CallByName(Application, "ActiveDocument", vbGet)
            If Not doc Is Nothing Then Set GetHostVBProject = doc.VBProject
        Case "PowerPoint.Application"
            Dim pres As Object
            Set pres = CallByName(Application, "ActivePresentation", vbGet)
            If Not pres Is Nothing Then Set GetHostVBProject = pres.VBProject
    End Select
    
    ' Fallback: active VBProject in VBE
    If GetHostVBProject Is Nothing Then
        Set GetHostVBProject = Application.VBE.ActiveVBProject
    End If
    
    ' Final selection: prefer the VBProject that contains the builder modules
    ' This ensures new forms are created in the same project as the builder code,
    ' avoiding creation under Normal.dotm/global template.
    Dim vbProj As Object
    Set vbProj = GetHostVBProject
    If Not ProjectContainsBuilderModules(vbProj) Then
        Dim i As Integer
        For i = 1 To Application.VBE.VBProjects.Count
            Dim candidate As Object
            Set candidate = Application.VBE.VBProjects(i)
            If ProjectContainsBuilderModules(candidate) Then
                Set GetHostVBProject = candidate
                Exit For
            End If
        Next i
    End If
    On Error GoTo 0
End Function

Private Function ProjectContainsBuilderModules(vbProj As Object) As Boolean
    Dim found As Boolean
    Dim i As Integer
    On Error Resume Next
    found = False
    If vbProj Is Nothing Then
        ProjectContainsBuilderModules = False
        Exit Function
    End If
    For i = 1 To vbProj.VBComponents.Count
        Select Case vbProj.VBComponents(i).Name
            Case "modVBABuilder", "modFormBuilder", "modVBAProject", "modJsonParser"
                found = True
                Exit For
        End Select
    Next i
    ProjectContainsBuilderModules = found
    On Error GoTo 0
End Function

' =====================================================================================
' VBA PROJECT EXPORT UTILITIES
' =====================================================================================
' Functions for exporting VBA projects (simplified implementation)

' =====================================================================================
' FILE I/O UTILITIES
' =====================================================================================
' Low-level file reading and writing functions

Public Function ReadTextFile(filePath As String) As String
    ' Read entire text file content into a string
    ' 
    ' PARAMETERS:
    '   filePath - Full path to the text file to read
    '
    ' RETURNS: File content as string, empty string on error
    Dim fileNum As Integer
    Dim fileContent As String
    
    On Error GoTo ErrorHandler
    
    fileNum = FreeFile
    Open filePath For Input As fileNum
    fileContent = Input(LOF(fileNum), fileNum)
    Close fileNum
    
    ReadTextFile = fileContent
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close fileNum
    ReadTextFile = ""
End Function

' =====================================================================================
' MANIFEST & CONFIGURATION LOADING
' =====================================================================================
' Functions for loading and parsing application manifest files

' Load and Parse Application Manifest File
'
' Reads a manifest.json file and parses it into a Dictionary object.
' The manifest defines the application structure, dependencies, and metadata.
'
' PARAMETERS:
'   manifestPath - Full path to the manifest.json file
'
' RETURNS: Dictionary object with parsed manifest data, Nothing on error
'
' EXPECTED MANIFEST STRUCTURE:
' {
'   "name": "ApplicationName",
'   "version": "1.0.0",
'   "modules": "module1,module2",
'   "forms": "form1,form2",
'   "dependencies": {
'     "references": ["Microsoft Forms 2.0 Object Library"]
'   }
' }
Public Function LoadManifest(manifestPath As String) As Object
    Dim fileContent As String
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Loading Manifest ==="
    Debug.Print "Manifest Path: " & manifestPath
    
    ' Check if manifest exists
    If Dir(manifestPath) = "" Then
        Debug.Print "❌ Manifest file not found: " & manifestPath
        MsgBox "Manifest file not found: " & manifestPath, vbCritical, "Manifest Error"
        Set LoadManifest = Nothing
        Exit Function
    End If
    
    Debug.Print "✅ Manifest file found"
    
    ' Read manifest file
    fileContent = ReadTextFile(manifestPath)
    If fileContent = "" Then
        Debug.Print "❌ Failed to read manifest file"
        MsgBox "Failed to read manifest file: " & manifestPath, vbCritical, "Manifest Error"
        Set LoadManifest = Nothing
        Exit Function
    End If
    
    Debug.Print "Manifest file size: " & Len(fileContent) & " characters"
    
    ' Parse JSON
    Debug.Print "=== MANIFEST PARSING DEBUG ==="
    Debug.Print "About to call ParseSimpleJSON with content:"
    Debug.Print Left(fileContent, 500)
    
    Dim parsedManifest As Object
    Set parsedManifest = ParseSimpleJSON(fileContent)
    
    Debug.Print "ParseSimpleJSON returned, checking result..."
    If parsedManifest Is Nothing Then
        Debug.Print "❌ ParseSimpleJSON returned Nothing"
    Else
        Debug.Print "✅ ParseSimpleJSON returned valid object"
    End If
    
    ' Validate that we have the required fields
    If parsedManifest Is Nothing Then
        Debug.Print "❌ Failed to parse manifest file"
        MsgBox "Failed to parse manifest file: " & manifestPath, vbCritical, "Manifest Parse Error"
        Set LoadManifest = Nothing
        Exit Function
    End If
    
    Debug.Print "✅ Manifest parsed successfully"
    Debug.Print "Parsed fields: " & parsedManifest.Count & " total"
    
    ' Check for required fields
    If Not parsedManifest.Exists("name") Then
        Debug.Print "❌ Manifest missing required 'name' field"
        MsgBox "Manifest missing required 'name' field: " & manifestPath, vbCritical, "Manifest Validation Error"
        Set LoadManifest = Nothing
        Exit Function
    End If
    
    Debug.Print "✅ Manifest validation passed"
    Debug.Print "Project Name: " & parsedManifest("name")
    If parsedManifest.Exists("version") Then Debug.Print "Project Version: " & parsedManifest("version")
    
    Set LoadManifest = parsedManifest
    Exit Function
    
ErrorHandler:
    Debug.Print "❌ Error in LoadManifest: " & Err.Description
    Set LoadManifest = Nothing
End Function

' =====================================================================================
' PROJECT EXPORT FUNCTIONALITY
' =====================================================================================
' Simplified VBA project export capabilities

Public Function ExportVBAProject(exportPath As String) As Boolean
    ' Simplified VBA project export (no rollback system)
    Dim vbProj As Object
    Dim vbComp As Object
    Dim i As Integer
    Dim fso As Object
    
    On Error GoTo ErrorHandler
    
    Set vbProj = GetHostVBProject()
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If vbProj Is Nothing Then
        ExportVBAProject = False
        Exit Function
    End If
    
    ' Export each component
    For i = 1 To vbProj.VBComponents.Count
        Set vbComp = vbProj.VBComponents(i)
        
        ' Skip document modules (they can't be exported)
        If vbComp.Type <> vbext_ct_Document Then
            Dim exportFile As String
            Select Case vbComp.Type
                Case vbext_ct_StdModule
                    exportFile = exportPath & "\" & vbComp.Name & ".bas"
                Case vbext_ct_ClassModule
                    exportFile = exportPath & "\" & vbComp.Name & ".cls"
                Case vbext_ct_MSForm
                    exportFile = exportPath & "\" & vbComp.Name & ".frm"
            End Select
            
            ' Export the component
            vbComp.Export exportFile
        End If
    Next i
    
    ExportVBAProject = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error exporting VBA project: " & Err.Description
    ExportVBAProject = False
End Function

' =====================================================================================
' FORM CREATION & BUILDING - CORE ENGINE
' =====================================================================================
' Primary form creation functions using export+import methodology

' Creates a UserForm with design and code-behind injection
' CRITICAL SEQUENCE: 1) Create component, 2) Rename BEFORE accessing Designer, 3) Apply design
' The Designer property locks the component and prevents renaming after access
Public Function CreateFormFromDesign(formName As String, designFile As String, codeFile As String) As Object
    Dim vbProj As Object
    Dim formComp As Object
    Dim formObj As Object
    Dim designData As Object
    Dim defaultName As String
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Form Creation (Direct Add) ==="
    Debug.Print "Target Form Name: " & formName
    Debug.Print "Design File: " & designFile
    Debug.Print "Code File: " & codeFile
    
    Set vbProj = GetHostVBProject()
    
    ' If a form with this name already exists, update it in place
    On Error Resume Next
    Set formComp = vbProj.VBComponents(formName)
    On Error GoTo ErrorHandler
    If Not formComp Is Nothing Then
        Debug.Print "Form already exists, updating in place: " & formName
        
        Set formObj = formComp.Designer
        If formObj Is Nothing Then
            Debug.Print "ERROR: Could not get designer for existing form"
            GoTo ErrorHandler
        End If
        
        ' PRE-SIZE THE EXISTING FORM before applying design
        If Dir(designFile) <> "" Then
            Debug.Print "Pre-sizing existing form based on design file..."
            Dim existingFormWidth As Long, existingFormHeight As Long
            Call CalculateFormSizeFromDesign(designFile, existingFormWidth, existingFormHeight)
            formObj.Width = existingFormWidth
            formObj.Height = existingFormHeight
            Debug.Print "Existing form pre-sized to: " & existingFormWidth & " x " & existingFormHeight & " twips (" & (existingFormWidth / 20) & " x " & (existingFormHeight / 20) & " points)"
        End If
        
        ' Apply design and sizing corrections
        If Dir(designFile) <> "" Then
            Debug.Print "Loading design from: " & designFile
            Set designData = ParseFormDesign(designFile)
            If Not designData Is Nothing Then
                Call ApplyFormDesign(formObj, designData)
            End If
        Else
            Debug.Print "No design file found, using defaults"
            Call ApplyDefaultDesign(formObj, formName)
        End If
        
        ' Import code-behind
        If Dir(codeFile) <> "" Then
            Debug.Print "Importing code-behind from: " & codeFile
            Call ImportCodeBehind(formComp, codeFile)
        Else
            Debug.Print "No code-behind file found"
        End If
        
        Debug.Print "✅ Form update complete!"
        Debug.Print "   VBA Name: " & formComp.Name
        Debug.Print "   Target Name: " & formName
        Debug.Print "   Final size: " & formObj.Width & " x " & formObj.Height & " (twips)"
        Debug.Print "   Final size in points: " & (formObj.Width / 20) & " x " & (formObj.Height / 20)
        Set CreateFormFromDesign = formObj
        Exit Function
    End If
    
    ' Add a new form component
    Set formComp = vbProj.VBComponents.Add(vbext_ct_MSForm)
    If formComp Is Nothing Then
        Debug.Print "ERROR: Failed to add UserForm component"
        GoTo ErrorHandler
    End If
    
    ' Capture the default name and attempt to rename BEFORE accessing Designer
    defaultName = formComp.Name
    If StrComp(defaultName, formName, vbTextCompare) <> 0 Then
        If Not SafeRenameVBComponent(vbProj, formComp, formName) Then
            Debug.Print "⚠️ Rename via VBComponents failed. Trying Designer.Name..."
            On Error Resume Next
            Dim tmpDesigner As Object
            Set tmpDesigner = formComp.Designer
            If Not tmpDesigner Is Nothing Then
                tmpDesigner.Name = formName
            End If
            If Err.Number <> 0 Then
                Debug.Print "❌ Designer.Name rename failed: " & Err.Description
                Err.Clear
                ' Final fallback: import a temp .frm with the desired name
                Dim tempFrm As String
                tempFrm = CreateTempFormFile(formName, designFile, codeFile)
                If tempFrm <> "" Then
                    Debug.Print "Attempting import-based creation to guarantee name..."
                    ' Remove the just-created component and import
                    On Error Resume Next
                    vbProj.VBComponents.Remove formComp
                    Set formComp = Nothing
                    Dim imported As Object
                    Set imported = vbProj.VBComponents.Import(tempFrm)
                    If Not imported Is Nothing Then
                        If StrComp(imported.Name, formName, vbTextCompare) <> 0 Then
                            On Error Resume Next
                            imported.Name = formName
                            On Error GoTo ErrorHandler
                        End If
                        Set formComp = imported
                    Else
                        Debug.Print "❌ Import-based creation failed"
                        GoTo ErrorHandler
                    End If
                    On Error GoTo ErrorHandler
                Else
                    Debug.Print "❌ Could not create temp .frm for import"
                    GoTo ErrorHandler
                End If
            End If
            On Error GoTo ErrorHandler
        End If
    End If
    
    ' Refresh formComp reference by explicit lookup
    On Error Resume Next
    Set formComp = vbProj.VBComponents(formName)
    On Error GoTo ErrorHandler
    If formComp Is Nothing Then
        ' As a last resort, try defaultName (in case of unusual rename behavior)
        On Error Resume Next
        Set formComp = vbProj.VBComponents(defaultName)
        On Error GoTo ErrorHandler
    End If
    
    ' Get the form designer object AFTER final name is set
    Set formObj = formComp.Designer
    If formObj Is Nothing Then
        Debug.Print "ERROR: Could not get form designer"
        GoTo ErrorHandler
    End If
    
    ' PRE-SIZE THE FORM before applying design - this is critical!
    If Dir(designFile) <> "" Then
        Debug.Print "Pre-sizing form based on design file..."
        Dim formWidth As Long, formHeight As Long
        Call CalculateFormSizeFromDesign(designFile, formWidth, formHeight)
        formObj.Width = formWidth
        formObj.Height = formHeight
        Debug.Print "Form pre-sized to: " & formWidth & " x " & formHeight & " twips (" & (formWidth / 20) & " x " & (formHeight / 20) & " points)"
    End If
    
    ' Apply design if available
    If Dir(designFile) <> "" Then
        Debug.Print "Loading design from: " & designFile
        Set designData = ParseFormDesign(designFile)
        If Not designData Is Nothing Then
            Call ApplyFormDesign(formObj, designData)
        End If
    Else
        Debug.Print "No design file found, using defaults"
        Call ApplyDefaultDesign(formObj, formName)
    End If
    
    ' Import code-behind if available
    If Dir(codeFile) <> "" Then
        Debug.Print "Importing code-behind from: " & codeFile
        Call ImportCodeBehind(formComp, codeFile)
    Else
        Debug.Print "No code-behind file found"
    End If
    
    Debug.Print "✅ Form creation complete!"
    Debug.Print "   VBA Name: " & formComp.Name
    Debug.Print "   Target Name: " & formName
    Debug.Print "   Direct Add/Import method used"
    
    Set CreateFormFromDesign = formObj
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR in CreateFormFromDesign: " & Err.Description
    Set CreateFormFromDesign = Nothing
End Function

' High-level builder: create temp form → export → normalize → import as target
' Build and Import a Complete UserForm
'
' This is the primary form building function that creates a UserForm using
' the export+import methodology for maximum compatibility. The process:
' 1. Creates a temporary form with design and code-behind
' 2. Exports it to a .frm file using VBA Editor
' 3. Normalizes the .frm metadata and sizing
' 4. Imports the .frm back as the final form
'
' PARAMETERS:
'   formName   - Target name for the form (e.g., "frmMainApp")
'   designFile - Path to design.json file with form layout
'   codeFile   - Path to code-behind.vba file with form code
'   appPath    - Application root path for export location
'
' RETURNS: Form Designer object if successful, Nothing on error
Public Function BuildAndImportForm(formName As String, designFile As String, codeFile As String, appPath As String) As Object
    Dim exportPath As String
    On Error GoTo ErrorHandler
    
    exportPath = ExportFormAsFile(formName, designFile, codeFile, appPath)
    If exportPath = "" Then
        Set BuildAndImportForm = Nothing
        Exit Function
    End If
    
    Set BuildAndImportForm = ImportFormFile(exportPath, formName)
    Exit Function
    
ErrorHandler:
    Debug.Print "BuildAndImportForm ERROR: " & Err.Description
    Set BuildAndImportForm = Nothing
End Function

' =====================================================================================
' FORM IMPORT & INTEGRATION
' =====================================================================================
' Functions for importing pre-built .frm files into VBA projects

Public Function ImportFormFile(frmFile As String, targetName As String) As Object
    Dim vbProj As Object
    Dim imported As Object
    On Error GoTo ErrorHandler
    
    If Dir(frmFile) = "" Then
        Debug.Print "ImportFormFile: file not found - " & frmFile
        Set ImportFormFile = Nothing
        Exit Function
    End If
    
    Set vbProj = GetHostVBProject()
    If vbProj Is Nothing Then GoTo ErrorHandler
    
    ' Remove any existing form with the target name first
    Call RemoveExistingForm(vbProj, targetName)
    
    ' Import the .frm
    Set imported = vbProj.VBComponents.Import(frmFile)
    If imported Is Nothing Then
        Debug.Print "ImportFormFile: import returned Nothing"
        Set ImportFormFile = Nothing
        Exit Function
    End If
    
    ' Rename immediately before accessing Designer
    Call SafeRenameVBComponent(vbProj, imported, targetName)
    
    ' Access Designer and apply sizing corrections
    Dim formObj As Object
    Set formObj = imported.Designer
    If Not formObj Is Nothing Then
        ' Apply sizing corrections after import
        Call ApplyFormSizingAfterImport(formObj, targetName)
        
        ' Force additional resize to ensure proper sizing
        Call ResizeFormToFitControls(formObj)
    End If
    
    Set ImportFormFile = formObj
    Debug.Print "ImportFormFile: successfully imported as " & imported.Name
    Exit Function
    
ErrorHandler:
    Debug.Print "ImportFormFile ERROR: " & Err.Number & " - " & Err.Description
    Set ImportFormFile = Nothing
End Function

' =====================================================================================
' FORM DESIGN PARSING
' =====================================================================================
' Functions for parsing JSON form design specifications

Private Function ParseFormDesign(designFile As String) As Object
    Dim fileContent As String
    Dim designData As Object
    
    On Error GoTo ErrorHandler
    
    ' Read the design file
    fileContent = ReadTextFile(designFile)
    If fileContent = "" Then
        Set ParseFormDesign = Nothing
        Exit Function
    End If
    
    ' Parse JSON
    Set designData = ParseSimpleJSON(fileContent)
    Set ParseFormDesign = designData
    Exit Function
    
ErrorHandler:
    Debug.Print "Error parsing design file: " & Err.Description
    Set ParseFormDesign = Nothing
End Function

' =====================================================================================
' FORM DESIGN APPLICATION
' =====================================================================================
' Functions for applying design specifications to form objects

Private Sub ApplyFormDesign(formObj As Object, designData As Object)
    On Error Resume Next
    
    Debug.Print "Applying form design..."
    
    ' Apply basic form properties
    If designData.Exists("caption") Then
        formObj.Caption = designData("caption")
    End If
    
    ' Clear existing controls before re-creating
    Call ClearFormControls(formObj)
    
    ' Create controls if they exist
    If designData.Exists("controls") Then
        Call CreateControls(formObj, designData("controls"))
    End If
    
    ' Try to apply form sizing only if the form object seems stable
    On Error Resume Next
    Dim currentWidthPts As Single, currentHeightPts As Single
    Dim canAccessSize As Boolean
    
    ' Test if we can safely access the form size properties
    currentWidthPts = formObj.Width / 20
    currentHeightPts = formObj.Height / 20
    canAccessSize = (Err.Number = 0)
    
    If canAccessSize Then
        Debug.Print "ApplyFormDesign - Current form size: " & currentWidthPts & " x " & currentHeightPts & " points"
        
        If currentWidthPts >= 300 And currentHeightPts >= 200 Then
            Debug.Print "Form already has reasonable size: " & currentWidthPts & " x " & currentHeightPts & " points - skipping resize"
        Else
            Debug.Print "Form has small size: " & currentWidthPts & " x " & currentHeightPts & " points - attempting calculated size"
            
            ' Try to apply size, but don't fail if it doesn't work
            Err.Clear
            Call CalculateAndApplyFormSize(formObj, designData)
            
            If Err.Number = 0 Then
                ' Check size after applying
                currentWidthPts = formObj.Width / 20
                currentHeightPts = formObj.Height / 20
                Debug.Print "Form size after CalculateAndApplyFormSize: " & currentWidthPts & " x " & currentHeightPts & " points"
            Else
                Debug.Print "Warning: Could not apply form sizing: " & Err.Description
                Err.Clear
            End If
        End If
    Else
        Debug.Print "Warning: Cannot access form size properties - skipping form sizing (Error: " & Err.Description & ")"
        Err.Clear
    End If
    
    On Error GoTo 0
    
    ' Validate form size and controls fit properly (using calculated dimensions)
    Call ValidateFormLayoutWithCalculatedSize(designData)
    
    ' Also add a simple validation against the actual form object
    Debug.Print "Form object validation - Width: " & formObj.Width & " twips (" & (formObj.Width / 20) & " points)"
    Debug.Print "Form object validation - Height: " & formObj.Height & " twips (" & (formObj.Height / 20) & " points)"
    
    If designData.Exists("startUpPosition") Then
        formObj.StartUpPosition = designData("startUpPosition")
    End If
    
    Debug.Print "Form design applied"
    On Error GoTo 0
End Sub

Private Sub ApplyDefaultDesign(formObj As Object, formName As String)
    On Error Resume Next
    
    formObj.Caption = formName
    ' Use reasonable default dimensions (convert points to twips)
    formObj.Width = 450 * 20   ' 450 points = 9000 twips
    formObj.Height = 300 * 20  ' 300 points = 6000 twips
    
    ' Add a default button
    Dim btn As Object
    Set btn = formObj.Controls.Add("Forms.CommandButton.1", "btnDefault")
    btn.Caption = "Click Me"
    btn.Left = 100
    btn.Top = 80
    btn.Width = 100
    btn.Height = 30
    
    Debug.Print "Applied default design"
    On Error GoTo 0
End Sub

' =====================================================================================
' FORM CONTROL CREATION
' =====================================================================================
' Functions for creating and configuring form controls

Private Sub CreateControls(formObj As Object, controlsCollection As Object)
    Dim i As Integer
    Dim controlData As Object
    Dim ctrl As Object
    
    On Error Resume Next
    
    Debug.Print "Creating " & controlsCollection.Count & " controls..."
    
    For i = 1 To controlsCollection.Count
        Set controlData = controlsCollection(i)
        
        If Not controlData Is Nothing Then
            
            Dim controlType As String
            controlType = "Forms.CommandButton.1" ' Default
            
            If controlData.Exists("type") Then
                Select Case controlData("type")
                    Case "CommandButton"
                        controlType = "Forms.CommandButton.1"
                    Case "Label"
                        controlType = "Forms.Label.1"
                    Case "TextBox"
                        controlType = "Forms.TextBox.1"
                    Case "ListBox"
                        controlType = "Forms.ListBox.1"
                    Case "ComboBox"
                        controlType = "Forms.ComboBox.1"
                    Case "CheckBox"
                        controlType = "Forms.CheckBox.1"
                    Case "OptionButton"
                        controlType = "Forms.OptionButton.1"
                End Select
            End If
            
            ' Create control
            Dim controlName As String
            controlName = "Control" & i
            If controlData.Exists("name") Then
                controlName = controlData("name")
            End If
            
            Set ctrl = formObj.Controls.Add(controlType, controlName)
            
            ' Apply properties
            If controlData.Exists("caption") Then ctrl.Caption = controlData("caption")
            If controlData.Exists("left") Then 
                ctrl.Left = controlData("left")
                Debug.Print "Control " & controlName & " left set to: " & ctrl.Left
            End If
            If controlData.Exists("top") Then 
                ctrl.Top = controlData("top")
                Debug.Print "Control " & controlName & " top set to: " & ctrl.Top
            End If
            If controlData.Exists("width") Then 
                ctrl.Width = controlData("width")
                Debug.Print "Control " & controlName & " width set to: " & ctrl.Width
            End If
            If controlData.Exists("height") Then 
                ctrl.Height = controlData("height")
                Debug.Print "Control " & controlName & " height set to: " & ctrl.Height
            End If
            If controlData.Exists("text") Then On Error Resume Next: ctrl.Text = controlData("text"): On Error GoTo 0
            ' ListBox/ComboBox specific
            If TypeName(ctrl) = "ListBox" Then
                If controlData.Exists("multiSelect") Then
                    Dim ms As String
                    ms = LCase$(CStr(controlData("multiSelect")))
                    Select Case ms
                        Case "fmmultiselectsingle", "0": ctrl.MultiSelect = 0
                        Case "fmmultiselectmulti", "1": ctrl.MultiSelect = 1
                        Case "fmmultiselectextended", "2": ctrl.MultiSelect = 2
                    End Select
                End If
            End If
            
            ' Apply font properties if they exist
            If controlData.Exists("font") Then
                Dim fontData As Object
                Set fontData = controlData("font")
                If fontData.Exists("name") Then ctrl.Font.Name = fontData("name")
                If fontData.Exists("size") Then ctrl.Font.Size = fontData("size")
                If fontData.Exists("bold") Then ctrl.Font.Bold = fontData("bold")
            End If
            
            Debug.Print "Created control: " & controlName
        End If
    Next i
    
    Debug.Print "Controls created successfully"
    On Error GoTo 0
End Sub

' =====================================================================================
' CODE-BEHIND INTEGRATION
' =====================================================================================
' Functions for importing VBA code into form modules

Private Sub ImportCodeBehind(formComp As Object, codeFile As String)
    Dim fileContent As String
    Dim codeModule As Object
    
    On Error GoTo ErrorHandler
    
    ' Read code-behind file
    fileContent = ReadTextFile(codeFile)
    If fileContent = "" Then
        Debug.Print "Code-behind file is empty"
        Exit Sub
    End If
    
    ' Get form's code module
    Set codeModule = formComp.CodeModule
    
    ' Clear existing code and insert new
    If codeModule.CountOfLines > 0 Then
        codeModule.DeleteLines 1, codeModule.CountOfLines
    End If
    
    codeModule.InsertLines 1, fileContent
    
    Debug.Print "Code-behind imported successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error importing code-behind: " & Err.Description
End Sub

' =====================================================================================
' FORM IMPORT HELPER FUNCTIONS
' =====================================================================================
' Utility functions supporting form import operations

Private Sub RemoveExistingForm(vbProj As Object, formName As String)
    ' Remove any existing form with the target name
    On Error Resume Next
    
    Dim i As Integer
    For i = vbProj.VBComponents.Count To 1 Step -1
        If vbProj.VBComponents(i).Name = formName And vbProj.VBComponents(i).Type = vbext_ct_MSForm Then
            Debug.Print "Removing existing form: " & formName
            vbProj.VBComponents.Remove vbProj.VBComponents(i)
            Exit For
        End If
    Next i
    
    On Error GoTo 0
End Sub

Private Function SafeRenameVBComponent(vbProj As Object, vbComp As Object, newName As String) As Boolean
    Dim oldName As String
    Dim i As Integer
    
    On Error Resume Next
    SafeRenameVBComponent = False
    If vbComp Is Nothing Then Exit Function
    
    oldName = vbComp.Name
    If StrComp(oldName, newName, vbTextCompare) = 0 Then
        SafeRenameVBComponent = True
        Exit Function
    End If
    
    ' If a component with the target name already exists and is a form, remove it
    For i = vbProj.VBComponents.Count To 1 Step -1
        If vbProj.VBComponents(i).Name = newName And vbProj.VBComponents(i).Type = vbext_ct_MSForm Then
            vbProj.VBComponents.Remove vbProj.VBComponents(i)
            Exit For
        End If
    Next i
    
    ' Attempt rename via collection indexer (more reliable than setting vbComp.Name directly)
    vbProj.VBComponents(oldName).Name = newName
    If Err.Number = 0 Then
        SafeRenameVBComponent = True
        On Error GoTo 0
        Exit Function
    End If
    
    Err.Clear
    
    ' Fallback: set Name property directly
    vbComp.Name = newName
    If Err.Number = 0 Then
        SafeRenameVBComponent = True
    Else
        Err.Clear
        SafeRenameVBComponent = False
    End If
    On Error GoTo 0
End Function

Private Sub ClearFormControls(formObj As Object)
    On Error Resume Next
    Dim idx As Long
    For idx = formObj.Controls.Count - 1 To 0 Step -1
        formObj.Controls.Remove formObj.Controls(idx).Name
    Next idx
    On Error GoTo 0
End Sub

Private Function CreateTempFormFile(formName As String, designFile As String, codeFile As String) As String
    ' Create a minimal .frm file that can be imported with the correct name and size
    Dim tempPath As String
    Dim tempFormFile As String
    Dim fileContent As String
    Dim formWidth As Long
    Dim formHeight As Long
    
    On Error GoTo ErrorHandler
    
    ' Create temp file path
    tempPath = Environ("TEMP")
    tempFormFile = tempPath & "\" & formName & ".frm"
    
    Debug.Print "Creating minimal form file: " & tempFormFile
    
    ' Calculate proper form size from design file BEFORE creating the form
    Call CalculateFormSizeFromDesign(designFile, formWidth, formHeight)
    
    Debug.Print "Initial form size: " & formWidth & " x " & formHeight & " twips"
    Debug.Print "Initial form size: " & (formWidth / 20) & " x " & (formHeight / 20) & " points"
    
    ' Create minimal .frm file content with calculated size
    fileContent = "VERSION 5.00" & vbCrLf
    fileContent = fileContent & "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " & formName & " " & vbCrLf
    fileContent = fileContent & "   Caption         =   """ & formName & """" & vbCrLf
    fileContent = fileContent & "   ClientHeight    =   " & formHeight & vbCrLf
    fileContent = fileContent & "   ClientWidth     =   " & formWidth & vbCrLf
    fileContent = fileContent & "   ClientLeft      =   45" & vbCrLf
    fileContent = fileContent & "   ClientTop       =   390" & vbCrLf
    fileContent = fileContent & "   StartUpPosition =   1  'CenterOwner" & vbCrLf
    fileContent = fileContent & "End" & vbCrLf
    fileContent = fileContent & "Attribute VB_Name = """ & formName & """" & vbCrLf
    fileContent = fileContent & "Attribute VB_GlobalNameSpace = False" & vbCrLf
    fileContent = fileContent & "Attribute VB_Creatable = False" & vbCrLf
    fileContent = fileContent & "Attribute VB_PredeclaredId = True" & vbCrLf
    fileContent = fileContent & "Attribute VB_Exposed = False" & vbCrLf
    
    ' Add minimal code-behind (we'll apply the real code after import)
    fileContent = fileContent & "Option Explicit" & vbCrLf
    
    ' Write the .frm file
    Dim fileNum As Integer
    fileNum = FreeFile
    Open tempFormFile For Output As fileNum
    Print #fileNum, fileContent
    Close fileNum
    
    Debug.Print "Minimal form file created successfully"
    CreateTempFormFile = tempFormFile
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close fileNum
    Debug.Print "Error creating temp form file: " & Err.Description
    CreateTempFormFile = ""
End Function

Private Sub CalculateFormSizeFromDesign(designFile As String, ByRef formWidth As Long, ByRef formHeight As Long)
    ' Calculate proper form dimensions from design file before creating the form
    On Error Resume Next
    
    Dim designData As Object
    Dim requiredWidth As Single
    Dim requiredHeight As Single
    Dim padding As Single
    Dim minWidth As Single
    Dim minHeight As Single
    
    ' Set defaults first
    padding = 40
    minWidth = 450
    minHeight = 300
    formWidth = CLng(minWidth * 20)  ' Convert to twips
    formHeight = CLng(minHeight * 20)
    
    ' Try to load and parse the design file
    If Dir(designFile) <> "" Then
        Set designData = ParseFormDesign(designFile)
        
        If Not designData Is Nothing Then
            ' Check for explicit dimensions first
            If designData.Exists("width") And designData.Exists("height") Then
                formWidth = CLng(CSng(designData("width")) * 20)
                formHeight = CLng(CSng(designData("height")) * 20)
                Debug.Print "Using explicit form dimensions from design"
            ElseIf designData.Exists("controls") Then
                ' Calculate from controls
                Call CalculateControlExtentsFromDesign(designData("controls"), requiredWidth, requiredHeight)
                
                ' Add padding
                requiredWidth = requiredWidth + (padding * 2)
                requiredHeight = requiredHeight + (padding * 2)
                
                ' Ensure minimum dimensions
                If requiredWidth < minWidth Then requiredWidth = minWidth
                If requiredHeight < minHeight Then requiredHeight = minHeight
                
                ' Convert to twips
                formWidth = CLng(requiredWidth * 20)
                formHeight = CLng(requiredHeight * 20)
                
                Debug.Print "Calculated form size from controls: " & requiredWidth & " x " & requiredHeight & " points"
            End If
        End If
    End If
    
    Debug.Print "Final initial form size: " & formWidth & " x " & formHeight & " twips"
    On Error GoTo 0
End Sub

' =====================================================================================
' FORM SIZING & LAYOUT CALCULATIONS
' =====================================================================================
' Functions for calculating optimal form dimensions based on controls

Private Sub CalculateAndApplyFormSize(formObj As Object, designData As Object)
    On Error Resume Next
    
    Dim requiredWidth As Single
    Dim requiredHeight As Single
    Dim padding As Single
    Dim minWidth As Single
    Dim minHeight As Single
    
    ' Set padding (minimum space around controls) - increased for better spacing
    padding = 40  ' Increased to 40 for more generous spacing
    
    ' Set minimum form dimensions - increased for better appearance  
    minWidth = 450  ' Increased minimum width
    minHeight = 300  ' Increased minimum height
    
    ' Calculate required size based on controls
    If designData.Exists("controls") Then
        Call CalculateControlExtents(formObj, designData("controls"), requiredWidth, requiredHeight)
        
        Debug.Print "Control extents calculated: " & requiredWidth & " x " & requiredHeight
        
        ' Add padding to required dimensions
        requiredWidth = requiredWidth + (padding * 2)
        requiredHeight = requiredHeight + (padding * 2)
        
        ' Ensure minimum dimensions
        If requiredWidth < minWidth Then requiredWidth = minWidth
        If requiredHeight < minHeight Then requiredHeight = minHeight
        
        Debug.Print "Control extents with padding: " & requiredWidth & " x " & requiredHeight
        Debug.Print "Final form size: " & requiredWidth & " x " & requiredHeight
    Else
        ' No controls - use design.json dimensions or defaults
        If designData.Exists("width") Then
            requiredWidth = CSng(designData("width"))
        Else
            requiredWidth = minWidth
        End If
        
        If designData.Exists("height") Then
            requiredHeight = CSng(designData("height"))
        Else
            requiredHeight = minHeight
        End If
    End If
    
    ' Apply the calculated dimensions (convert points to twips)
    ' VBA forms use twips internally (20 twips = 1 point)
    formObj.Width = requiredWidth * 20
    formObj.Height = requiredHeight * 20
    
    Debug.Print "Form size applied: " & formObj.Width & " x " & formObj.Height & " (twips)"
    Debug.Print "Form size in points: " & (formObj.Width / 20) & " x " & (formObj.Height / 20)
    On Error GoTo 0
End Sub

Private Sub CalculateControlExtents(formObj As Object, controlsCollection As Object, ByRef maxWidth As Single, ByRef maxHeight As Single)
    On Error Resume Next
    
    Dim i As Integer
    Dim controlData As Object
    Dim rightEdge As Single
    Dim bottomEdge As Single
    
    maxWidth = 0
    maxHeight = 0
    
    ' Find the rightmost and bottommost edges of all controls
    For i = 1 To controlsCollection.Count
        Set controlData = controlsCollection(i)
        
        If Not controlData Is Nothing Then
            ' Calculate right edge (left + width)
            If controlData.Exists("left") And controlData.Exists("width") Then
                rightEdge = CSng(controlData("left")) + CSng(controlData("width"))
                If rightEdge > maxWidth Then maxWidth = rightEdge
            End If
            
            ' Calculate bottom edge (top + height)
            If controlData.Exists("top") And controlData.Exists("height") Then
                bottomEdge = CSng(controlData("top")) + CSng(controlData("height"))
                If bottomEdge > maxHeight Then maxHeight = bottomEdge
            End If
        End If
    Next i
    
    Debug.Print "Control extents calculated: " & maxWidth & " x " & maxHeight
    On Error GoTo 0
End Sub

Private Sub ValidateFormLayout(formObj As Object, designData As Object)
    On Error Resume Next
    
    Dim i As Integer
    Dim controlData As Object
    Dim control As Object
    Dim rightEdge As Single
    Dim bottomEdge As Single
    Dim formWidth As Single
    Dim formHeight As Single
    Dim issues As String
    
    issues = ""
    ' Convert form dimensions from twips to points for comparison
    formWidth = formObj.Width / 20
    formHeight = formObj.Height / 20
    
    ' Debug form dimensions for troubleshooting
    Debug.Print "Form validation - Width: " & formObj.Width & " twips (" & formWidth & " points)"
    Debug.Print "Form validation - Height: " & formObj.Height & " twips (" & formHeight & " points)"
    
    ' Check if all controls fit within form boundaries
    If designData.Exists("controls") Then
        For i = 1 To designData("controls").Count
            Set controlData = designData("controls")(i)
            
            If Not controlData Is Nothing Then
                ' Calculate control edges
                If controlData.Exists("left") And controlData.Exists("width") Then
                    rightEdge = CSng(controlData("left")) + CSng(controlData("width"))
                    If rightEdge > formWidth Then
                        issues = issues & "Control '" & controlData("name") & "' extends beyond form width (" & rightEdge & " > " & formWidth & ")" & vbCrLf
                    End If
                End If
                
                If controlData.Exists("top") And controlData.Exists("height") Then
                    bottomEdge = CSng(controlData("top")) + CSng(controlData("height"))
                    If bottomEdge > formHeight Then
                        issues = issues & "Control '" & controlData("name") & "' extends beyond form height (" & bottomEdge & " > " & formHeight & ")" & vbCrLf
                    End If
                End If
            End If
        Next i
    End If
    
    ' Report validation results
    If issues <> "" Then
        Debug.Print "⚠️ FORM LAYOUT VALIDATION WARNINGS:"
        Debug.Print issues
    Else
        Debug.Print "✅ Form layout validation passed - all controls fit properly"
    End If
    
    On Error GoTo 0
End Sub

Private Sub ApplyFormDesignControlsOnly(formObj As Object, designData As Object)
    ' Apply form design but skip all sizing operations to avoid object corruption
    On Error Resume Next
    
    Debug.Print "Applying form design (controls only, no sizing)..."
    
    ' Apply basic form properties
    If designData.Exists("caption") Then
        formObj.Caption = designData("caption")
    End If
    
    ' Clear existing controls before re-creating
    Call ClearFormControls(formObj)
    
    ' Create controls if they exist
    If designData.Exists("controls") Then
        Call CreateControls(formObj, designData("controls"))
    End If
    
    ' Apply startup position
    If designData.Exists("startUpPosition") Then
        formObj.StartUpPosition = designData("startUpPosition")
    End If
    
    Debug.Print "Form design applied (controls only)"
    On Error GoTo 0
End Sub

Private Sub ValidateFormLayoutWithCalculatedSize(designData As Object)
    On Error Resume Next
    
    Dim i As Integer
    Dim controlData As Object
    Dim rightEdge As Single
    Dim bottomEdge As Single
    Dim formWidth As Single
    Dim formHeight As Single
    Dim padding As Single
    Dim minWidth As Single
    Dim minHeight As Single
    Dim issues As String
    
    issues = ""
    
    ' Use the same calculation logic as CalculateAndApplyFormSize
    padding = 40
    minWidth = 450
    minHeight = 300
    
    ' Calculate what the form size should be
    If designData.Exists("controls") Then
        Dim requiredWidth As Single
        Dim requiredHeight As Single
        
        ' Calculate control extents
        For i = 1 To designData("controls").Count
            Set controlData = designData("controls")(i)
            
            If Not controlData Is Nothing Then
                ' Calculate right edge (left + width)
                If controlData.Exists("left") And controlData.Exists("width") Then
                    rightEdge = CSng(controlData("left")) + CSng(controlData("width"))
                    If rightEdge > requiredWidth Then requiredWidth = rightEdge
                End If
                
                ' Calculate bottom edge (top + height)
                If controlData.Exists("top") And controlData.Exists("height") Then
                    bottomEdge = CSng(controlData("top")) + CSng(controlData("height"))
                    If bottomEdge > requiredHeight Then requiredHeight = bottomEdge
                End If
            End If
        Next i
        
        ' Add padding
        formWidth = requiredWidth + (padding * 2)
        formHeight = requiredHeight + (padding * 2)
        
        ' Ensure minimum dimensions
        If formWidth < minWidth Then formWidth = minWidth
        If formHeight < minHeight Then formHeight = minHeight
        
        Debug.Print "Validation using calculated form size: " & formWidth & " x " & formHeight & " points"
        
        ' Now validate that all controls fit within this calculated size
        For i = 1 To designData("controls").Count
            Set controlData = designData("controls")(i)
            
            If Not controlData Is Nothing Then
                ' Calculate control edges
                If controlData.Exists("left") And controlData.Exists("width") Then
                    rightEdge = CSng(controlData("left")) + CSng(controlData("width"))
                    If rightEdge > formWidth Then
                        issues = issues & "Control '" & controlData("name") & "' extends beyond calculated form width (" & rightEdge & " > " & formWidth & ")" & vbCrLf
                    End If
                End If
                
                If controlData.Exists("top") And controlData.Exists("height") Then
                    bottomEdge = CSng(controlData("top")) + CSng(controlData("height"))
                    If bottomEdge > formHeight Then
                        issues = issues & "Control '" & controlData("name") & "' extends beyond calculated form height (" & bottomEdge & " > " & formHeight & ")" & vbCrLf
                    End If
                End If
            End If
        Next i
    End If
    
    ' Report validation results
    If issues <> "" Then
        Debug.Print "⚠️ FORM LAYOUT VALIDATION WARNINGS:"
        Debug.Print issues
    Else
        Debug.Print "✅ Form layout validation passed - all controls fit within calculated form size"
    End If
    
    On Error GoTo 0
End Sub

Private Sub ApplyFormSizingAfterImport(formObj As Object, formName As String)
    On Error Resume Next
    
    Debug.Print "=== Applying Form Sizing After Import ==="
    Debug.Print "Form: " & formName
    Debug.Print "Current size: " & formObj.Width & " x " & formObj.Height & " (twips)"
    Debug.Print "Current size in points: " & (formObj.Width / 20) & " x " & (formObj.Height / 20)
    
    ' Calculate proper size based on controls
    Dim requiredWidth As Single
    Dim requiredHeight As Single
    Dim padding As Single
    Dim minWidth As Single
    Dim minHeight As Single
    
    padding = 40  ' Increased to 40 for more generous spacing
    minWidth = 450  ' Increased minimum width
    minHeight = 300  ' Increased minimum height
    
    ' Calculate based on actual controls on the form
    Call CalculateActualControlExtents(formObj, requiredWidth, requiredHeight)
    
    ' Add padding
    requiredWidth = requiredWidth + (padding * 2)
    requiredHeight = requiredHeight + (padding * 2)
    
    ' Ensure minimum dimensions
    If requiredWidth < minWidth Then requiredWidth = minWidth
    If requiredHeight < minHeight Then requiredHeight = minHeight
    
    ' Apply the calculated dimensions (convert points to twips)
    formObj.Width = requiredWidth * 20
    formObj.Height = requiredHeight * 20
    
    Debug.Print "Applied size: " & formObj.Width & " x " & formObj.Height & " (twips)"
    Debug.Print "Applied size in points: " & (formObj.Width / 20) & " x " & (formObj.Height / 20)
    On Error GoTo 0
End Sub

Private Sub CalculateActualControlExtents(formObj As Object, ByRef maxWidth As Single, ByRef maxHeight As Single)
    On Error Resume Next
    
    Dim i As Integer
    Dim ctrl As Object
    Dim rightEdge As Single
    Dim bottomEdge As Single
    
    maxWidth = 0
    maxHeight = 0
    
    ' Find the rightmost and bottommost edges of actual controls on the form
    For i = 0 To formObj.Controls.Count - 1
        Set ctrl = formObj.Controls(i)
        
        If Not ctrl Is Nothing Then
            ' Calculate right edge (left + width)
            rightEdge = ctrl.Left + ctrl.Width
            If rightEdge > maxWidth Then maxWidth = rightEdge
            
            ' Calculate bottom edge (top + height)
            bottomEdge = ctrl.Top + ctrl.Height
            If bottomEdge > maxHeight Then maxHeight = bottomEdge
        End If
    Next i
    
    Debug.Print "Actual control extents calculated: " & maxWidth & " x " & maxHeight & " (in points)"
    Debug.Print "Number of controls analyzed: " & formObj.Controls.Count
    On Error GoTo 0
End Sub

' =====================================================================================
' FORM RESIZE UTILITIES
' =====================================================================================
' Public utilities for dynamically resizing forms

' Public function to be called from UserForm_Initialize event
Public Sub ResizeFormToFitControls(formObj As Object)
    On Error Resume Next
    
    Debug.Print "=== Resizing Form to Fit Controls ==="
    Debug.Print "Current form size: " & formObj.Width & " x " & formObj.Height & " (twips)"
    Debug.Print "Current form size in points: " & (formObj.Width / 20) & " x " & (formObj.Height / 20)
    
    Dim requiredWidth As Single
    Dim requiredHeight As Single
    Dim padding As Single
    Dim minWidth As Single
    Dim minHeight As Single
    
    padding = 40  ' Increased to 40 for more generous spacing
    minWidth = 450  ' Increased minimum width
    minHeight = 300  ' Increased minimum height
    
    ' Calculate based on actual controls on the form
    Call CalculateActualControlExtents(formObj, requiredWidth, requiredHeight)
    
    ' Add padding
    requiredWidth = requiredWidth + (padding * 2)
    requiredHeight = requiredHeight + (padding * 2)
    
    ' Ensure minimum dimensions
    If requiredWidth < minWidth Then requiredWidth = minWidth
    If requiredHeight < minHeight Then requiredHeight = minHeight
    
    ' Apply the calculated dimensions (convert points to twips)
    formObj.Width = requiredWidth * 20
    formObj.Height = requiredHeight * 20
    
    Debug.Print "Form resized to: " & formObj.Width & " x " & formObj.Height & " (twips)"
    Debug.Print "Form resized to points: " & (formObj.Width / 20) & " x " & (formObj.Height / 20)
    On Error GoTo 0
End Sub

' =====================================================================================
' GENERAL UTILITY FUNCTIONS
' =====================================================================================
' Miscellaneous helper functions


' =====================================================================================
' FORM EXPORT SYSTEM - VBE EXPORT METHOD
' =====================================================================================
' Functions for exporting forms using VBA Editor's native export capabilities

Public Function ExportFormAsFile(formName As String, designFile As String, codeFile As String, appPath As String) As String
    ' Build a valid .frm by letting VBE generate it via Export
    Dim vbProj As Object
    Dim formComp As Object
    Dim formObj As Object
    Dim exportFolder As String
    Dim exportPath As String
    Dim tmpName As String
    Dim designData As Object
    
    On Error GoTo ErrorHandler
    
    Set vbProj = GetHostVBProject()
    exportFolder = appPath & "\forms\"
    exportPath = exportFolder & formName & ".frm"
    
    ' Ensure folder exists
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(exportFolder) Then fso.CreateFolder exportFolder
    
    ' Create a temporary form component
    Set formComp = vbProj.VBComponents.Add(vbext_ct_MSForm)
    If formComp Is Nothing Then GoTo ErrorHandler
    
    ' Try to rename to desired name BEFORE accessing Designer
    tmpName = formComp.Name
    Call SafeRenameVBComponent(vbProj, formComp, formName)
    
    ' Get designer and apply design
    Set formObj = formComp.Designer
    If formObj Is Nothing Then GoTo ErrorHandler
    
    ' Apply design WITHOUT trying to manipulate form size during creation
    If Dir(designFile) <> "" Then
        Set designData = ParseFormDesign(designFile)
        If Not designData Is Nothing Then 
            Debug.Print "Applying form design (controls only)..."
            Call ApplyFormDesignControlsOnly(formObj, designData)
        End If
    End If
    
    If Dir(codeFile) <> "" Then
        Call ImportCodeBehind(formComp, codeFile)
    End If
    
    Debug.Print "Skipping form pre-sizing to avoid object corruption - will rely on post-export processing"
    
    ' Export using VBE to guarantee correct format and resources
    On Error GoTo ErrorHandler
    formComp.Export exportPath
    Debug.Print "Form exported successfully to: " & exportPath
    
    ' Post-process exported .frm to enforce metadata and size from design
    If Not designData Is Nothing Then
        Debug.Print "Post-processing exported form file..."
        Call AdjustExportedFormMetadata(exportPath, formName, designData)
        Call AdjustExportedFormSize(exportPath, designData)
        ' CRITICAL: Calculate and apply correct form size based on controls
        Call AdjustExportedFormSizeFromControls(exportPath, designData)
        Debug.Print "Post-processing complete - form should have correct dimensions in .frm file"
    Else
        ' Always ensure VB_Name and Begin name align even without design
        Call AdjustExportedFormNames(exportPath, formName)
    End If
    
    ' Clean up: remove the temporary form from the project
    vbProj.VBComponents.Remove formComp
    
    Debug.Print "Exported form for manual import: " & exportPath
    ExportFormAsFile = exportPath
    Exit Function
    
ErrorHandler:
    Dim errDesc As String
    Dim errNum As Long
    errDesc = Err.Description
    errNum = Err.Number
    
    On Error Resume Next
    If Not formComp Is Nothing Then vbProj.VBComponents.Remove formComp
    
    Debug.Print "Error exporting form file via VBE:"
    Debug.Print "  Error Number: " & errNum
    Debug.Print "  Error Description: " & errDesc
    Debug.Print "  Export Path: " & exportPath
    Debug.Print "  Form Name: " & formName
    If Not formObj Is Nothing Then
        Debug.Print "  Form Size: " & formObj.Width & " x " & formObj.Height & " twips"
    End If
    
    ExportFormAsFile = ""
End Function

Private Sub AdjustExportedFormSize(frmFile As String, designData As Object)
    On Error Resume Next
    
    Dim content As String
    content = ReadTextFile(frmFile)
    If content = "" Then Exit Sub
    
    Dim widthPts As Double, heightPts As Double
    Dim hasExplicitW As Boolean, hasExplicitH As Boolean
    
    ' Check for explicit dimensions first
    hasExplicitW = Not designData Is Nothing And designData.Exists("width")
    hasExplicitH = Not designData Is Nothing And designData.Exists("height")
    
    If hasExplicitW Then
        widthPts = CDbl(designData("width"))
    End If
    If hasExplicitH Then
        heightPts = CDbl(designData("height"))
    End If
    
    ' If no explicit dimensions, calculate from controls
    If Not hasExplicitW Or Not hasExplicitH Then
        Dim requiredWidth As Single, requiredHeight As Single
        Dim padding As Single, minWidth As Single, minHeight As Single
        
        padding = 40
        minWidth = 450
        minHeight = 300
        
        If Not designData Is Nothing And designData.Exists("controls") Then
            Call CalculateControlExtentsFromDesign(designData("controls"), requiredWidth, requiredHeight)
            
            ' Add padding
            requiredWidth = requiredWidth + (padding * 2)
            requiredHeight = requiredHeight + (padding * 2)
            
            ' Ensure minimum dimensions
            If requiredWidth < minWidth Then requiredWidth = minWidth
            If requiredHeight < minHeight Then requiredHeight = minHeight
            
            If Not hasExplicitW Then widthPts = requiredWidth
            If Not hasExplicitH Then heightPts = requiredHeight
            
            Debug.Print "Calculated form size for export: " & widthPts & " x " & heightPts & " points"
        Else
            ' No controls and no explicit size - use defaults
            If Not hasExplicitW Then widthPts = minWidth
            If Not hasExplicitH Then heightPts = minHeight
        End If
    End If
    
    ' Convert points → twips (20 twips per point)
    Dim widthTwips As Long, heightTwips As Long
    widthTwips = CLng(widthPts * 20#)
    heightTwips = CLng(heightPts * 20#)
    
    Debug.Print "Export sizing - Width: " & widthTwips & " twips (" & widthPts & " pts)"
    Debug.Print "Export sizing - Height: " & heightTwips & " twips (" & heightPts & " pts)"
    
    ' Replace ClientWidth/ClientHeight lines if present
    content = ReplaceClientMetric(content, "ClientWidth", CStr(widthTwips))
    content = ReplaceClientMetric(content, "ClientHeight", CStr(heightTwips))
    
    Dim fnum As Integer: fnum = FreeFile
    Open frmFile For Output As fnum
    Print #fnum, content
    Close fnum
    On Error GoTo 0
End Sub

Private Sub CalculateControlExtentsFromDesign(controlsCollection As Object, ByRef maxWidth As Single, ByRef maxHeight As Single)
    On Error Resume Next
    
    Dim i As Integer
    Dim controlData As Object
    Dim rightEdge As Single
    Dim bottomEdge As Single
    
    maxWidth = 0
    maxHeight = 0
    
    ' Find the rightmost and bottommost edges of all controls
    For i = 1 To controlsCollection.Count
        Set controlData = controlsCollection(i)
        
        If Not controlData Is Nothing Then
            ' Calculate right edge (left + width)
            If controlData.Exists("left") And controlData.Exists("width") Then
                rightEdge = CSng(controlData("left")) + CSng(controlData("width"))
                If rightEdge > maxWidth Then maxWidth = rightEdge
            End If
            
            ' Calculate bottom edge (top + height)
            If controlData.Exists("top") And controlData.Exists("height") Then
                bottomEdge = CSng(controlData("top")) + CSng(controlData("height"))
                If bottomEdge > maxHeight Then maxHeight = bottomEdge
            End If
        End If
    Next i
    
    Debug.Print "Control extents from design: " & maxWidth & " x " & maxHeight
    On Error GoTo 0
End Sub

Private Function ReplaceClientMetric(content As String, key As String, value As String) As String
    ' Replace the numeric value for a given key formatted like: key =   <number>
    Dim pos As Long, lineStart As Long, lineEnd As Long
    pos = InStr(1, content, key & " ")
    If pos = 0 Then
        ReplaceClientMetric = content
        Exit Function
    End If
    ' Find end of line
    lineStart = pos
    lineEnd = InStr(pos, content, vbCrLf)
    If lineEnd = 0 Then lineEnd = Len(content) + 1
    Dim line As String
    line = Mid$(content, lineStart, lineEnd - lineStart)
    ' Build replacement line with same spacing prefix
    Dim prefix As String
    prefix = Left$(line, InStr(1, line, "=") + 1)
    Dim spaces As String
    spaces = Mid$(line, Len(prefix) + 1)
    ' Normalize to standard spacing
    ReplaceClientMetric = Left$(content, lineStart - 1) & key & Space(5) & "=   " & value & Mid$(content, lineEnd)
End Function

Private Sub AdjustExportedFormMetadata(frmFile As String, formName As String, designData As Object)
    On Error Resume Next
    Dim content As String
    content = ReadTextFile(frmFile)
    If content = "" Then Exit Sub
    
    ' Ensure the first Begin line uses the target form name
    content = Replace(content, "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1", _
                              "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " & formName)
    content = Replace(content, "Attribute VB_Name = ""UserForm1""", _
                              "Attribute VB_Name = """ & formName & """")
    
    ' Caption
    If Not designData Is Nothing Then
        If designData.Exists("caption") Then
            Dim cap As String: cap = CStr(designData("caption"))
            content = ReplaceCaptionLine(content, cap)
        End If
    End If
    
    Dim fnum As Integer: fnum = FreeFile
    Open frmFile For Output As fnum
    Print #fnum, content
    Close fnum
    On Error GoTo 0
End Sub

Private Sub AdjustExportedFormNames(frmFile As String, formName As String)
    On Error Resume Next
    Dim content As String
    content = ReadTextFile(frmFile)
    If content = "" Then Exit Sub
    content = Replace(content, "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1", _
                              "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " & formName)
    content = Replace(content, "Attribute VB_Name = ""UserForm1""", _
                              "Attribute VB_Name = """ & formName & """")
    Dim fnum As Integer: fnum = FreeFile
    Open frmFile For Output As fnum
    Print #fnum, content
    Close fnum
    On Error GoTo 0
End Sub

Private Sub AdjustExportedFormSizeFromControls(frmFile As String, designData As Object)
    On Error Resume Next
    
    If designData Is Nothing Or Not designData.Exists("controls") Then Exit Sub
    
    ' Calculate required size based on controls
    Dim requiredWidth As Single
    Dim requiredHeight As Single
    Dim padding As Single
    Dim minWidth As Single
    Dim minHeight As Single
    
    padding = 40  ' Increased to 40 for more generous spacing
    minWidth = 450  ' Increased minimum width
    minHeight = 300  ' Increased minimum height
    
    ' Calculate control extents from design data
    Call CalculateControlExtentsFromDesign(designData("controls"), requiredWidth, requiredHeight)
    
    ' Add padding
    requiredWidth = requiredWidth + (padding * 2)
    requiredHeight = requiredHeight + (padding * 2)
    
    ' Ensure minimum dimensions
    If requiredWidth < minWidth Then requiredWidth = minWidth
    If requiredHeight < minHeight Then requiredHeight = minHeight
    
    ' Convert to twips
    Dim widthTwips As Long, heightTwips As Long
    widthTwips = CLng(requiredWidth * 20#)
    heightTwips = CLng(requiredHeight * 20#)
    
    Debug.Print "Pre-calculated form size: " & widthTwips & " x " & heightTwips & " (twips)"
    Debug.Print "Pre-calculated form size in points: " & requiredWidth & " x " & requiredHeight
    
    ' Read the .frm file content
    Dim content As String
    content = ReadTextFile(frmFile)
    If content = "" Then Exit Sub
    
    ' Replace ClientWidth/ClientHeight with calculated values
    content = ReplaceClientMetric(content, "ClientWidth", CStr(widthTwips))
    content = ReplaceClientMetric(content, "ClientHeight", CStr(heightTwips))
    
    ' Write the updated content back to the file
    Dim fnum As Integer: fnum = FreeFile
    Open frmFile For Output As fnum
    Print #fnum, content
    Close fnum
    
    Debug.Print "Updated .frm file with pre-calculated size"
    On Error GoTo 0
End Sub

Private Function ReplaceCaptionLine(content As String, newCaption As String) As String
    Dim pos As Long, lineStart As Long, lineEnd As Long
    pos = InStr(1, content, "Caption         =")
    If pos = 0 Then
        ReplaceCaptionLine = content
        Exit Function
    End If
    lineStart = InStrRev(content, vbCrLf, pos)
    If lineStart = 0 Then lineStart = 1 Else lineStart = lineStart + 2
    lineEnd = InStr(pos, content, vbCrLf)
    If lineEnd = 0 Then lineEnd = Len(content) + 1
    ReplaceCaptionLine = Left$(content, lineStart - 1) & _
                         "   Caption         =   """ & newCaption & """" & vbCrLf & _
                         Mid$(content, lineEnd)
End Function

' Parse JSON Text into VBA Dictionary Object
'
' A robust JSON parser specifically designed for VBA Build System needs.
' Handles the specific JSON structures used in manifests and form designs.
' Supports nested objects, arrays, and mixed data types.
'
' PARAMETERS:
'   jsonText - Raw JSON string to parse
'
' RETURNS: Dictionary object with parsed data, Nothing on error
'
' SUPPORTED STRUCTURES:
' • Simple key-value pairs (strings, numbers, booleans)
' • Nested objects (font, dependencies)
' • Arrays (controls, references)
' • Mixed data types with automatic conversion
Public Function ParseSimpleJSON(jsonText As String) As Object
    ' Robust JSON parser for VBA applications
    On Error GoTo ErrorHandler
    
    Debug.Print "=== JSON PARSER DEBUG ==="
    Debug.Print "Input JSON length: " & Len(jsonText)
    Debug.Print "First 200 chars: " & Left(jsonText, 200)
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Debug.Print "Dictionary created successfully"
    
    ' Extract basic key-value pairs (only from top level, not from nested objects)
    Dim topLevelJson As String
    
    ' If this has a controls array, extract only the part before it for top-level properties
    Dim controlsPos As Integer
    controlsPos = InStr(jsonText, """controls""")
    If controlsPos > 0 Then
        topLevelJson = Left(jsonText, controlsPos - 1) & "}"
    Else
        topLevelJson = jsonText
    End If
    
    ' Extract basic key-value pairs from top-level only
    If InStr(topLevelJson, """name""") > 0 Then
        dict("name") = ExtractJSONValue(topLevelJson, "name")
    End If
    If InStr(topLevelJson, """version""") > 0 Then
        dict("version") = ExtractJSONValue(topLevelJson, "version")
    End If
    If InStr(topLevelJson, """modules""") > 0 Then
        dict("modules") = ExtractJSONValue(topLevelJson, "modules")
    End If
    If InStr(topLevelJson, """forms""") > 0 Then
        dict("forms") = ExtractJSONValue(topLevelJson, "forms")
    End If
    If InStr(topLevelJson, """entryPoint""") > 0 Then
        dict("entryPoint") = ExtractJSONValue(topLevelJson, "entryPoint")
    End If
    If InStr(topLevelJson, """caption""") > 0 Then
        dict("caption") = ExtractJSONValue(topLevelJson, "caption")
    End If
    If InStr(topLevelJson, """width""") > 0 Then
        dict("width") = CLng(ExtractJSONValue(topLevelJson, "width"))
    End If
    If InStr(topLevelJson, """height""") > 0 Then
        dict("height") = CLng(ExtractJSONValue(topLevelJson, "height"))
    End If
    If InStr(topLevelJson, """startUpPosition""") > 0 Then
        Dim supRaw As String
        supRaw = ExtractJSONValue(topLevelJson, "startUpPosition")
        ' Allow numeric or string forms (e.g., "CenterOwner")
        If IsNumeric(supRaw) Then
            dict("startUpPosition") = CLng(supRaw)
        Else
            ' Normalize common MSForms constants to numeric values
            Select Case LCase$(Trim$(supRaw))
                Case "centerscreen", "0"
                    dict("startUpPosition") = 0
                Case "centerowner", "1"
                    dict("startUpPosition") = 1
                Case "manual", "2"
                    dict("startUpPosition") = 2
                Case Else
                    ' Default to CenterOwner if unknown string
                    dict("startUpPosition") = 1
            End Select
        End If
    End If
    
    ' Note: Removed left, top, type, etc. since these are typically control properties
    
    ' Handle "controls" array
    If InStr(jsonText, """controls""") > 0 Then
        Dim controlsArray As Object
        Set controlsArray = ParseJSONArray(jsonText, "controls")
        If Not controlsArray Is Nothing Then
            ' Use Set for object assignment since controlsArray is a Collection
            On Error Resume Next
            Set dict("controls") = controlsArray
            If Err.Number <> 0 Then
                Err.Clear
                ' Try without Set
                dict("controls") = controlsArray
                If Err.Number <> 0 Then
                    Err.Clear
                End If
            End If
            On Error GoTo 0
        End If
    End If
    
    ' Handle "dependencies" object
    If InStr(jsonText, """dependencies""") > 0 Then
        Dim depsObj As Object
        Set depsObj = ParseNestedObject(jsonText, "dependencies")
        If Not depsObj Is Nothing Then
            Set dict("dependencies") = depsObj
        End If
    End If
    
    Debug.Print "JSON parsing completed successfully"
    Debug.Print "Dictionary contains " & dict.Count & " items"
    
    Set ParseSimpleJSON = dict
    Exit Function
    
ErrorHandler:
    Debug.Print "❌ JSON PARSER ERROR: " & Err.Number & " - " & Err.Description
    Debug.Print "Error occurred at line in ParseSimpleJSON"
    Set ParseSimpleJSON = Nothing
End Function

Public Function ExtractJSONValue(jsonText As String, key As String) As String
    ' Very basic JSON value extraction with improved type handling
    Dim startPos As Integer
    Dim endPos As Integer
    Dim searchKey As String
    
    searchKey = """" & key & """:"
    startPos = InStr(jsonText, searchKey)
    If startPos = 0 Then
        ExtractJSONValue = ""
        Exit Function
    End If
    
    startPos = startPos + Len(searchKey)
    ' Skip whitespace
    Do While Mid(jsonText, startPos, 1) = " " Or Mid(jsonText, startPos, 1) = vbTab
        startPos = startPos + 1
    Loop
    
    ' Check if value is quoted (string) or unquoted (number/boolean)
    Dim isQuoted As Boolean
    isQuoted = (Mid(jsonText, startPos, 1) = """")
    
    If isQuoted Then
        startPos = startPos + 1  ' Skip opening quote
    End If
    
    ' Find end of value
    endPos = startPos
    Do While endPos <= Len(jsonText)
        Dim ch As String
        ch = Mid(jsonText, endPos, 1)
        
        If isQuoted Then
            ' For quoted strings, look for closing quote
            If ch = """" Then Exit Do
        Else
            ' For unquoted values, look for comma, brace, or bracket
            If ch = "," Or ch = "}" Or ch = "]" Or ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then Exit Do
        End If
        endPos = endPos + 1
    Loop
    
    ExtractJSONValue = Mid(jsonText, startPos, endPos - startPos)
End Function

Public Function ParseJSONArray(jsonText As String, key As String) As Object
    ' Robust JSON array parser with proper object splitting
    Dim arr As Collection
    Set arr = New Collection ' Use VBA Collection instead of ArrayList
    
    Dim arrayContent As String
    arrayContent = ExtractJSONArrayContent(jsonText, key)
    
    If arrayContent = "" Then
        Set ParseJSONArray = arr
        Exit Function
    End If
    
    ' Parse objects using proper brace counting instead of crude splitting
    Dim objects As Collection
    Set objects = SplitJSONObjects(arrayContent)
    
    Dim i As Integer
    For i = 1 To objects.Count
        Dim objText As String
        objText = objects(i)
        
        Dim controlDict As Object
        On Error Resume Next
        Set controlDict = ParseControlObject(objText) ' Use dedicated control parser
        If Err.Number <> 0 Then
            Err.Clear
            Set controlDict = Nothing
        End If
        On Error GoTo 0
        
        If Not controlDict Is Nothing Then
            arr.Add controlDict ' Collection.Add instead of ArrayList.Add
        End If
    Next i
    
    Set ParseJSONArray = arr
End Function

Private Function SplitJSONObjects(arrayContent As String) As Collection
    ' Split JSON array content into individual objects using proper brace counting
    Dim objects As Collection
    Set objects = New Collection
    
    Dim i As Integer
    Dim braceCount As Integer
    Dim startPos As Integer
    Dim inString As Boolean
    Dim escaped As Boolean
    
    startPos = 1
    braceCount = 0
    inString = False
    escaped = False
    
    ' Skip leading whitespace and find first brace
    Do While startPos <= Len(arrayContent) And Mid(arrayContent, startPos, 1) <> "{"
        startPos = startPos + 1
    Loop
    
    If startPos > Len(arrayContent) Then
        Set SplitJSONObjects = objects
        Exit Function
    End If
    
    Dim objStart As Integer
    objStart = startPos
    
    For i = startPos To Len(arrayContent)
        Dim char As String
        char = Mid(arrayContent, i, 1)
        
        ' Handle string parsing to ignore braces inside strings
        If char = """" And Not escaped Then
            inString = Not inString
        ElseIf char = "\" Then
            escaped = Not escaped
        Else
            escaped = False
        End If
        
        If Not inString Then
            If char = "{" Then
                braceCount = braceCount + 1
            ElseIf char = "}" Then
                braceCount = braceCount - 1
                
                ' When we close all braces, we have a complete object
                If braceCount = 0 Then
                    Dim objText As String
                    objText = Mid(arrayContent, objStart, i - objStart + 1)
                    objects.Add objText
                    
                    ' Find next object start
                    objStart = i + 1
                    Do While objStart <= Len(arrayContent) And Mid(arrayContent, objStart, 1) <> "{"
                        objStart = objStart + 1
                    Loop
                    
                    If objStart <= Len(arrayContent) Then
                        i = objStart - 1  ' Will be incremented by For loop
                    End If
                End If
            End If
        End If
    Next i
    
    Set SplitJSONObjects = objects
End Function

Public Function ParseControlObject(objText As String) As Object
    ' Parse a single control object without recursing back to ParseSimpleJSON
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    ' Extract control-specific properties
    If InStr(objText, """name""") > 0 Then
        dict("name") = ExtractJSONValue(objText, "name")
    End If
    If InStr(objText, """type""") > 0 Then
        dict("type") = ExtractJSONValue(objText, "type")
    End If
    If InStr(objText, """caption""") > 0 Then
        dict("caption") = ExtractJSONValue(objText, "caption")
    End If
    If InStr(objText, """left""") > 0 Then
        dict("left") = CLng(ExtractJSONValue(objText, "left"))
    End If
    If InStr(objText, """top""") > 0 Then
        dict("top") = CLng(ExtractJSONValue(objText, "top"))
    End If
    If InStr(objText, """width""") > 0 Then
        dict("width") = CLng(ExtractJSONValue(objText, "width"))
    End If
    If InStr(objText, """height""") > 0 Then
        dict("height") = CLng(ExtractJSONValue(objText, "height"))
    End If
    If InStr(objText, """text""") > 0 Then
        dict("text") = ExtractJSONValue(objText, "text")
    End If
    If InStr(objText, """multiSelect""") > 0 Then
        dict("multiSelect") = ExtractJSONValue(objText, "multiSelect")
    End If
    
    ' Handle nested font object
    If InStr(objText, """font""") > 0 Then
        Dim fontDict As Object
        Set fontDict = ParseNestedObject(objText, "font")
        If Not fontDict Is Nothing Then
            Set dict("font") = fontDict
        End If
    End If
    
    Set ParseControlObject = dict
    On Error GoTo 0
End Function

Public Function ParseNestedObject(jsonText As String, key As String) As Object
    ' Parse a nested object like font properties or dependencies
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    Dim objContent As String
    objContent = ExtractNestedObjectContent(jsonText, key)
    
    If objContent = "" Then
        Set ParseNestedObject = Nothing
        Exit Function
    End If
    
    ' Parse font-specific properties
    If InStr(objContent, """size""") > 0 Then
        dict("size") = CLng(ExtractJSONValue(objContent, "size"))
    End If
    If InStr(objContent, """bold""") > 0 Then
        dict("bold") = CBool(ExtractJSONValue(objContent, "bold"))
    End If
    If InStr(objContent, """italic""") > 0 Then
        dict("italic") = CBool(ExtractJSONValue(objContent, "italic"))
    End If
    If InStr(objContent, """name""") > 0 Then
        dict("name") = ExtractJSONValue(objContent, "name")
    End If
    
    ' Parse dependencies-specific properties
    If InStr(objContent, """references""") > 0 Then
        Dim referencesContent As String
        referencesContent = ExtractJSONArrayContent(objContent, "references")
        If referencesContent <> "" Then
            ' Parse string array manually since it's simpler than object array
            Dim refArray As Collection
            Set refArray = New Collection
            
            ' Split by comma and clean up quotes
            Dim refs As Variant
            refs = Split(referencesContent, ",")
            
            Dim i As Integer
            For i = 0 To UBound(refs)
                Dim refStr As String
                refStr = Trim(refs(i))
                
                ' Remove all quotes and extra whitespace/newlines
                refStr = Replace(refStr, """", "")
                refStr = Replace(refStr, vbCr, "")
                refStr = Replace(refStr, vbLf, "")
                refStr = Replace(refStr, vbTab, "")
                refStr = Trim(refStr)
                
                ' Only add non-empty references
                If refStr <> "" Then
                    refArray.Add refStr
                End If
            Next i
            
            Set dict("references") = refArray
        End If
    End If
    
    Set ParseNestedObject = dict
    On Error GoTo 0
End Function

Public Function ExtractNestedObjectContent(jsonText As String, key As String) As String
    ' Extract content of a nested object like {"font": {"size": 14, "bold": true}}
    Dim startPos As Integer
    Dim endPos As Integer
    Dim searchKey As String
    Dim braceCount As Integer
    
    ' Find the key followed by colon and opening brace
    searchKey = """" & key & """"
    Dim keyPos As Integer
    keyPos = InStr(jsonText, searchKey)
    
    If keyPos = 0 Then
        ExtractNestedObjectContent = ""
        Exit Function
    End If
    
    ' Look for colon after key
    Dim colonPos As Integer
    colonPos = InStr(keyPos, jsonText, ":")
    If colonPos = 0 Then
        ExtractNestedObjectContent = ""
        Exit Function
    End If
    
    ' Find opening brace after colon
    Dim bracePos As Integer
    bracePos = colonPos + 1
    Do While bracePos <= Len(jsonText)
        Dim char As String
        char = Mid(jsonText, bracePos, 1)
        If char = "{" Then
            Exit Do
        ElseIf char <> " " And char <> vbTab And char <> vbCr And char <> vbLf Then
            ExtractNestedObjectContent = ""
            Exit Function
        End If
        bracePos = bracePos + 1
    Loop
    
    If bracePos > Len(jsonText) Then
        ExtractNestedObjectContent = ""
        Exit Function
    End If
    
    startPos = bracePos + 1
    
    ' Find matching closing brace
    braceCount = 1
    endPos = startPos
    Do While endPos <= Len(jsonText) And braceCount > 0
        Dim charX As String
        charX = Mid(jsonText, endPos, 1)
        If charX = "{" Then
            braceCount = braceCount + 1
        ElseIf charX = "}" Then
            braceCount = braceCount - 1
        End If
        
        If braceCount = 0 Then Exit Do
        endPos = endPos + 1
    Loop
    
    If braceCount = 0 Then
        ExtractNestedObjectContent = Mid(jsonText, startPos, endPos - startPos)
    Else
        ExtractNestedObjectContent = ""
    End If
End Function

Public Function ExtractJSONArrayContent(jsonText As String, key As String) As String
    ' Extracts the content string from a JSON array
    Dim startPos As Integer
    Dim endPos As Integer
    Dim searchKey As String
    Dim bracketCount As Integer
    
    ' First find the key, then look for the array bracket
    Dim keyPos As Integer
    Dim searchPattern As String
    searchPattern = """" & key & """"
    keyPos = InStr(jsonText, searchPattern)
    
    If keyPos = 0 Then
        ExtractJSONArrayContent = ""
        Exit Function
    End If
    
    ' Look for the colon after the key
    Dim colonPos As Integer
    colonPos = InStr(keyPos, jsonText, ":")
    If colonPos = 0 Then
        ExtractJSONArrayContent = ""
        Exit Function
    End If
    
    ' Look for opening bracket after the colon, skipping whitespace
    Dim bracketPos As Integer
    bracketPos = colonPos + 1
    Do While bracketPos <= Len(jsonText)
        Dim char As String
        char = Mid(jsonText, bracketPos, 1)
        If char = "[" Then
            Exit Do
        ElseIf char <> " " And char <> vbTab And char <> vbCr And char <> vbLf Then
            ExtractJSONArrayContent = ""
            Exit Function
        End If
        bracketPos = bracketPos + 1
    Loop
    
    If bracketPos > Len(jsonText) Then
        ExtractJSONArrayContent = ""
        Exit Function
    End If
    
    startPos = bracketPos + 1
    
    ' Find the closing "]"
    bracketCount = 1
    endPos = startPos
    Do While endPos <= Len(jsonText) And bracketCount > 0
        Dim charX As String
        charX = Mid(jsonText, endPos, 1)
        If charX = "[" Then
            bracketCount = bracketCount + 1
        ElseIf charX = "]" Then
            bracketCount = bracketCount - 1
        End If
        
        If bracketCount = 0 Then Exit Do
        endPos = endPos + 1
    Loop
    
    If bracketCount = 0 Then
        Dim content As String
        content = Mid(jsonText, startPos, endPos - startPos)
        ExtractJSONArrayContent = content
    Else
        ExtractJSONArrayContent = "" ' Unmatched brackets
    End If
End Function