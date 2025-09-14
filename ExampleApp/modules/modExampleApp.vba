' ===== EXAMPLE APPLICATION MODULE =====
' This is a reference template for VBA Build System applications
' 
' PURPOSE: Demonstrates common patterns for AI assistants and developers
' PATTERNS SHOWN:
'   - Simple message display functions
'   - Direct VBA form launching (recommended approach)
'   - Error handling with fallback strategies
'   - Module-to-form communication
'   - Debug output for troubleshooting
'
' FOR AI ASSISTANTS: Use this as a template when generating VBA applications
' FOR DEVELOPERS: Copy this structure for new applications

Attribute VB_Name = "modExampleApp"
Option Explicit

' ===== BASIC FUNCTIONALITY EXAMPLES =====

Public Sub ShowHelloMessage()
    ' Simple message box - basic VBA interaction pattern
    MsgBox "Hello from ExampleApp!", vbInformation, "Example Application"
End Sub

Public Sub ShowCustomMessage(message As String, Optional title As String = "ExampleApp")
    ' Parameterized message - shows function parameter patterns
    MsgBox message, vbInformation, title
End Sub

' ===== FORM LAUNCHING PATTERNS =====

Public Sub LaunchExampleForm()
    ' DIRECT FORM REFERENCE ONLY: The build creates forms BEFORE importing modules
    ' Never reference placeholder names like UserForm1 in modules
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Launching Example Form ==="
    Debug.Print "Attempting to launch: frmExampleApp"
    
    frmExampleApp.Show
    Debug.Print "✅ Successfully launched form: frmExampleApp"
    Exit Sub
    
ErrorHandler:
    Debug.Print "❌ Failed to launch frmExampleApp: " & Err.Description
    MsgBox "Could not launch frmExampleApp. Ensure the build completed successfully and the form was created.", _
          vbCritical, "Form Launch Error"
End Sub

' ===== UTILITY FUNCTIONS =====

Public Sub ShowSystemInfo()
    ' Demonstrates system information gathering
    Dim info As String
    info = "VBA Build System Example" & vbCrLf & vbCrLf
    info = info & "Application: ExampleApp" & vbCrLf
    info = info & "Version: 1.0.0" & vbCrLf
    info = info & "Host: " & Application.Name & vbCrLf
    info = info & "VBA Version: " & Application.Version
    
    MsgBox info, vbInformation, "System Information"
End Sub

' ===== PATTERN EXAMPLES FOR AI REFERENCE =====

Public Sub DemonstrateErrorHandling()
    ' Shows proper error handling patterns for VBA
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Error Handling Demo ==="
    
    ' Simulate some operation that might fail
    Dim result As String
    result = "Operation completed successfully"
    
    Debug.Print "✅ " & result
    Exit Sub
    
ErrorHandler:
    Debug.Print "❌ Error occurred: " & Err.Description
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    Resume Next
End Sub

Public Sub DemonstrateDebugging()
    ' Shows debugging output patterns
    Debug.Print "=== Debug Output Demo ==="
    Debug.Print "This appears in VBA Immediate Window (Ctrl+G)"
    Debug.Print "Timestamp: " & Now()
    Debug.Print "User: " & Environ("USERNAME")
    Debug.Print "Computer: " & Environ("COMPUTERNAME")
End Sub

' ===== AI ASSISTANT GUIDELINES =====
'
' When generating VBA code based on this template:
'
' 1. ALWAYS use DIRECT form references: frmYourForm.Show
'    - The build system creates forms BEFORE importing modules
'    - By the time module code runs, forms already exist
'    - Keep it simple and clean
'    - Never reference placeholder names like UserForm1 in modules
' 2. Include proper error handling for runtime issues
' 3. Add Debug.Print statements for troubleshooting
' 4. Use descriptive function names and comments
' 5. Follow the module naming pattern: modYourAppName
' 6. Include proper error handling in all functions
' 7. Use Option Explicit at the top of every module
' 8. Add Attribute VB_Name for proper module identification
'
' FORM LAUNCHING PATTERN:
' frmYourForm.Show  ' Simple, direct, and works because build creates forms first
'
' BUILD PROCESS ORDER:
' 1. Build system creates forms from design.json
' 2. Build system imports modules (this code)
' 3. Module code can safely reference forms because they exist
'
' AVOID:
' - Complex dynamic form discovery (unnecessary)
' - Collection-based form management
' - External dependencies when possible
' - Hardcoded paths or system-specific references
