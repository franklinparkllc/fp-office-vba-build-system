VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExampleApp 
   Caption         =   "Example Application - VBA Build System Template"

   ClientHeight     =   6400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   OleObjectBlob   =   "frmExampleApp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExampleApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===== EXAMPLE APPLICATION FORM CODE-BEHIND =====
' This demonstrates event handling patterns for VBA Build System forms
'
' PURPOSE: Reference template for AI assistants and developers
' PATTERNS SHOWN:
'   - Button click event handlers
'   - Form lifecycle events (Initialize, Activate, QueryClose)
'   - Module function calls from form events
'   - Error handling in event procedures
'   - User interaction patterns
'
' FOR AI ASSISTANTS: Use these patterns when generating form code-behind
' FOR DEVELOPERS: Copy these event handling approaches

Option Explicit

' ===== FORM LIFECYCLE EVENTS =====

Private Sub UserForm_Initialize()
    ' Runs when form is first created - use for setup
    Debug.Print "=== ExampleApp Form Initializing ==="
    
    ' Example: Set default values, configure controls
    Me.Caption = "Example Application - VBA Build System"
    
    ' Configure the title label
    lblTitle.Caption = "VBA Build System Example"
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 12
    
    Debug.Print "Form initialized successfully"
End Sub

Private Sub UserForm_Activate()
    ' Runs when form becomes active - use for refresh operations
    Debug.Print "ExampleApp form activated"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Runs before form closes - use for cleanup or confirmation
    If CloseMode = vbFormControlMenu Then
        ' User clicked the X button
        Debug.Print "User closing form via X button"
    End If
    
    ' Example: Uncomment to show confirmation dialog
    ' Dim result As VbMsgBoxResult
    ' result = MsgBox("Are you sure you want to close?", vbYesNo + vbQuestion, "Confirm Close")
    ' If result = vbNo Then Cancel = True
End Sub

' ===== BUTTON CLICK EVENT HANDLERS =====

Private Sub btnHello_Click()
    ' Primary action button - demonstrates module function call
    Debug.Print "Hello button clicked"
    
    ' Call module function - this is the recommended pattern
    Call modExampleApp.ShowHelloMessage
End Sub

Private Sub btnInfo_Click()
    ' Secondary action button - shows system information
    Debug.Print "Info button clicked"
    
    ' Another module function call example
    Call modExampleApp.ShowSystemInfo
End Sub

Private Sub btnClose_Click()
    ' Standard close button pattern
    Debug.Print "Close button clicked"
    
    ' Hide the form (recommended over Unload for reusability)
    Me.Hide
    
    ' Alternative: Unload the form completely
    ' Unload Me
End Sub

' ===== ADVANCED EVENT HANDLING EXAMPLES =====

Private Sub UserForm_Click()
    ' Form click event - useful for focus management
    Debug.Print "Form background clicked"
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Double-click on form - example of advanced interaction
    Debug.Print "Form double-clicked"
    Call modExampleApp.DemonstrateDebugging
End Sub

' ===== ERROR HANDLING PATTERNS =====

Private Sub btnHello_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Example of mouse event with error handling
    On Error GoTo ErrorHandler
    
    If Button = 1 Then ' Left mouse button
        Debug.Print "Left mouse button pressed on Hello button"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in mouse event: " & Err.Description
    ' Don't show message box in mouse events - can cause issues
End Sub

' ===== AI ASSISTANT GUIDELINES FOR FORM CODE-BEHIND =====
'
' When generating form code-behind based on this template:
'
' 1. ALWAYS include Option Explicit at the top
' 2. Use descriptive event handler names matching control names
' 3. Call module functions rather than putting logic in form events
' 4. Include Debug.Print statements for troubleshooting
' 5. Handle form lifecycle events when needed (Initialize, QueryClose)
' 6. Use proper error handling in complex event procedures
' 7. Prefer Me.Hide over Unload Me for form closure (better reusability)
' 8. Keep event handlers simple - complex logic belongs in modules
'
' CONTROL EVENT PATTERNS:
' - CommandButton: _Click() is primary event
' - TextBox: _Change(), _Enter(), _Exit() for validation
' - ListBox/ComboBox: _Click(), _Change() for selection handling
' - Form: _Initialize(), _Activate(), _QueryClose() for lifecycle
'
' ERROR HANDLING:
' - Use On Error GoTo ErrorHandler in complex events
' - Avoid message boxes in mouse/keyboard events
' - Log errors to Debug.Print for troubleshooting
' - Resume Next or Exit Sub appropriately




