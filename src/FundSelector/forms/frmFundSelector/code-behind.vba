' ===== FUND SELECTOR FORM CODE-BEHIND =====
' This demonstrates event handling patterns for VBA Build System forms
'
' PURPOSE: Fund selection tool with database connectivity
' PATTERNS SHOWN:
'   - Button click event handlers
'   - Form lifecycle events (Initialize, Activate, QueryClose)
'   - Module function calls from form events
'   - Error handling in event procedures
'   - ListBox interaction patterns
'   - Database connectivity management
'
' FOR AI ASSISTANTS: Use these patterns when generating form code-behind
' FOR DEVELOPERS: Copy these event handling approaches

Option Explicit

' ===== FORM LIFECYCLE EVENTS =====

Private Sub UserForm_Initialize()
    ' Runs when form is first created - use for setup
    Debug.Print "=== FundSelector Form Initializing ==="
    
    ' Example: Set default values, configure controls
    Me.Caption = "Fund Selection Tool - VBA Build System"
    
    ' Configure the title label
    lblTitle.Caption = "General Partner Fund Selection"
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 14
    
    ' Configure description label
    lblDescription.Caption = "Select a General Partner from the list below to view their fund information."
    
    ' Load initial data
    Call LoadGeneralPartners
    
    Debug.Print "Form initialized successfully"
End Sub

Private Sub UserForm_Activate()
    ' Runs when form becomes active - use for refresh operations
    Debug.Print "FundSelector form activated"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Runs before form closes - use for cleanup or confirmation
    If CloseMode = vbFormControlMenu Then
        ' User clicked the X button
        Debug.Print "User closing form via X button"
    End If
    
    ' Clean up database connections
    Call modFundSelector.CloseDatabase
End Sub

' ===== BUTTON CLICK EVENT HANDLERS =====

Private Sub btnRefresh_Click()
    ' Refresh the general partners list
    Debug.Print "Refresh button clicked"
    
    ' Call module function - this is the recommended pattern
    Call modFundSelector.RefreshGeneralPartnersList
End Sub

Private Sub btnTestConnection_Click()
    ' Test database connection
    Debug.Print "Test Connection button clicked"
    
    ' Another module function call example
    Call modFundSelector.TestDatabaseConnection
End Sub

Private Sub btnClose_Click()
    ' Standard close button pattern
    Debug.Print "Close button clicked"
    
    ' Clean up database connections
    Call modFundSelector.CloseDatabase
    
    ' Hide the form (recommended over Unload for reusability)
    Me.Hide
    
    ' Alternative: Unload the form completely
    ' Unload Me
End Sub

' ===== LISTBOX INTERACTION EVENTS =====

Private Sub lstGeneralPartners_Click()
    ' Handle list box selection
    Debug.Print "General Partner selected: " & lstGeneralPartners.ListIndex
    
    ' Call module function to handle selection
    Call modFundSelector.HandleGeneralPartnerSelection(lstGeneralPartners.ListIndex)
End Sub

Private Sub lstGeneralPartners_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Handle double-click on list box item
    Debug.Print "General Partner double-clicked: " & lstGeneralPartners.ListIndex
    
    ' Call module function for double-click action
    Call modFundSelector.HandleGeneralPartnerDoubleClick(lstGeneralPartners.ListIndex)
End Sub

' ===== INTERNAL HELPER FUNCTIONS =====

Private Sub LoadGeneralPartners()
    ' Load general partners into the list box
    On Error GoTo ErrorHandler
    
    Debug.Print "Loading general partners..."
    
    ' Call module function to get data
    Call modFundSelector.LoadGeneralPartnersIntoList(lstGeneralPartners)
    
    Debug.Print "General partners loaded successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error loading general partners: " & Err.Description
    lstGeneralPartners.AddItem "Error loading data. Please try refreshing."
End Sub

' ===== ADVANCED EVENT HANDLING EXAMPLES =====

Private Sub UserForm_Click()
    ' Form click event - useful for focus management
    Debug.Print "Form background clicked"
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Double-click on form - example of advanced interaction
    Debug.Print "Form double-clicked"
    Call modFundSelector.DemonstrateDebugging
End Sub

' ===== ERROR HANDLING PATTERNS =====

Private Sub btnRefresh_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Example of mouse event with error handling
    On Error GoTo ErrorHandler
    
    If Button = 1 Then ' Left mouse button
        Debug.Print "Left mouse button pressed on Refresh button"
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
' - ListBox: _Click(), _DblClick() for selection handling
' - TextBox: _Change(), _Enter(), _Exit() for validation
' - Form: _Initialize(), _Activate(), _QueryClose() for lifecycle
'
' ERROR HANDLING:
' - Use On Error GoTo ErrorHandler in complex events
' - Avoid message boxes in mouse/keyboard events
' - Log errors to Debug.Print for troubleshooting
' - Resume Next or Exit Sub appropriately 