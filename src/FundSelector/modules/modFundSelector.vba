' ===== FUND SELECTOR MODULE =====
' This is a reference template for VBA Build System applications
' 
' PURPOSE: Fund selection tool with database connectivity
' PATTERNS SHOWN:
'   - Database connection management
'   - Direct VBA form launching (recommended approach)
'   - Error handling with fallback strategies
'   - Module-to-form communication
'   - Debug output for troubleshooting
'   - ListBox data management
'
' FOR AI ASSISTANTS: Use this as a template when generating VBA applications
' FOR DEVELOPERS: Copy this structure for new applications

Attribute VB_Name = "modFundSelector"
Option Explicit

' ===== DATABASE CONNECTION MANAGEMENT =====

' Azure SQL Database connection string
Private Const CONNECTION_STRING = "Provider=SQLOLEDB;Data Source=prod-sqlmserver-eastus.public.22fe579d4c3a.database.windows.net,3342;Initial Catalog=FPProd;User ID=fpreadonly;Password=2lNZUHYehwXv4bai5Mo60ok8;Encrypt=yes;TrustServerCertificate=no;"

' Global variables for database connectivity
Public conn As Object
Public rs As Object

' ===== FORM LAUNCHING PATTERNS =====

Public Sub LaunchFundSelectorForm()
    ' DIRECT FORM REFERENCE ONLY: The build creates forms BEFORE importing modules
    ' Never reference placeholder names like UserForm1 in modules
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Launching Fund Selector Form ==="
    Debug.Print "Attempting to launch: frmFundSelector"
    
    ' Initialize database connection first
    If InitializeDatabase() Then
        frmFundSelector.Show
        Debug.Print "✅ Successfully launched form: frmFundSelector"
    Else
        MsgBox "Could not connect to database. Please check your connection settings.", vbCritical, "Database Connection Error"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "❌ Failed to launch frmFundSelector: " & Err.Description
    MsgBox "Could not launch frmFundSelector. Ensure the build completed successfully and the form was created.", _
          vbCritical, "Form Launch Error"
End Sub

' ===== DATABASE CONNECTION FUNCTIONS =====

Private Function InitializeDatabase() As Boolean
    ' Initialize database connection with better error handling
    On Error GoTo ErrorHandler
    
    ' Check if already connected
    If Not conn Is Nothing Then
        If conn.State = 1 Then
            InitializeDatabase = True
            Debug.Print "✅ Database already connected"
            Exit Function
        End If
    End If
    
    Debug.Print "=== Initializing Database Connection ==="
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = CONNECTION_STRING
    conn.CommandTimeout = 30
    conn.Open
    
    InitializeDatabase = True
    Debug.Print "✅ Database connection established successfully"
    Exit Function
    
ErrorHandler:
    InitializeDatabase = False
    Debug.Print "❌ Database connection failed: " & Err.Description
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Function

Public Sub TestDatabaseConnection()
    ' Test database connection and show result
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Testing Database Connection ==="
    
    If InitializeDatabase() Then
        MsgBox "Database connection successful!", vbInformation, "Connection Test"
        Debug.Print "✅ Database connection test passed"
    Else
        MsgBox "Database connection failed. Please check your connection settings.", vbCritical, "Connection Test"
        Debug.Print "❌ Database connection test failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "❌ Error testing database connection: " & Err.Description
    MsgBox "Error testing database connection: " & Err.Description, vbCritical, "Connection Test Error"
End Sub

' ===== GENERAL PARTNERS DATA MANAGEMENT =====

Public Sub LoadGeneralPartnersIntoList(listBox As Object)
    ' Load general partners into the specified list box
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Loading General Partners into ListBox ==="
    
    ' Clear existing items
    listBox.Clear
    
    ' Get the list of general partners
    Dim gpList As Collection
    Set gpList = GetGeneralPartners(500) ' Get up to 500
    
    If gpList.Count > 0 Then
        Dim i As Integer
        Dim gpItem As Variant
        
        For i = 1 To gpList.Count
            gpItem = gpList(i)
            listBox.AddItem gpItem(1) ' GP Name
        Next i
        
        Debug.Print "✅ Loaded " & gpList.Count & " general partners"
    Else
        listBox.AddItem "No General Partners found."
        Debug.Print "⚠️ No general partners found"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "❌ Error loading general partners: " & Err.Description
    listBox.AddItem "Error loading data. Please try refreshing."
End Sub

Public Sub RefreshGeneralPartnersList()
    ' Refresh the general partners list in the form
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Refreshing General Partners List ==="
    
    ' Call the form's load function
    Call LoadGeneralPartnersIntoList(frmFundSelector.lstGeneralPartners)
    
    Debug.Print "✅ General partners list refreshed"
    Exit Sub
    
ErrorHandler:
    Debug.Print "❌ Error refreshing general partners list: " & Err.Description
    MsgBox "Error refreshing data: " & Err.Description, vbCritical, "Refresh Error"
End Sub

Private Function GetGeneralPartners(Optional limit As Integer = 100) As Collection
    ' Get a simple list of General Partners from database
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Querying General Partners from Database ==="
    
    ' Initialize collection
    Set GetGeneralPartners = New Collection
    
    ' Build simple SQL query to get top GPs
    Dim sql As String
    sql = "SELECT TOP " & limit & " GeneralPartnerID, GeneralPartner " & _
          "FROM tbl_GeneralPartners " & _
          "WHERE GeneralPartner IS NOT NULL " & _
          "ORDER BY GeneralPartner"
    
    ' Execute query
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1 ' adOpenKeyset, adLockReadOnly
    
    ' Build collection
    Do While Not rs.EOF
        GetGeneralPartners.Add Array(rs("GeneralPartnerID").value, rs("GeneralPartner").value)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    Debug.Print "✅ Retrieved " & GetGeneralPartners.Count & " general partners from database"
    Exit Function
    
ErrorHandler:
    Debug.Print "❌ Error querying general partners: " & Err.Description
    Set GetGeneralPartners = New Collection
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
End Function

' ===== LISTBOX INTERACTION HANDLERS =====

Public Sub HandleGeneralPartnerSelection(selectedIndex As Integer)
    ' Handle general partner selection from list box
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Handling General Partner Selection ==="
    Debug.Print "Selected index: " & selectedIndex
    
    If selectedIndex >= 0 Then
        ' Get the selected general partner name
        Dim selectedGP As String
        selectedGP = frmFundSelector.lstGeneralPartners.List(selectedIndex)
        
        Debug.Print "Selected General Partner: " & selectedGP
        
        ' Here you could add logic to show more details about the selected GP
        ' For example, load fund information, show details, etc.
        MsgBox "Selected: " & selectedGP, vbInformation, "General Partner Selected"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "❌ Error handling general partner selection: " & Err.Description
End Sub

Public Sub HandleGeneralPartnerDoubleClick(selectedIndex As Integer)
    ' Handle general partner double-click from list box
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Handling General Partner Double-Click ==="
    Debug.Print "Double-clicked index: " & selectedIndex
    
    If selectedIndex >= 0 Then
        ' Get the selected general partner name
        Dim selectedGP As String
        selectedGP = frmFundSelector.lstGeneralPartners.List(selectedIndex)
        
        Debug.Print "Double-clicked General Partner: " & selectedGP
        
        ' Here you could add logic to open a detailed view or perform an action
        ' For example, open a fund details form, launch a report, etc.
        MsgBox "Double-clicked: " & selectedGP & vbCrLf & "This could open detailed fund information.", vbInformation, "General Partner Details"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "❌ Error handling general partner double-click: " & Err.Description
End Sub

' ===== UTILITY FUNCTIONS =====

Public Sub ShowSystemInfo()
    ' Demonstrates system information gathering
    Dim info As String
    info = "Fund Selector Application" & vbCrLf & vbCrLf
    info = info & "Application: FundSelector" & vbCrLf
    info = info & "Version: 1.0.0" & vbCrLf
    info = info & "Host: " & Application.Name & vbCrLf
    info = info & "VBA Version: " & Application.Version & vbCrLf
    info = info & "Database: Azure SQL (FPProd)"
    
    MsgBox info, vbInformation, "System Information"
End Sub

Public Sub DemonstrateDebugging()
    ' Shows debugging output patterns
    Debug.Print "=== Debug Output Demo ==="
    Debug.Print "This appears in VBA Immediate Window (Ctrl+G)"
    Debug.Print "Timestamp: " & Now()
    Debug.Print "User: " & Environ("USERNAME")
    Debug.Print "Computer: " & Environ("COMPUTERNAME")
    Debug.Print "Database Connection: " & IIf(Not conn Is Nothing And conn.State = 1, "Active", "Inactive")
End Sub

' ===== CLEANUP FUNCTIONS =====

Public Sub CloseDatabase()
    ' Clean up database connections
    On Error Resume Next
    
    Debug.Print "=== Closing Database Connections ==="
    
    ' Close database connections
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
        Debug.Print "✅ Recordset closed"
    End If
    
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
        Debug.Print "✅ Database connection closed"
    End If
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
' DATABASE PATTERNS:
' - Always check connection state before operations
' - Use proper error handling for database operations
' - Clean up connections in error handlers
' - Provide user feedback for connection issues
'
' AVOID:
' - Complex dynamic form discovery (unnecessary)
' - Collection-based form management
' - External dependencies when possible
' - Hardcoded paths or system-specific references
