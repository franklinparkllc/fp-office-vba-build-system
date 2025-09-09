Attribute VB_Name = "modDBInterface"
' ===== FUND SELECTOR DATABASE INTERFACE (Minimal) =====
' Application-specific database functions for the fund selector

Option Explicit

' Azure SQL Database connection string
Private Const CONNECTION_STRING = "Provider=SQLOLEDB;Data Source=prod-sqlmserver-eastus.public.22fe579d4c3a.database.windows.net,3342;Initial Catalog=FPProd;User ID=fpreadonly;Password=2lNZUHYehwXv4bai5Mo60ok8;Encrypt=yes;TrustServerCertificate=no;"

' Global variables
Public conn As Object
Public rs As Object

' Main entry point - called by the generic build system
Public Sub OpenFundSelector()
    ' Initialize database connection
    If InitializeDatabase() Then
        ' Launch the form using the generic build system
        Call modAppBuilder.LaunchForm("frmFundSelector")
    Else
        MsgBox "Could not connect to database. Please check your connection settings.", vbCritical
    End If
End Sub

' Initialize database connection with better error handling
Private Function InitializeDatabase() As Boolean
    On Error GoTo ErrorHandler
    
    ' Check if already connected
    If Not conn Is Nothing Then
        If conn.State = 1 Then
            InitializeDatabase = True
            Exit Function
        End If
    End If
    
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = CONNECTION_STRING
    conn.CommandTimeout = 30
    conn.Open
    
    InitializeDatabase = True
    Exit Function
    
ErrorHandler:
    InitializeDatabase = False
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Function

' Get a simple list of General Partners
Public Function GetGeneralPartners(Optional limit As Integer = 100) As Collection
    Dim gpList As Collection
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    ' Initialize collection
    Set gpList = CreateCollection()
    
    ' Build simple SQL query to get top GPs
    sql = "SELECT TOP " & limit & " GeneralPartnerID, GeneralPartner " & _
          "FROM tbl_GeneralPartners " & _
          "WHERE GeneralPartner IS NOT NULL " & _
          "ORDER BY GeneralPartner"
    
    ' Execute query
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1 ' adOpenKeyset, adLockReadOnly
    
    ' Build collection
    Do While Not rs.EOF
        gpList.Add Array(rs("GeneralPartnerID").value, rs("GeneralPartner").value)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    Set GetGeneralPartners = gpList
    Exit Function
    
ErrorHandler:
    Set GetGeneralPartners = CreateCollection()
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
End Function

' Helper function
Private Function CreateCollection() As Collection
    Set CreateCollection = New Collection
End Function

' Cleanup
Public Sub CloseDatabase()
    On Error Resume Next
    
    ' Close database connections
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Sub 