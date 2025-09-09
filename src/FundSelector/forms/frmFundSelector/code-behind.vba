Option Explicit

Private Sub UserForm_Initialize()
    ' Load initial general partners
    LoadGeneralPartners
End Sub

Private Sub LoadGeneralPartners()
    Dim gpList As Collection
    Dim gpItem As Variant
    Dim i As Integer
    
    ' Get the list of general partners
    Set gpList = GetGeneralPartners(500) ' Get up to 500
    
    ' Clear and populate the list box
    lstGeneralPartners.Clear
    
    If gpList.Count > 0 Then
        For i = 1 To gpList.Count
            gpItem = gpList(i)
            lstGeneralPartners.AddItem gpItem(1) ' GP Name
        Next i
    Else
        lstGeneralPartners.AddItem "No General Partners found."
    End If
End Sub

Private Sub btnClose_Click()
    CloseDatabase
    Unload Me
End Sub 