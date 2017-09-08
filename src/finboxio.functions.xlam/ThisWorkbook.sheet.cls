Private handler As AppEventHandler

Private Sub Workbook_Activate()
    Call AddUDFCategoryDescription
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    On Error GoTo Finish
    If VBA.Dir(ThisWorkbook.path & Application.PathSeparator & ".git", vbDirectory Or vbHidden) <> "" Then
        Dim openName As String
        openName = Workbooks("vbaDeveloper.xlam").name
        Call Application.Run("vbaDeveloper.xlam!exportVbProject", "finboxioFunctions")
    End If
Finish:
End Sub

Private Sub Workbook_Open()
    LogMessage "Welcome to the finbox.io excel add-in!"
    LogMessage "Running add-in version: " & GetAPIHeader
    
    Set handler = New AppEventHandler
    
    Call SetAPIKeyHandler(New APIKeyHandler)
    Call UpdateCustomMenu
    
    CheckQuota blockEvents:=True
End Sub

