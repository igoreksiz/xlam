Private handler As AppEventHandler

Private Sub Workbook_Activate()
    Call AddUDFCategoryDescription
End Sub

Private Sub Workbook_Open()
    LogMessage "Welcome to the finbox.io excel add-in!"
    LogMessage "Running add-in version: " & GetAPIHeader
    
    Set handler = New AppEventHandler
    
    Call SetAPIKeyHandler(New APIKeyHandler)
    Call UpdateCustomMenu
    
    CheckQuota blockEvents:=True
End Sub
