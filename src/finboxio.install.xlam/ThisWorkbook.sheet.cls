Private Sub Workbook_Activate()
    Call AddUDFCategoryDescription
End Sub

Private Sub Workbook_Open()
    Call InstallAddin(Me)
    
    Call SetAPIKeyHandler(New APIKeyHandler)
    
    Call AddCustomMenu
    Call CheckUpdates

    LogMessage "Welcome to " & AppTitle
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Call DeleteCustomMenu
End Sub