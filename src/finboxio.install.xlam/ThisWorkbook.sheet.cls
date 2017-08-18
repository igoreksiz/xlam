Private replacer As WorkbookLinkReplacer

Private Sub Workbook_Activate()
    Call AddUDFCategoryDescription
End Sub

Private Sub Workbook_Open()
    Call InstallAddin(Me)
    
    Call SetAPIKeyHandler(New APIKeyHandler)
    Call SetWorkbookLinkReplacer(New WorkbookLinkReplacer)
    
    Call AddCustomMenu

    LogMessage "Welcome to " & AppTitle
    
    CheckQuota
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Call DeleteCustomMenu
End Sub

Private Sub SetWorkbookLinkReplacer(wlr As WorkbookLinkReplacer)
    Set replacer = wlr
End Sub
