Private replacer As WorkbookLinkReplacer
Private isClosing As Boolean

Private Sub Workbook_Activate()
    Call AddUDFCategoryDescription
End Sub

Private Sub Workbook_Open()
    LogMessage "Welcome to the finbox.io excel add-in!"
    LogMessage "Running add-in version: " & GetAPIHeader
    
    Call SetAPIKeyHandler(New APIKeyHandler)
    Call SetWorkbookLinkReplacer(New WorkbookLinkReplacer)
    
    Call UpdateCustomMenu
    
    CheckQuota blockEvents:=True
End Sub

Private Sub SetWorkbookLinkReplacer(wlr As WorkbookLinkReplacer)
    Set replacer = wlr
End Sub
