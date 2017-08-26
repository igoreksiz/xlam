Private replacer As WorkbookLinkReplacer

Private Sub Workbook_Activate()
    Call AddUDFCategoryDescription
End Sub

Private Sub Workbook_Open()
    AppVersion = ThisWorkbook.Sheets("finboxio").range("AppVersion").value
    ReleaseDate = ThisWorkbook.Sheets("finboxio").range("ReleaseDate").value
    AppTitle = "finbox.io excel add-in v" & AppVersion & " (" & EXCEL_VERSION & " edition), released " & ReleaseDate

    LogMessage "Welcome to the " & AppTitle
    
    Call SetAPIKeyHandler(New APIKeyHandler)
    Call SetWorkbookLinkReplacer(New WorkbookLinkReplacer)
    
    Call UpdateCustomMenu
    
    CheckQuota blockEvents:=True
End Sub

Private Sub SetWorkbookLinkReplacer(wlr As WorkbookLinkReplacer)
    Set replacer = wlr
End Sub