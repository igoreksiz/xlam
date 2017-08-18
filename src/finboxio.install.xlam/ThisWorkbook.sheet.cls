Private replacer As WorkbookLinkReplacer

Private Sub Workbook_Activate()
    Call AddUDFCategoryDescription
End Sub

Private Sub Workbook_Open()
    Call InstallAddin(Me)
    
    Call SetAPIKeyHandler(New APIKeyHandler)
    Call SetWorkbookLinkReplacer(New WorkbookLinkReplacer)
    
    Call AddCustomMenu

    AppVersion = ThisWorkbook.Sheets("finboxio").range("AppVersion")
    ReleaseDate = ThisWorkbook.Sheets("finboxio").range("ReleaseDate")
    AppTitle = "finbox.io excel add-in v" & AppVersion & " (" & EXCEL_VERSION & " edition), released " & ReleaseDate

    LogMessage "Welcome to the " & AppTitle
    
    CheckQuota
End Sub

Private Sub SetWorkbookLinkReplacer(wlr As WorkbookLinkReplacer)
    Set replacer = wlr
End Sub