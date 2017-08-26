Private Sub Workbook_Open()
    ' Load AddIn Details
    AddInVersion = ThisWorkbook.Sheets("finboxio").Range("AppVersion").value
    AddInDate = ThisWorkbook.Sheets("finboxio").Range("ReleaseDate").value
    AddInDetail = "finbox.io excel add-in v" & AddInVersion & " (" & EXCEL_VERSION & " edition), released " & AddInReleaseDate
    
    ' Install AddIn on first open
    AddInInstalled = InstallAddIn(Me)
    
    ' Download newer version if available
    If AddInInstalled Then
        Call DownloadUpdates
        If HasStagedUpdates(AddInFunctionsFile) Then
            PromoteStagedUpdate AddInFunctionsFile
        End If
    End If
    
    ' Load the FNBX functions
    Call LoadAddInFunctions
End Sub