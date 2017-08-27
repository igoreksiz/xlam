Private Sub Workbook_Open()
    loadingManager = True
    
    ' Install AddIn on first open
    AddInInstalled = InstallAddIn(Me)
    
    ' Download newer version if available
    If AddInInstalled And Not HasUpdates Then
        Call DownloadUpdates(blockEvents:=True)
    End If
    
    ' Unload the current functions add-in
    ' if an update is available
    If HasUpdates Then
        Call UnloadAddInFunctions
    End If
    
    ' Load the FNBX functions
    Call LoadAddInFunctions
    
    loadingManager = False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' The functions add-in should NOT be formally installed
    ' (as indicated by a checkmark in the add-ins window),
    ' since it will be loaded by the manager add-in.
    ' If it is installed, we must uninstall it to
    ' avoid problems on startup.
    ' Theoretically it should be very unlikely that the
    ' functions add-in actually ends up in an installed
    ' state since it is a hidden file and won't show up
    ' in the add-ins list, but it doesn't hurt to check.
    Call UninstallAddInFunctions
End Sub