Attribute VB_Name = "LoaderModule"
Option Explicit
Option Private Module

Public loadingManager As Boolean
Public updatingFunctions As Boolean

Public Function IsLoadingManager() As Boolean
    IsLoadingManager = loadingManager
End Function

Public Function IsUpdatingFunctions() As Boolean
    IsUpdatingFunctions = updatingFunctions
End Function

' Check if the functions add-in is installed alongside
Public Function HasAddInFunctions() As Boolean
    HasAddInFunctions = _
        SafeDir(LocalPath(AddInFunctionsFile)) <> "" Or _
        SafeDir(LocalPath(AddInFunctionsFile), vbHidden) <> ""
End Function

' Load the functions add-in installed alongside
Public Sub LoadAddInFunctions()

    ' Make sure add-in is installed on Mac 2011
    If ExcelVersion = "Mac2011" Then
        Dim addin As addin, installed As addin
        For Each addin In Application.AddIns
            If addin.name = AddInFunctionsFile Then
                Set installed = addin
                Exit For
            End If
        Next addin
        
        If addin Is Nothing Then
            Set installed = Application.AddIns.Add(LocalPath(AddInFunctionsFile), True)
        End If
        
        installed.installed = True
    End If

    ' If the functions add-in is already loaded,
    ' we should just exit.
    If LoadedAddInFunctions Then Exit Sub
    
    Dim appSec As MsoAutomationSecurity
    appSec = Application.AutomationSecurity
    
    ' If an update is staged, promote it to the active
    ' add-in. Only do this if this is an installed add-in
    ' so that we don't accidentally overwrite a dev
    ' version of the functions add-in.
    If HasStagedUpdate And AddInInstalled Then
        PromoteStagedUpdate
    End If
    
    On Error GoTo RemoveAddInFunctions
    
    If ExcelVersion = "Mac2011" Then
        Call Workbooks.Open(LocalPath(AddInFunctionsFile))
    Else
        Application.AutomationSecurity = msoAutomationSecurityLow
        ' Load the functions add-in
        Call Workbooks.Open(LocalPath(AddInFunctionsFile))
        Application.AutomationSecurity = appSec
    End If
    
    Exit Sub

RemoveAddInFunctions:
    ' If for some reason we can't open the functions
    ' component, the workbook may be corrupted.
    ' Just remove all traces so it will be re-downloaded
    ' on the next restart.
    
    Application.AutomationSecurity = appSec
    
    RemoveAddInFunctions
    
    MsgBox _
        Title:="[finbox.io] Add-in Error", _
        Prompt:="The finbox.io add-in functions were not loaded correctly. " & _
                "Please try restarting Excel and contact support@finbox.io if this problem persists.", _
        Buttons:=vbCritical
End Sub

' Ensures that functions add-in is uninstalled and unloaded
Public Function UninstallAddInFunctions() As Boolean
    Dim addin As addin
    For Each addin In Application.AddIns
        If addin.name = AddInFunctionsFile And addin.installed Then
            addin.installed = False
            UninstallAddInFunctions = True
            Exit Function
        End If
    Next addin
End Function

' Checks if the functions add-in is currently loaded
Public Function LoadedAddInFunctions() As Boolean
    ' The add-in may be loaded as a hidden file, so
    ' it won't always show up in the add-ins list.
    ' So the safest thing to do is check if the workbook
    ' itself is open. If the call below succeeds, then
    ' we know it's loaded.
    On Error GoTo Finish
    Dim name As String
    name = Workbooks(AddInFunctionsFile).name
    LoadedAddInFunctions = True
Finish:

End Function

' Unloads the currently loaded functions add-in.
' Does nothing if the add-in is not loaded.
Public Function UnloadAddInFunctions() As Boolean
    Dim openName As String, canUnloadFunctions As Boolean

    ' If the workbook isn't open, this will fail
    On Error GoTo Unloaded
    openName = Workbooks(AddInFunctionsFile).name

    ' If the functions module is in the process of
    ' updating this add-in, we shouldn't unload it
    canUnloadFunctions = _
        Not Application.Run(AddInFunctionsFile & "!IsUpdatingManager") And _
        Not Application.Run(AddInFunctionsFile & "!IsCheckingUpdates")
    If Not canUnloadFunctions Then Exit Function

    ' Try to close workbook. If either of these
    ' calls fail it likely means the workbook is
    ' closed.
    Workbooks(AddInFunctionsFile).Close
    openName = Workbooks(AddInFunctionsFile).name
    
    ' Workbook must still be open
    Exit Function
    
Unloaded:
    ' Workbook is not loaded
    UnloadAddInFunctions = True
End Function

' Check if staged functions add-in is available
Private Function HasStagedUpdate() As Boolean
    HasStagedUpdate = _
        SafeDir(StagingPath(AddInFunctionsFile)) <> "" Or _
        SafeDir(StagingPath(AddInFunctionsFile), vbHidden) <> ""
End Function

' Promotes the staged functions add-in to active
Public Sub PromoteStagedUpdate()
    If updatingFunctions Or Not HasStagedUpdate Then Exit Sub
    
    On Error GoTo Finish
    updatingFunctions = True
    If UnloadAddInFunctions Then
        If HasAddInFunctions Then
            SetAttr LocalPath(AddInFunctionsFile), vbNormal
            Kill LocalPath(AddInFunctionsFile)
        End If
        Name StagingPath(AddInFunctionsFile) As LocalPath(AddInFunctionsFile)
        VBA.SetAttr LocalPath(AddInFunctionsFile), vbHidden
        
        #If Mac Then
            MsgBox _
                Title:="[finbox.io] Add-In Functions Updated", _
                Prompt:="A new version of the add-in functions have been installed. " & _
                        "You may be prompted to enable the updated macros. " & _
                        "Macros must be enabled or the add-in will not function properly."
        #End If
        
        LoadAddInFunctions
    End If
    
Finish:
    updatingFunctions = False
End Sub

