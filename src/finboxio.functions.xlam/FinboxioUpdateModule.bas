Attribute VB_Name = "FinboxioUpdateModule"
Option Explicit
Option Private Module

Public updatingManager As Boolean
Public checkingUpdates As Boolean

Public Function IsUpdatingManager() As Boolean
    IsUpdatingManager = updatingManager
End Function

Public Function IsCheckingUpdates() As Boolean
    IsCheckingUpdates = checkingUpdates
End Function

Public Function HasInstalledAddInManager() As Boolean
    HasInstalledAddInManager = _
        Dir(LocalPath(AddInInstalledFile)) <> "" Or _
        Dir(LocalPath(AddInInstalledFile), vbHidden) <> ""
End Function

Public Function HasStagedUpdate() As Boolean
    HasStagedUpdate = _
        Dir(StagingPath(AddInInstalledFile)) <> "" Or _
        Dir(StagingPath(AddInInstalledFile), vbHidden) <> ""
End Function

' Promotes the staged add-in manager to active
Public Sub PromoteStagedUpdate()
    If updatingManager Or Not HasStagedUpdate Then Exit Sub

    ' Test open the workbook to guarantee macros are
    ' available before trying to run them
    On Error GoTo NoManager
    Dim openName As String, canUnloadManager As Boolean
    openName = Workbooks(AddInManagerFile).name

    ' Make sure manager isn't doing something that would
    ' prevent us from unloading it properly
    canUnloadManager = _
        Not Application.Run(AddInManagerFile & "!IsLoadingManager") And _
        Not Application.Run(AddInManagerFile & "!IsUpdatingFunctions")
        
    If Not canUnloadManager Then Exit Sub
    
NoManager:
    Dim appSec As MsoAutomationSecurity
    appSec = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityLow

    On Error GoTo ReportError

    updatingManager = True
   
    ' Uninstall the active manager
    Dim addIn As addIn
    For Each addIn In Application.AddIns
        If addIn.name = AddInInstalledFile And addIn.Installed Then
            addIn.Installed = False
            Exit For
        End If
    Next addIn
    
    ' Ensure the manager is unloaded
    UnloadAddInManager

    ' Promote staged manager
    If HasInstalledAddInManager Then
        SetAttr LocalPath(AddInInstalledFile), vbNormal
        Kill LocalPath(AddInInstalledFile)
    End If
    Name StagingPath(AddInInstalledFile) As LocalPath(AddInInstalledFile)
    VBA.SetAttr LocalPath(AddInInstalledFile), vbNormal
    
    #If Mac Then
        MsgBox _
            Title:="[finbox.io] Add-In Manager Updated", _
            Prompt:="A new version of the add-in manager has been installed. " & _
                    "You may be prompted to enable the updated macros. " & _
                    "Macros must be enabled or the add-in will not function properly."
    #End If
    
    ' Reinstall the manager
    If Not addIn Is Nothing Then
        addIn.Installed = True
    Else
        Set addIn = Application.AddIns.Add(LocalPath(AddInInstalledFile), True)
    End If
    
    ' Ensure the manager workbook is opened
    Call Workbooks.Open(LocalPath(AddInInstalledFile))
    
    GoTo Finish

ReportError:

    MsgBox _
        Title:="[finbox.io] Add-in Error", _
        Prompt:="The finbox.io add-in manager was not loaded correctly. " & _
                "Please try restarting Excel and contact support@finbox.io if this problem persists.", _
        Buttons:=vbCritical

Finish:
    updatingManager = False
    Application.AutomationSecurity = appSec
End Sub

' Unloads the currently loaded add-in manager.
' Does nothing if the add-in is not loaded.
Private Function UnloadAddInManager() As Boolean
    Dim openName As String

    ' If the workbook isn't open, this will fail
    On Error GoTo Unloaded
    openName = Workbooks(AddInInstalledFile).name

    ' Try to close workbook. If either of these
    ' calls fail it likely means the workbook is
    ' closed.
    Workbooks(AddInFunctionsFile).Close
    openName = Workbooks(AddInFunctionsFile).name
    
    ' Workbook must still be open
    Exit Function
    
Unloaded:
    ' Workbook is not loaded
    UnloadAddInManager = True
End Function
