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
    If Not HasStagedUpdate Then Exit Sub
    
    On Error GoTo NoManager

    Dim openName As String, canUnloadManager As Boolean
    openName = Workbooks(AddInManagerFile).name
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
        If addIn.name = AddInInstalledFile Then
            On Error GoTo PromoteStaged
            Workbooks(AddInInstalledFile).Close
            On Error GoTo ReportError
            Exit For
        End If
    Next addIn

PromoteStaged:
    ' Promote staged manager
    If HasInstalledAddInManager Then
        SetAttr LocalPath(AddInInstalledFile), vbNormal
        Kill LocalPath(AddInInstalledFile)
    End If
    Name StagingPath(AddInInstalledFile) As LocalPath(AddInInstalledFile)
    VBA.SetAttr LocalPath(AddInInstalledFile), vbNormal
    
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


