Attribute VB_Name = "FinboxioUpdateModule"
Option Explicit
Option Private Module

Private updatingManager As Boolean

Public Function IsUpdatingManager() As Boolean
    IsUpdatingManager = updatingManager
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
    
    On Error Resume Next
    Dim loadingManager As Boolean
    loadingManager = Application.Run(AddInManagerFile & "!IsLoadingManager")
    If loadingManager Then Exit Sub
    
    On Error GoTo Finish
    
    updatingManager = True
    
    ' Uninstall the active manager
    Dim addIn As addIn
    For Each addIn In Application.AddIns
        If addIn.name = AddInInstalledFile And addIn.Installed Then
            addIn.Installed = False
            Exit For
        End If
    Next addIn
    
    ' Promote staged manager
    If HasInstalledAddInManager Then
        SetAttr LocalPath(AddInInstalledFile), vbNormal
        Kill LocalPath(AddInInstalledFile)
    End If
    Name StagingPath(AddInInstalledFile) As LocalPath(AddInInstalledFile)
    VBA.SetAttr LocalPath(AddInInstalledFile), vbNormal
    
    ' Reinstall the manager
    If Not addIn Is Nothing Then addIn.Installed = True

Finish:
    updatingManager = False
End Sub
