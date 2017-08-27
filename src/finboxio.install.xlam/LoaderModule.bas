Attribute VB_Name = "LoaderModule"
Option Explicit
Option Private Module

Public loadingManager As Boolean

Public Function IsLoadingManager() As Boolean
    IsLoadingManager = loadingManager
End Function

' Check if the functions add-in is installed alongside
Public Function HasAddInFunctions() As Boolean
    HasAddInFunctions = _
        Dir(LocalPath(AddInFunctionsFile)) <> "" Or _
        Dir(LocalPath(AddInFunctionsFile), vbHidden) <> ""
End Function

' Load the functions add-in installed alongside
Public Sub LoadAddInFunctions()
    ' If the functions add-in is already loaded,
    ' we should just exit.
    If LoadedAddInFunctions Then Exit Sub
    
    ' If an update is staged, promote it to the active
    ' add-in. Only do this if this is an installed add-in
    ' so that we don't accidentally overwrite a dev
    ' version of the functions add-in.
    If HasStagedUpdate And AddInInstalled Then
        PromoteStagedUpdate
    End If
    
    ' Load the functions add-in
    Call Workbooks.Open(LocalPath(AddInFunctionsFile))
End Sub

' Ensures that functions add-in is uninstalled and unloaded
Public Function UninstallAddInFunctions() As Boolean
    Dim addIn As addIn
    For Each addIn In Application.AddIns
        If addIn.name = AddInFunctionsFile And addIn.installed Then
            addIn.installed = False
            UninstallAddInFunctions = True
            Exit Function
        End If
    Next addIn
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
Public Sub UnloadAddInFunctions()
    On Error Resume Next
    
    ' If the functions module is in the process of
    ' updating this add-in, we shouldn't unload it
    Dim midUpdate As Boolean
    midUpdate = Application.Run(AddInFunctionsFile & "!IsUpdatingManager")
    If Not midUpdate Then Workbooks(AddInFunctionsFile).Close
End Sub

' Check if staged functions add-in is available
Private Function HasStagedUpdate() As Boolean
    HasStagedUpdate = _
        Dir(StagingPath(AddInFunctionsFile)) <> "" Or _
        Dir(StagingPath(AddInFunctionsFile), vbHidden) <> ""
End Function

' Promotes the staged functions add-in to active
Private Sub PromoteStagedUpdate()
    If HasAddInFunctions Then
        SetAttr LocalPath(AddInFunctionsFile), vbNormal
        Kill LocalPath(AddInFunctionsFile)
    End If
    Name StagingPath(AddInFunctionsFile) As LocalPath(AddInFunctionsFile)
    VBA.SetAttr LocalPath(AddInFunctionsFile), vbHidden
End Sub
