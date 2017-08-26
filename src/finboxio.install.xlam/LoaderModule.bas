Attribute VB_Name = "LoaderModule"
Option Explicit
Option Private Module

Public Function HasAddInFunctions() As Boolean
    HasAddInFunctions = Dir(XlamPath(AddInFunctionsFile)) <> ""
End Function

Public Sub LoadAddInFunctions()
    If Not AddInInstalled And LoadedAddInFunctions Then UnloadAddInFunctions
    Call Workbooks.Open(XlamPath(AddInFunctionsFile))
End Sub

Private Function LoadedAddInFunctions() As Boolean
    LoadedAddInFunctions = False
    Dim addIn As addIn
    For Each addIn In Application.AddIns
        If addIn.name = XlamFile(AddInFunctionsFile) Then LoadedAddInFunctions = True
    Next addIn
End Function

Private Sub UnloadAddInFunctions()
    Workbooks(XlamFile(AddInFunctionsFile)).Close
End Sub
