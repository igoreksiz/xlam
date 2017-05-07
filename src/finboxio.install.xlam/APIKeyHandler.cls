VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APIKeyHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' It is impossible to write other cell values when executing code
' triggered by a UDF call from another cell. So in order to support
' a login prompt on first use of the FNBX formula (if not logged in)
' this sets up a listener to save the collected API key in an independent
' execution context once the UDF is completed

Option Explicit

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private keyToStore As String

Private Sub Class_Initialize()
    Set App = Application
    keyToStore = ""
End Sub

Public Sub StoreAPIKey(key As String)
    keyToStore = key
    StoreKey
End Sub

Public Sub ClearAPIKey()
    keyToStore = ""
    StoreKey
End Sub

Public Function GetAPIKey()
    If keyToStore <> "" Then
        GetAPIKey = keyToStore
    Else
        GetAPIKey = VBA.Trim(ThisWorkbook.Sheets("API_Credentials").Range("APIkey").value)
    End If
End Function

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Source As Range)
    StoreKey
End Sub

Private Sub StoreKey()
    On Error GoTo EnableEvents
    App.EnableEvents = False
    ThisWorkbook.Sheets("API_Credentials").Range("APIkey").value = keyToStore
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    keyToStore = ""
    AppRibbon.Invalidate
    Exit Sub
EnableEvents:
    App.EnableEvents = True
End Sub
