Attribute VB_Name = "FinboxioRefreshModule"
Option Explicit

Public Sub RefreshData()
    DisplayDataLimit
    FixAddinLinks
    ClearCache
    Application.CalculateFull
End Sub

Public Sub DisplayDataLimit()
    Dim ack As Integer
    ack = MsgBox("You have exhausted your finbox.io data limit. Please try again later.", vbCritical)
    RedisplayDataLimit = Now() + (5 / (60 * 24))
End Sub
