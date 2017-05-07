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
    ack = MsgBox("You have exhausted your finbox.io data limit. Click OK to view your usage history.", vbCritical + vbOKCancel)
    RedisplayDataLimit = Now() + (5 / (60 * 24))
    If ack = vbOK Then
        ActiveWorkbook.FollowHyperlink USAGE_URL
    End If
End Sub
