Attribute VB_Name = "FinboxioRefreshModule"
Option Explicit

Public Sub RefreshData()
    ' If we already know data limit has been reached,
    ' just display notification
    If TypeName(RedisplayDataLimit) = "Date" Then
        If RedisplayDataLimit > Now() Then
            DisplayDataLimit False
        End If
    End If
    
    FixAddinLinks
    ClearCache
    Application.CalculateFull
End Sub

Public Sub DisplayDataLimit(Optional reset As Boolean = True)
    Dim ack As Integer
    ack = MsgBox("You have exhausted your finbox.io data limit. Click OK to view your usage history.", vbCritical + vbOKCancel)
    If reset Then
        RedisplayDataLimit = Now() + (5 / (60 * 24))
    End If
    If ack = vbOK Then
        ActiveWorkbook.FollowHyperlink USAGE_URL
    End If
End Sub
