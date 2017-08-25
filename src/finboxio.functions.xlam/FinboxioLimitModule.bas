Attribute VB_Name = "FinboxioLimitModule"
Option Explicit
Option Private Module

Private RedisplayWarning As Date

Public Sub ShowRateLimitWarning(Optional reset As Boolean = True)
    Dim ack As Integer
    ack = MsgBox("You have exhausted your finbox.io data limit. Click OK to view your usage history.", vbCritical + vbOKCancel)
    If reset Then SetRateLimitTimer
    If ack = vbOK Then ThisWorkbook.FollowHyperlink USAGE_URL
End Sub

Public Function IsRateLimited()
    If RedisplayWarning > Now() Then
        IsRateLimited = True
    Else
        IsRateLimited = False
    End If
End Function

Public Sub ClearRateLimit()
    RedisplayWarning = Now() - 1
End Sub

Private Sub SetRateLimitTimer()
    RedisplayWarning = Now() + (5 / (60 * 24))
End Sub

