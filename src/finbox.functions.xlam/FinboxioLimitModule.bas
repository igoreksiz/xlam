Attribute VB_Name = "FinboxioLimitModule"
Option Explicit
Option Private Module

Private RedisplayWarning As Date

Public Sub ShowRateLimitWarning(Optional reset As Boolean = True)
    MsgBox _
        Title:="[Finbox] Limit Exceeded", _
        Prompt:="You have exhausted your finbox data limit. Try again later or contact support@finbox.com to request a limit increase.", _
        Buttons:=vbCritical
    If reset Then SetRateLimitTimer
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

