Attribute VB_Name = "LogModule"
Option Explicit
Option Private Module

Public trigger As String

Public Sub LogMessage(msg As String)
    Dim source As String
    source = "v" & AddInVersion & " " & ThisWorkbook.name & " -"
    If VBA.Len(source) < 40 Then source = source & String(40 - VBA.Len(source), "-")
    If trigger <> "" Then msg = "(" & trigger & ") -> " & msg
    msg = "[" & VBA.Format(VBA.Now(), "yyyy-MM-dd hh:mm:ss") & "] " & source & "- " & msg
    Debug.Print (msg)
    Open SavePath(AddInLogFile) For Append As #1
        Print #1, msg
    Close #1
End Sub

Public Sub TrimLog(Optional days As Integer = 0)
    If days = 0 Then days = GetSetting("logRetentionDays", 30)
    Dim line As String, timestamp As String, time As Date, trimmed As Integer
    trimmed = 0
    VBA.FileCopy SavePath(AddInLogFile), SavePath(AddInLogFile & ".tmp")
    Open SavePath(AddInLogFile) For Output As #1
    Open SavePath(AddInLogFile & ".tmp") For Input As #2
        While Not EOF(2)
            Line Input #2, line
            line = VBA.Trim(Application.Clean(line))
            timestamp = VBA.Mid(line, 2, VBA.InStr(line, "]") - 2)
            time = CDate(timestamp)
            If time > VBA.Now() - days Then
                Print #1, line
            Else
                trimmed = trimmed + 1
            End If
        Wend
    Close #2
    Close #1
    
    VBA.Kill SavePath(AddInLogFile & ".tmp")
    
    If trimmed > 0 Then
        LogMessage "Trimmed " & trimmed & " messages older than " & (VBA.Now() - days)
    End If
End Sub
