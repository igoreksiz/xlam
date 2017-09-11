Attribute VB_Name = "FinboxioLogModule"
Option Explicit
Option Private Module

Public trigger As String

Public Sub ShowMessages()
    Application.Run (AddInManagerFile & "!TrimLog")
    ThisWorkbook.FollowHyperlink address:=LocalPath(AddInLogFile), AddHistory:=False
End Sub

Public Sub LogMessage(ByVal msg As String, Optional ByVal key As String = "")
    Dim source As String
    source = "v" & AddInVersion & " " & ThisWorkbook.name & " -"
    If VBA.Len(source) < 40 Then source = source & String(40 - VBA.Len(source), "-")
    If trigger <> "" Then msg = "(" & trigger & ") -> " & msg
    If key <> "" Then msg = "(" & key & ") -> " & msg
    msg = "[" & VBA.Format(VBA.Now(), "yyyy-MM-dd hh:mm:ss") & "] " & source & "- " & msg
    Debug.Print msg
    Open LocalPath(AddInLogFile) For Append As #1
        Print #1, msg
    Close #1
End Sub