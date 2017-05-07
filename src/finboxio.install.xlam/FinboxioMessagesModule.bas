Attribute VB_Name = "FinboxioMessagesModule"
' finbox.io API Integration

Option Explicit

Private CachedMessages As New Collection

Public Function TestMessages(count As Long)
    Dim c As Long
    For c = 1 To count
        LogMessage ("Test message")
    Next c
End Function

Public Sub ShowMessages()
    Dim msgs As Long
    msgs = CachedMessages.count
    
    If msgs = 0 Then
        MsgBox "No finbox.io messages to display.", vbInformation, AppTitle
        Exit Sub
    End If

    Dim r() As String
    Dim m As Long
    ReDim r(1 To msgs, 1 To 1)
    For m = msgs To 1 Step -1
        r(m, 1) = CachedMessages(m)
    Next m

    Dim LogWorksheet As Worksheet
    Dim LogWorkbook As Workbook
    
    Set LogWorkbook = Workbooks.Add(1)
      
    With LogWorkbook.Sheets(1)
        .Columns("A:A").ColumnWidth = 120
        .Columns("A:A").WrapText = True
        .Range(.Cells(1, 1), .Cells(msgs, 1)).value = r
        .name = "finbox.io messages"
    End With

    Set LogWorksheet = Nothing
End Sub

Public Sub LogMessage(ByVal Msg As String, Optional ByVal key As String = "")
    If key <> "" Then Msg = Msg & " (" & key & ")"
    CachedMessages.Add Now() & "  " & Msg
End Sub

