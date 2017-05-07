Attribute VB_Name = "FinboxioMessagesModule"
' finbox.io API Integration

' Written by Michael Chambers, April 2017
' michael@mrchambers.f9.co.uk

' Upwork Contract Id 17916950

Option Explicit

Private CachedMessages As New Collection

Public Sub FinboxioMessages(Optional control As IRibbonControl)
    Dim msgs As Long
    msgs = CachedMessages.Count
    
    If msgs = 0 Then
        MsgBox "No finbox.io messages to display.", vbInformation, AppTitle
        Exit Sub
    End If

    Dim m As Long
    Dim LogWorksheet As Worksheet
    Dim LogWorkbook As Workbook
    
    Set LogWorkbook = Workbooks.Add(1)
      
    With LogWorkbook.Sheets(1)
        .Columns("A:A").ColumnWidth = 120
        .Columns("A:A").WrapText = True
        
        For m = msgs To 1 Step -1
            .Cells(msgs - m + 1, "A").value = CachedMessages(m)
        Next m
        .name = "finbox.io messages"
    End With

    Set LogWorksheet = Nothing
End Sub

Public Sub LogMessage(ByVal Msg As String, Optional ByVal key As String = "")
    If key <> "" Then Msg = Msg & " (" & key & ")"
    CachedMessages.Add Now() & "  " & Msg
End Sub
