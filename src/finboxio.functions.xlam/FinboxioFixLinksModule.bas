Attribute VB_Name = "FinboxioFixLinksModule"
Option Explicit
Option Private Module

Public Const FNBXFND = "!FNBX("

Public IsReplacingLinks As Boolean

Public Function FixAddInLinks(Optional wb As Workbook)
    On Error GoTo CleanExit
    Dim sheet As Worksheet, replaced As Boolean, rng As Range
    
    IsReplacingLinks = True
    Application.ScreenUpdating = False
    
    replaced = False
    
    Dim ws
    If TypeName(wb) = "Empty" Or wb Is Nothing Then
        Set ws = Worksheets
    Else
        Set ws = wb.Worksheets
    End If

    For Each sheet In ws
        Set rng = sheet.UsedRange
        Dim FirstFound As String, LastCell As Range, FoundCell As Range
        Set LastCell = rng.Cells(rng.Cells.count)
        Set FoundCell = rng.Find(What:=FNBXFND, LookIn:=xlFormulas, LookAt:=xlPart, After:=LastCell, MatchCase:=False)
        If Not FoundCell Is Nothing Then
            FirstFound = FoundCell.address
            On Error Resume Next
            Do Until FoundCell Is Nothing
                Set FoundCell = rng.Find(What:=FNBXFND, LookIn:=xlFormulas, LookAt:=xlPart, After:=FoundCell, MatchCase:=False)
                If FoundCell.HasFormula Then
                    FoundCell.Formula = DereferenceFNBX(FoundCell.Formula)
                    replaced = True
                End If
                If FoundCell.address = FirstFound Then Exit Do
            Loop
        End If
    Next sheet
    
CleanExit:
    ResetFindReplace
    Application.ScreenUpdating = True
    IsReplacingLinks = False
    If replaced Then Application.CalculateFull
End Function

Function DereferenceFNBX(ByVal Formula As String)
    DereferenceFNBX = Formula
    Dim replaced As String
    replaced = Formula
    On Error GoTo Finish
    While VBA.InStr(replaced, FNBXFND) > 0
        Dim i As Integer: i = VBA.InStr(replaced, FNBXFND)
        Dim p As String: p = VBA.Mid(replaced, i - 1, 1)
        If p = "'" Then
            replaced = _
                VBA.Left(replaced, VBA.InStrRev(replaced, "'", i - 2) - 1) & _
                VBA.Mid(replaced, i + 1)
        Else
            replaced = _
                VBA.Left(replaced, VBA.InStrRev(replaced, "finboxio", i - 1) - 1) & _
                VBA.Mid(replaced, i + 1)
        End If
    Wend
    DereferenceFNBX = replaced
Finish:
End Function
