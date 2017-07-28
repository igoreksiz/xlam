Attribute VB_Name = "FinboxioFixLinksModule"
Option Explicit

Public IsReplacingLinks As Boolean

Public Function FixAddinLinks(Optional wb As Workbook)
    On Error GoTo CleanExit
    
    IsReplacingLinks = True
    
    Dim calc As Long
    Dim Sheet As Worksheet
    Dim replaced As Boolean
    
    replaced = False
    
    Dim ws
    If TypeName(wb) = "Empty" Or wb Is Nothing Then
        Set ws = Worksheets
    Else
        Set ws = wb.Worksheets
    End If
    
    Application.ScreenUpdating = False
    For Each Sheet In ws
        If Not Sheet.Cells.Find("'*finboxio.install.xlam'!", , xlFormulas, xlPart, xlByRows, , False) Is Nothing And Not Sheet.ProtectionMode Then
            Sheet.Cells.Replace _
                What:="'*finboxio.install.xlam'!", _
                Replacement:="", _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                MatchCase:=False
            replaced = True
        End If
        
        If Not Sheet.Cells.Find("'*finboxio.xlam'!", , xlFormulas, xlPart, xlByRows, , False) Is Nothing And Not Sheet.ProtectionMode Then
            Sheet.Cells.Replace _
                What:="'*finboxio.xlam'!", _
                Replacement:="", _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                MatchCase:=False
            replaced = True
        End If
    Next Sheet

CleanExit:
    ResetFindReplace
    Application.ScreenUpdating = True
    IsReplacingLinks = False
    If replaced Then Application.CalculateFull
End Function

