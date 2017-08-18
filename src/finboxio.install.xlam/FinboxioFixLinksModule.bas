Attribute VB_Name = "FinboxioFixLinksModule"
Option Explicit
Option Private Module

Public IsReplacingLinks As Boolean

Public Function FixAddinLinks(Optional wb As Workbook)
    On Error GoTo CleanExit
    
    IsReplacingLinks = True
    
    Dim calc As Long
    Dim prefix As String
    Dim sheet As Worksheet
    Dim replaced As Boolean
    
    #If Mac Then
        prefix = "file:///*"
    #Else
        prefix = "?:\*"
    #End If
    
    replaced = False
    
    Dim ws
    If TypeName(wb) = "Empty" Or wb Is Nothing Then
        Set ws = Worksheets
    Else
        Set ws = wb.Worksheets
    End If
    
    For Each sheet In ws
        If Not sheet.Cells.Find("'" & prefix & "finboxio.install.xlam'!", , xlFormulas, xlPart, xlByRows, , False) Is Nothing And Not sheet.ProtectionMode Then
            sheet.Cells.Replace _
                What:="'" & prefix & "finboxio.install.xlam'!", _
                Replacement:="", _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                MatchCase:=False
            replaced = True
        End If
        
        If Not sheet.Cells.Find("'" & prefix & "finboxio.xlam'!", , xlFormulas, xlPart, xlByRows, , False) Is Nothing And Not sheet.ProtectionMode Then
            sheet.Cells.Replace _
                What:="'" & prefix & "finboxio.xlam'!", _
                Replacement:="", _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                MatchCase:=False
            replaced = True
        End If
    Next sheet

CleanExit:
    ResetFindReplace
    IsReplacingLinks = False
    If replaced Then Application.CalculateFull
    If replaced Then Application.CalculateFull
End Function



