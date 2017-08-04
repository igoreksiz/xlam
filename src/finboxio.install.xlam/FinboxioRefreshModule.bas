Attribute VB_Name = "FinboxioRefreshModule"
Option Explicit

Public Sub RefreshData()
    On Error GoTo EnableCache
    
    StartRecache
    
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.Calculate
    Next
    
EnableCache:
    StopRecache
End Sub

