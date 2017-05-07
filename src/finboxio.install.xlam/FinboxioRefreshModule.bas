Attribute VB_Name = "FinboxioRefreshModule"
Option Explicit

Public Sub RefreshData()
    ClearCache
    Application.CalculateFull
End Sub
