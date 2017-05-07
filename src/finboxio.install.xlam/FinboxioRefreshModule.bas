Attribute VB_Name = "FinboxioRefreshModule"
Option Explicit

Public Sub FinboxioRefresh(Optional control As IRibbonControl)
    ClearCache
    Application.CalculateFull
End Sub
