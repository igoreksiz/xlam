Attribute VB_Name = "FinboxioRefreshModule"
Option Explicit

Public Sub RefreshData()
    FixAddinLinks
    ClearCache
    Application.CalculateFull
End Sub
