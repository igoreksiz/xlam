VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set App = Application
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    If Not GetSetting("updateOnSheetActivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_SheetDeactivate(ByVal Sh As Object)
    If Not GetSetting("updateOnSheetDeactivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_AfterCalculate()
    If Not GetSetting("updateOnAfterCalculate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    If Not GetSetting("updateOnWindowActivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)
    If Not GetSetting("updateOnWindowDeactivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    If Not GetSetting("updateOnWorkbookActivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook)
    If Not GetSetting("updateOnWorkbookDeactivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    If Not GetSetting("updateOnWorkbookOpen", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub
