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

Private Sub App_AfterCalculate()
    CloseInstaller
    If Not GetSetting("updateOnAfterCalculate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    CloseInstaller
    If Not GetSetting("updateOnWindowActivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)
    CloseInstaller
    If Not GetSetting("updateOnWindowDeactivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    CloseInstaller
    If Not GetSetting("updateOnWorkbookActivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook)
    CloseInstaller
    If Not GetSetting("updateOnWorkbookDeactivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    CloseInstaller
    If Not GetSetting("updateOnWorkbookOpen", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    CloseInstaller
    If Not GetSetting("updateOnWorkbookBeforeClose", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    CloseInstaller
    If Not GetSetting("updateOnWorkbookAfterSave", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_NewWorkbook(ByVal Wb As Workbook)
    CloseInstaller
    If Not GetSetting("updateOnNewWorkbook", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    CloseInstaller
    If Not GetSetting("updateOnSheetActivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_SheetCalculate(ByVal Sh As Object)
    CloseInstaller
    If Not GetSetting("updateOnSheetCalculate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_SheetDeactivate(ByVal Sh As Object)
    CloseInstaller
    If Not GetSetting("updateOnSheetDeactivate", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    CloseInstaller
    If Not GetSetting("updateOnSheetSelectionChange", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookNewSheet(ByVal Wb As Workbook, ByVal Sh As Object)
    CloseInstaller
    If Not GetSetting("updateOnWorkbookNewSheet", True) Then Exit Sub
    AutoUpdateCheck
    PromoteStagedUpdate
End Sub
