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

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    FinboxioFixLinksModule.FixAddinLinks Wb
    PromoteStagedUpdate
End Sub

Private Sub App_AfterCalculate()
    PromoteStagedUpdate
End Sub

Private Sub App_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    PromoteStagedUpdate
End Sub

Private Sub App_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook)
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    PromoteStagedUpdate
End Sub

Private Sub App_NewWorkbook(ByVal Wb As Workbook)
    PromoteStagedUpdate
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    PromoteStagedUpdate
End Sub

Private Sub App_SheetCalculate(ByVal Sh As Object)
    PromoteStagedUpdate
End Sub

Private Sub App_SheetDeactivate(ByVal Sh As Object)
    PromoteStagedUpdate
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    PromoteStagedUpdate
End Sub

Private Sub App_WorkbookNewSheet(ByVal Wb As Workbook, ByVal Sh As Object)
    PromoteStagedUpdate
End Sub

