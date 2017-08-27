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

Private Sub App_WorkbookOpen(ByVal wb As Workbook)
    FinboxioFixLinksModule.FixAddinLinks wb
End Sub

Private Sub App_AfterCalculate()
    If HasStagedUpdate Then
        Dim loadingManager As Boolean
        loadingManager = Application.Run(AddInManagerFile & "!IsLoadingManager")
        If Not loadingManager Then PromoteStagedUpdate
    End If
End Sub
