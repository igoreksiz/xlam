VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Mac2016CredentialsForm 
   Caption         =   "finbox.io Login"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8700
   OleObjectBlob   =   "Mac2016CredentialsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Mac2016CredentialsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SignUpLabel_Click()
    ThisWorkbook.FollowHyperlink USAGE_URL
End Sub

Private Sub UserForm_Initialize()
    Me.emailBox.SetFocus
End Sub

Private Sub emailBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Dim DataObj As MsForms.DataObject
   Set DataObj = New MsForms.DataObject
   On Error GoTo 0
   DataObj.GetFromClipboard
   Me.emailBox.value = DataObj.GetText(1)
End Sub

Private Sub LoginButton_Click()
    StoreApiKey (Me.emailBox.value)
    Unload Me
End Sub

Private Sub LoginButtonBg_Click()
    StoreApiKey (Me.emailBox.value)
    Unload Me
End Sub

Private Sub LoginButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.LoginButton.BackColor = RGB(10, 37, 88)
    Me.LoginButtonBg.BackColor = RGB(10, 37, 88)
End Sub

Private Sub LoginButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.LoginButton.BackColor = RGB(21, 81, 195)
    Me.LoginButtonBg.BackColor = RGB(21, 81, 195)
End Sub

Private Sub LoginButtonBg_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.LoginButton.BackColor = RGB(10, 37, 88)
    Me.LoginButtonBg.BackColor = RGB(10, 37, 88)
End Sub

Private Sub LoginButtonBg_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.LoginButton.BackColor = RGB(21, 81, 195)
    Me.LoginButtonBg.BackColor = RGB(21, 81, 195)
End Sub

