VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CredentialsForm 
   Caption         =   "finbox.io Login"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6045
   OleObjectBlob   =   "CredentialsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CredentialsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub SignUpLabel_Click()
    ThisWorkbook.FollowHyperlink SIGNUP_URL
End Sub

Private Sub UserForm_Initialize()
    Me.emailBox.SetFocus
End Sub

Private Sub emailBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = vbKeyRButton Then Call RightClickMenu
End Sub

Private Sub passBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = vbKeyRButton Then Call RightClickMenu
End Sub

Private Sub LoginButton_Click()
    Dim success As Boolean
    success = Login(Me.emailBox.value, Me.passBox.value)
    If success Then
        Unload Me
        Application.CalculateFull
    End If
End Sub

Private Sub LoginButtonBg_Click()
    Dim success As Boolean
    success = Login(Me.emailBox.value, Me.passBox.value)
    If success Then
        Unload Me
        Application.CalculateFull
    End If
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

