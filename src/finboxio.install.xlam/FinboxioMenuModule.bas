Attribute VB_Name = "FinboxioMenuModule"
Option Explicit

Public AppRibbon
Private ButtonDefs(1 To 7) As String

Public Sub InvalidateAppRibbon()
    If Not TypeName(AppRibbon) = "Empty" Then
        If Not AppRibbon Is Nothing Then
            AppRibbon.Invalidate
        End If
    End If
End Sub

Public Sub FinboxioRibbonLoad(ribbon)
    Set AppRibbon = ribbon
End Sub

Public Sub FinboxioLoggedIn(control, ByRef enabled)
    enabled = IsLoggedIn()
End Sub

Public Sub FinboxioLoggedOut(control, ByRef enabled)
    enabled = IsLoggedOut()
End Sub

Public Sub FinboxioShowLogin(Optional control)
    ShowLoginForm
End Sub

Public Sub FinboxioLogout(Optional control)
    Call Logout
End Sub

Public Sub FinboxioMessages(Optional control)
    Call ShowMessages
End Sub

Public Sub FinboxioHelp(Optional control)
    Call LoadHelp
End Sub

Public Sub FinboxioRefresh(Optional control)
    Call RefreshData
End Sub

Public Sub FinboxioUnlink(Optional control)
    Call UnlinkFormulas
End Sub

Public Sub FinboxioUpdate(Optional control)
    Call CheckUpdates(True)
End Sub

Public Sub AddCustomMenu()
    ' Add macro into menu bar (Mac Excel 2011)

    ' Button definitions:  Cap&tion,MacroName,ToolTip,IconId,BeginGroupBool
    '      (IconId 39 is blue right arrow, and is a good default option)
    
    ButtonDefs(1) = "Log&in,FinboxioShowLogin,Login to finbox.io API,39,True"
    ButtonDefs(2) = "Log&out,FinboxioLogout,Logout from finbox.io API,39,False"
    
    ButtonDefs(3) = "&Refresh data,FinboxioRefresh,Recalculate open Excel Workbooks,39,True"
    ButtonDefs(4) = "Un&link Formulas,FinboxioUnlink,Unlink finbox.io formulas,39,False"
    
    ButtonDefs(5) = "&Message Log,FinboxioMessages,Display message log,39,True"
    ButtonDefs(6) = "Check For &Updates,FinboxioUpdate,Check for updates,39,False"
    ButtonDefs(7) = "&Help,FinboxioHelp,Read the finbox.io add-in guide,39,False"

    Dim bd As Integer
    Dim butdefs() As String
    
    If EXCEL_VERSION = "Mac2011" Then
     
        ' Office 2003 and earlier or Mac 2011
        ' Add (or retrieve) top level menu "Add-Ins"
        Dim CustomMenu As CommandBarPopup
        
        On Error GoTo 0
        Set CustomMenu = Application.CommandBars("Worksheet Menu Bar").Controls.Add(msoControlPopup, _
                temporary:=True)  ' before "Data"
        With CustomMenu
            .Caption = "&finbox.io"
            .Tag = "finbox.io"
            .enabled = True
            .Visible = True
        End With
    
        ' Add buttons to top level menu "Add-Ins"
        With CustomMenu.Controls
            For bd = LBound(ButtonDefs) To UBound(ButtonDefs)
                butdefs = Split(ButtonDefs(bd), ",")
                With .Add(msoControlButton, temporary:=True)
                    .Caption = butdefs(0)
                    .Tag = Replace(butdefs(0), "&", "")
                    .OnAction = butdefs(1)
                    .TooltipText = butdefs(2)
                    .Style = 3
                    .FaceId = CInt(butdefs(3))
                    
                    If VBA.LCase(butdefs(4)) = "true" Then
                        .BeginGroup = True
                    Else
                        .BeginGroup = False
                    End If
                End With
            Next
         End With
    End If
End Sub

Public Sub DeleteCustomMenu()
    Dim bd As Integer
    Dim butdefs() As String
    
    For bd = UBound(ButtonDefs) To LBound(ButtonDefs) Step -1
        On Error Resume Next
        butdefs = Split(ButtonDefs(bd), ",")
          
        If EXCEL_VERSION = "Mac2011" Then
            ' Delete buttons and top level menu "Custom"
            With Application.CommandBars("Worksheet Menu Bar").Controls("Add-Ins")
                .Controls(Replace(butdefs(0), "&", "")).Delete
                If .Controls.count = 0 Then .Delete
            End With
        End If
        On Error GoTo 0
    Next bd
End Sub
