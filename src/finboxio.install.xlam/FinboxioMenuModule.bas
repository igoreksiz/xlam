Attribute VB_Name = "FinboxioMenuModule"
Option Explicit

Public AppRibbon
Private ButtonDefs(1 To 7) As String

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
    CredentialsForm.Show
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
    
    ButtonDefs(1) = "finbox.io Log&in,FinboxioShowLogin,Login to finbox.io API,39,True"
    ButtonDefs(2) = "finbox.io Log&out,FinboxioLogout,Logout from finbox.io API,39,True"
    ButtonDefs(3) = "finbox.io &Messages,FinboxioMessages,Display message log,39,True"
    ButtonDefs(4) = "finbox.io Re&calculate,FinboxioRefresh,Recalculate open Excel Workbooks,39,True"
    ButtonDefs(5) = "finbox.io Updates,FinboxioUpdate,Check for updates,39,True"
    ButtonDefs(6) = "finbox.io Unlink,FinboxioUnlink,Unlink finbox.io formulas,39,True"
    ButtonDefs(7) = "finbox.io Help,FinboxioHelp,Read the finbox.io add-in guide,39,True"

    Dim bd As Integer
    Dim butdefs() As String
    
    Dim MacExcel2011 As Boolean
    MacExcel2011 = False
    
    #If Mac Then
        #If MAC_OFFICE_VERSION < 15 Then
            MacExcel2011 = True
        #End If
    #End If
        
    If MacExcel2011 Then
     
        ' Office 2003 and earlier or Mac 2011
        ' Add (or retrieve) top level menu "Add-Ins"
        Dim CustomMenu As CommandBarPopup
        
        On Error Resume Next
        Set CustomMenu = Application.CommandBars("Worksheet Menu Bar").Controls("Add-Ins")
        On Error GoTo 0
        
        If CustomMenu Is Nothing Then
                Set CustomMenu = Application.CommandBars("Worksheet Menu Bar").Controls.Add(msoControlPopup, _
                        temporary:=True, Before:=7)  ' before "Data"
                With CustomMenu
                    .Caption = "&Add-Ins"
                    .Tag = "Add-Ins"
                    .enabled = True
                    .Visible = True
                End With
        End If
    
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
                    
                    If LCase(butdefs(4)) = "true" Then
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
    
    Dim MacExcel2011 As Boolean
    MacExcel2011 = False
    
    #If Mac Then
        #If MAC_OFFICE_VERSION < 15 Then
            MacExcel2011 = True
        #End If
    #End If
    
    For bd = UBound(ButtonDefs) To LBound(ButtonDefs) Step -1
        On Error Resume Next
        butdefs = Split(ButtonDefs(bd), ",")
          
        If MacExcel2011 Then
            ' Delete buttons and top level menu "Custom"
            With Application.CommandBars("Worksheet Menu Bar").Controls("Add-Ins")
                .Controls(Replace(butdefs(0), "&", "")).Delete
                If .Controls.Count = 0 Then .Delete
            End With
        End If
        On Error GoTo 0
    Next bd
End Sub
