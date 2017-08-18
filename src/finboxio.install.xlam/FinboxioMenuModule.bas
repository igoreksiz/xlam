Attribute VB_Name = "FinboxioMenuModule"
Option Explicit
Option Private Module

#If Mac Then
    Public AppRibbon
#Else
    Public AppRibbon As IRibbonUI
#End If

Private ButtonDefs(1 To 12) As String

Public Sub InvalidateAppRibbon()
    If Not TypeName(AppRibbon) = "Empty" Then
        If Not AppRibbon Is Nothing Then
            AppRibbon.Invalidate
        End If
    End If
End Sub

#If Mac Then
Public Sub FinboxioRibbonLoad(ribbon)
    Set AppRibbon = ribbon
End Sub
#Else
Public Sub FinboxioRibbonLoad(ByRef ribbon As IRibbonUI)
    Set AppRibbon = ribbon
End Sub
#End If

Public Sub FinboxioLoggedIn(control, ByRef enabled)
    enabled = IsLoggedIn()
End Sub

Public Sub FinboxioLoggedOut(control, ByRef enabled)
    enabled = IsLoggedOut()
End Sub

Public Sub FinboxioIsFree(control, ByRef free)
    Dim key As String
    Dim tier As String
    
    key = GetAPIKey()
    tier = GetTier()
    
    free = False
    If key <> "" Then
        If tier = "anonymous" Or tier = "free" Then
            free = True
        End If
    End If
End Sub

Public Sub FinboxioQuotaLabel(control, ByRef label)
    label = QuotaLabel
End Sub

Public Sub FinboxioQuotaImage(control, ByRef image)
    image = QuotaImage
End Sub

Public Sub FinboxioCheckQuota(Optional control)
    CheckQuota
    If QuotaTotal < 1 Then
        MsgBox "Quota usage is unavailable at this time."
    Else
        MsgBox "You have used " & QuotaUsed & " datapoints of your " & QuotaTotal & " quota limit."
    End If
End Sub

Public Sub FinboxioShowLogin(Optional control)
    ShowLoginForm
End Sub

Public Sub FinboxioLogout(Optional control)
    Call Logout
End Sub

Public Sub FinboxioAbout(Optional control)
    MsgBox "You are using the " & AppTitle & vbCrLf & _
        "This add-on is installed as " & ThisWorkbook.path & Application.PathSeparator & ThisWorkbook.name & "." & vbCrLf & _
        "You can contact support@finbox.io with any questions or concerns." & vbCrLf & vbCrLf & _
        "Happy investing!"
End Sub

Public Sub FinboxioMessages(Optional control)
    Call ShowMessages
End Sub

Public Sub FinboxioHelp(Optional control)
    Call LoadHelp
End Sub

Public Sub FinboxioUpgrade(Optional control)
    ThisWorkbook.FollowHyperlink UPGRADE_URL
End Sub

Public Sub FinboxioProfile(Optional control)
    ThisWorkbook.FollowHyperlink PROFILE_URL
End Sub

Public Sub FinboxioWatchlist(Optional control)
    ThisWorkbook.FollowHyperlink WATCHLIST_URL
End Sub

Public Sub FinboxioScreener(Optional control)
    ThisWorkbook.FollowHyperlink SCREENER_URL
End Sub

Public Sub FinboxioTemplates(Optional control)
    ThisWorkbook.FollowHyperlink TEMPLATES_URL
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
    ButtonDefs(3) = "&Pro,FinboxioUpgrade,Upgrade to premium access,39,False"
    
    ButtonDefs(4) = "&Watchlist,FinboxioWatchlist,Go to your watchlist,39,True"
    ButtonDefs(5) = "&Screener,FinboxioScreener,Go to the online screener,39,False"
    ButtonDefs(6) = "&Templates,FinboxioTemplates,Download pre-built templates,39,False"
    
    ButtonDefs(7) = "&Refresh data,FinboxioRefresh,Recalculate open Excel Workbooks,39,True"
    ButtonDefs(8) = "Un&link Formulas,FinboxioUnlink,Unlink finbox.io formulas,39,False"
    
    ButtonDefs(9) = "&Message Log,FinboxioMessages,Display message log,39,True"
    ButtonDefs(10) = "Check For &Updates,FinboxioUpdate,Check for updates,39,False"
    ButtonDefs(11) = "&Help,FinboxioHelp,Read the finbox.io add-in guide,39,False"
    ButtonDefs(12) = "&About,FinboxioAbout,Information about the add-on,39,False"

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




