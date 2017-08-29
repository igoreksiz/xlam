Attribute VB_Name = "FinboxioMenuModule"
Option Explicit
Option Private Module

#If Mac Then
    Public AppRibbon
#Else
    Public AppRibbon As IRibbonUI
#End If

Private ButtonDefs(1 To 13) As String

Public Sub InvalidateAppRibbon()
    If Not TypeName(AppRibbon) = "Empty" Then
        If Not AppRibbon Is Nothing Then
            AppRibbon.Invalidate
        End If
    End If
    UpdateCustomMenu
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
        MsgBox _
            Title:="[finbox.io] Quota Unavailable", _
            Prompt:="Quota usage is unavailable at this time. Please try again and contact support@finbox.io if this problem persists.", _
            Buttons:=vbCritical
    Else
        MsgBox _
            Title:="[finbox.io] Quota Usage", _
            Prompt:="You have used " & QuotaUsed & " datapoints of your " & QuotaTotal & " quota limit.", _
            Buttons:=vbInformation
    End If
End Sub

Public Sub FinboxioShowLogin(Optional control)
    ShowLoginForm
End Sub

Public Sub FinboxioLogout(Optional control)
    Call Logout
End Sub

Public Sub FinboxioAbout(Optional control)
    Dim ManagerVersion As String, ManagerDate As Date, ManagerLocation As String
    Dim FunctionsVersion As String, FunctionsDate As Date, FunctionsLocation As String
    Dim APIKey As String
    
    ManagerVersion = AddInVersion(AddInManagerFile)
    ManagerDate = AddInReleaseDate(AddInManagerFile)
    ManagerLocation = AddInLocation(AddInManagerFile)
    
    FunctionsVersion = AddInVersion(AddInFunctionsFile)
    FunctionsDate = AddInReleaseDate(AddInFunctionsFile)
    FunctionsLocation = AddInLocation(AddInFunctionsFile)
    
    APIKey = GetAPIKey()
    
    Dim msg As String
    msg = _
        "Installation Details" & vbCrLf & _
        "--------------------" & vbCrLf & _
        vbCrLf & _
        "  Add-In Components:" & vbCrLf & _
        vbCrLf & _
        "    * Add-In Manager (v" & ManagerVersion & ", " & VBA.Round(ManagerDate) & ")" & vbCrLf & _
        "      " & ManagerLocation & vbCrLf & _
        vbCrLf & _
        "    * Add-In Functions (v" & FunctionsVersion & ", " & VBA.Round(FunctionsDate) & ")" & vbCrLf & _
        "      " & FunctionsLocation & vbCrLf & _
        vbCrLf & _
        "  Current User:" & vbCrLf & _
        vbCrLf & _
        "    API Key " & APIKey & vbCrLf & _
        vbCrLf & _
        "  Contact Information: " & vbCrLf & _
        vbCrLf & _
        "    Please help us improve your experience by reporting " & vbCrLf & _
        "    any issues and sending suggestions to support@finbox.io, " & vbCrLf & _
        "    or visit https://finbox.io to chat with us live." & vbCrLf & _
        vbCrLf & _
        vbCrLf & _
        "  Thank you for using finbox.io!"
        
    MsgBox _
        Title:="[finbox.io] Add-in Information", _
        Prompt:=msg
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

Public Sub FinboxioUnlinkImage(control, ByRef image)
    image = "HyperlinkRemove"
    If ExcelVersion = "Win2010" Then
        image = "SkipOccurrence"
    End If
End Sub

Public Sub FinboxioUpdate(Optional control)
    On Error GoTo Finish
    updatingManager = True
    Application.Run (AddInInstalledFile & "!CheckUpdates")
    PromoteStagedUpdate
Finish:
    updatingManager = False
End Sub

Public Sub UpdateCustomMenu()
    If ExcelVersion = "Mac2011" Then
        Dim CustomMenu As CommandBarPopup
        Dim Controls, i As Integer
        Set Controls = Application.CommandBars("Worksheet Menu Bar").Controls
        For i = 1 To Controls.count
            Dim control
            Set control = Controls.Item(i)
            If control.Tag = "finbox.io" Then Set CustomMenu = control
        Next i
        
        On Error GoTo 0
        If IsEmpty(CustomMenu) Or CustomMenu Is Nothing Then
            ' Add macro into menu bar (Mac Excel 2011)

            ' Button definitions:  Cap&tion,MacroName,ToolTip,IconId,BeginGroupBool
            '      (IconId 39 is blue right arrow, and is a good default option)
            
            ButtonDefs(1) = "Log&in,FinboxioShowLogin,Login to finbox.io API,2882,True"
            ButtonDefs(2) = "Log&out,FinboxioLogout,Logout from finbox.io API,1019,False"
            ButtonDefs(3) = "&Pro,FinboxioUpgrade,Upgrade to premium access,225,False"
            ButtonDefs(4) = "Check &Quota,FinboxioCheckQuota,Check quota usage,52,False"
            
            ButtonDefs(5) = "&Watchlist,FinboxioWatchlist,Go to your watchlist,183,True"
            ButtonDefs(6) = "&Screener,FinboxioScreener,Go to the online screener,601,False"
            ButtonDefs(7) = "&Templates,FinboxioTemplates,Download pre-built templates,357,False"
            ButtonDefs(8) = "&Help,FinboxioHelp,Read the finbox.io add-in guide,49,False"
            
            ButtonDefs(9) = "&Refresh Data,FinboxioRefresh,Recalculate open Excel Workbooks,37,True"
            ButtonDefs(10) = "Un&link Formulas,FinboxioUnlink,Unlink finbox.io formulas,2309,False"
            
            ButtonDefs(11) = "&Message Log,FinboxioMessages,Display message log,588,True"
            ButtonDefs(12) = "Check For &Updates,FinboxioUpdate,Check for updates,273,False"
            ButtonDefs(13) = "&About,FinboxioAbout,Information about the add-on,487,False"
        
            Dim bd As Integer
            Dim butdefs() As String
        
            Set CustomMenu = Application.CommandBars("Worksheet Menu Bar").Controls.Add(msoControlPopup, temporary:=True)
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
        
        CustomMenu.Controls.Item(1).Visible = Not IsLoggedIn()
        CustomMenu.Controls.Item(2).Visible = IsLoggedIn()
        CustomMenu.Controls.Item(3).Visible = (GetTier() = "free")
        CustomMenu.Controls.Item(4).Caption = QuotaLabel
        CustomMenu.Controls.Item(4).Tag = QuotaLabel
        If QuotaImage = "Piggy" Then CustomMenu.Controls.Item(4).FaceId = 52
        If QuotaImage = "HappyFace" Then CustomMenu.Controls.Item(4).FaceId = 59
        If QuotaImage = "TraceError" Then CustomMenu.Controls.Item(4).FaceId = 463
        If QuotaImage = "HighImportance" Then CustomMenu.Controls.Item(4).FaceId = 459
    End If
End Sub
