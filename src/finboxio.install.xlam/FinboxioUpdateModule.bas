Attribute VB_Name = "FinboxioUpdateModule"
Option Explicit
Option Private Module

Public Sub CheckUpdates(Optional explicit As Boolean = False, Optional wb As Workbook)
    Dim latest As String, _
        current As String, _
        lReleased As String, _
        cReleased As String, _
        downloadUrl As String, _
        releaseUrl As String, _
        answer As Integer
        
    Dim webClient As New webClient, _
        webRequest As New webRequest, _
        webResponse As webResponse
    
GetCurrent:
    On Error GoTo GetLatest
    webClient.BaseUrl = RELEASES_URL & "/tags/v" & AppVersion
    webRequest.Method = WebMethod.HttpGet
    webRequest.ResponseFormat = WebFormat.Json
    Set webResponse = webClient.Execute(webRequest)
    Select Case webResponse.statusCode
    Case 200
        current = webResponse.data.Item("tag_name")
        cReleased = webResponse.data.Item("created_at")
    End Select
    
GetLatest:
    On Error GoTo Confirmation
    webClient.BaseUrl = RELEASES_URL & "/latest"
    webRequest.Method = WebMethod.HttpGet
    webRequest.ResponseFormat = WebFormat.Json
    Set webResponse = webClient.Execute(webRequest)
    Select Case webResponse.statusCode
    Case 200
        latest = webResponse.data.Item("tag_name")
        lReleased = webResponse.data.Item("created_at")
        releaseUrl = webResponse.data.Item("html_url")
        Dim assets, asset
        Set assets = webResponse.data.Item("assets")
        For Each asset In assets
            If asset.Item("name") = "finboxio.install.xlam" Then
                downloadUrl = asset.Item("browser_download_url")
            End If
        Next asset
    End Select
    
Confirmation:
    On Error GoTo Finish
    
    If lReleased = "" Or releaseUrl = "" Then
        answer = MsgBox("Unable to check for updates to the finbox.io add-on at this time. Please contact support@finbox.io if this problem persists.", vbCritical)
    ElseIf lReleased = cReleased And explicit Then
        answer = MsgBox("You are already using the latest version of the finbox.io add-on. Congratulations!")
    ElseIf cReleased = "" And explicit Then
        answer = MsgBox("You are using an unreleased version of the finbox.io add-on. Would you like to download the latest release?", vbYesNo + vbQuestion)
    ElseIf cReleased < lReleased Then
        answer = MsgBox("You are using an unreleased version of the finbox.io add-on. Would you like to download the latest release?", vbYesNo + vbQuestion)
    ElseIf lReleased > cReleased Then
        answer = MsgBox("A newer version of the finbox.io add-on is available! Would you like to download the latest release?", vbYesNo + vbQuestion)
    End If
    
    If answer = vbYes Then
        If wb Is Nothing Or TypeName(wb) = "Nothing" Or TypeName(wb) = "Empty" Then
            Set wb = ThisWorkbook
        End If
        If Not (wb Is Nothing Or TypeName(wb) = "Nothing" Or TypeName(wb) = "Empty") Then
            wb.FollowHyperlink releaseUrl
        End If
    End If

Finish:
    
End Sub
