Attribute VB_Name = "FinboxioUpdateModule"
Option Explicit

Public Sub CheckUpdates(Optional explicit As Boolean = False)
    Dim latest As String
    Dim url As String
    Dim answer As Integer
    
    latest = ""
    
    On Error GoTo Confirmation
    
    Dim webClient As New webClient
    webClient.BaseUrl = UPDATES_URL
    
    Dim webRequest As New webRequest
    webRequest.Method = WebMethod.HttpGet
    webRequest.ResponseFormat = WebFormat.Json
    
    Dim webResponse As webResponse
    Set webResponse = webClient.Execute(webRequest)

    Select Case webResponse.statusCode
    Case 200
        latest = webResponse.Data("version")
        url = webResponse.Data("download_url")
    End Select
    
Confirmation:
    If latest = "" Then
        answer = MsgBox("Unable to check for updates to the finbox.io add-on at this time. Please contact support@finbox.io if this problem persists.", vbCritical, AppTitle)
    ElseIf latest > AppVersion Then
        answer = MsgBox("A new version of the finbox.io add-on is available! Would you like to upgrade to " & latest & " now?", vbYesNo + vbQuestion, AppTitle)
        If answer = vbYes Then
            ThisWorkbook.FollowHyperlink UPDATE_URL
        End If
    ElseIf explicit And latest = AppVersion Then
        answer = MsgBox("You're already running the latest version of the finbox.io add-on! Please enjoy responsibly.", vbOKOnly, AppTitle)
    End If
End Sub

Public Function latestVersion()
    On Error GoTo Finish
    
    Dim latest As String
    latest = ""
    
    Dim webClient As New webClient
    webClient.BaseUrl = UPDATES_URL
    
    Dim webRequest As New webRequest
    webRequest.Method = WebMethod.HttpGet
    webRequest.ResponseFormat = WebFormat.Json
    
    Dim webResponse As webResponse
    Set webResponse = webClient.Execute(webRequest)

    Select Case webResponse.statusCode
    Case 200
        latest = webResponse.Data("version")
    End Select
Finish:
    latestVersion = latest
End Function
