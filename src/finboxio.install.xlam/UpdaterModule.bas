Attribute VB_Name = "UpdaterModule"
Option Explicit
Option Private Module

Public Sub DownloadUpdates(Optional explicit As Boolean = False, Optional wb As Workbook)
    Dim allowPrereleases As Boolean: allowPrereleases = True
    
    Dim latest As String, _
        current As String, _
        lReleased As String, _
        cReleased As String, _
        loaderUrl As String, _
        functionsUrl As String, _
        releaseUrl As String, _
        answer As Integer
        
    Dim WebClient As New WebClient, _
        WebRequest As New WebRequest, _
        WebResponse As WebResponse, _
        asset As Object
    
    WebClient.BlockEventLoop = Not explicit
    
    ' Skip update check if AppVersion is not set
    If AddInVersion = "" Then
        If explicit Then GoTo Confirmation
        GoTo Finish
    End If
    
GetCurrent:
    On Error GoTo GetLatest
    WebClient.BaseUrl = RELEASES_URL & "/tags/v" & AddInVersion
    WebRequest.Method = WebMethod.HttpGet
    WebRequest.ResponseFormat = WebFormat.Json
    Set WebResponse = WebClient.Execute(WebRequest)
    Select Case WebResponse.statusCode
    Case 200
        current = WebResponse.data.Item("tag_name")
        cReleased = WebResponse.data.Item("created_at")
        releaseUrl = WebResponse.data.Item("html_url")
        For Each asset In WebResponse.data.Item("assets")
            If asset.Item("name") = "finboxio.install.xlam" Then
                loaderUrl = asset.Item("browser_download_url")
            End If
            If asset.Item("name") = "finboxio.functions.xlam" Then
                functionsUrl = asset.Item("browser_download_url")
            End If
        Next asset
    End Select
    
GetLatest:
    On Error GoTo Confirmation

    WebClient.BaseUrl = RELEASES_URL & "/latest"
    If allowPrereleases Then WebClient.BaseUrl = RELEASES_URL

    WebRequest.Method = WebMethod.HttpGet
    WebRequest.ResponseFormat = WebFormat.Json
    Set WebResponse = WebClient.Execute(WebRequest)
    Select Case WebResponse.statusCode
    Case 200
        Dim release: Set release = WebResponse.data
        If TypeName(release) = "Collection" Then Set release = WebResponse.data(1)
        latest = release.Item("tag_name")
        lReleased = release.Item("created_at")
        releaseUrl = release.Item("html_url")
        For Each asset In release.Item("assets")
            If asset.Item("name") = "finboxio.install.xlam" Then
                loaderUrl = asset.Item("browser_download_url")
            End If
            If asset.Item("name") = "finboxio.functions.xlam" Then
                functionsUrl = asset.Item("browser_download_url")
            End If
        Next asset
    End Select

Confirmation:
    On Error GoTo Finish

    If AddInVersion = "" Or lReleased = "" Or releaseUrl = "" Then
        answer = MsgBox("Unable to check for updates to the finbox.io add-on at this time. Please contact support@finbox.io if this problem persists.", vbCritical)
        ' LogMessage "Failed to check for updates."
    ElseIf cReleased = "" Then
        answer = MsgBox("You are using an unreleased version of the finbox.io add-on. Would you like to download the latest release?", vbYesNo + vbQuestion)
        ' LogMessage "Unreleased add-on version detected."
    ElseIf cReleased < lReleased Then
        answer = MsgBox("A newer version of the finbox.io add-on is available! Would you like to download the latest release?", vbYesNo + vbQuestion)
        ' LogMessage "Add-on update " & latest & " - " & lReleased & " is available. (Upgrading from " & current & " - " & cReleased & ")"
    ElseIf lReleased = cReleased Then
        If explicit Then answer = MsgBox("You are already using the latest version of the finbox.io add-on. Congratulations!")
        ' LogMessage "No updates available."
    End If

    If answer = vbYes Or Not HasAddInFunctions Then
        DownloadFile loaderUrl, StagedXlamPath(AddInLoaderFile)
        DownloadFile functionsUrl, StagedXlamPath(AddInFunctionsFile)

        MsgBox "The update is ready and will take effect after you restart Excel."

        ' If wb Is Nothing Or VBA.IsEmpty(wb) Then Set wb = ThisWorkbook
        ' If Not (wb Is Nothing Or VBA.IsEmpty(wb)) Then wb.FollowHyperlink releaseUrl
    End If

Finish:
    
End Sub

Public Function HasStagedUpdates(file As String) As Boolean
    HasStagedUpdates = Dir(StagedXlamPath(AddInFunctionsFile)) <> ""
End Function

Public Function XlamFile(file As String) As String
    XlamFile = file & ".xlam"
End Function

Public Function StagedXlamFile(file As String) As String
    StagedXlamFile = XlamFile(file & ".staged")
End Function

Public Function XlamPath(file As String) As String
    XlamPath = ThisWorkbook.path & Application.PathSeparator & XlamFile(file)
End Function

Public Function StagedXlamPath(file As String) As String
    StagedXlamPath = ThisWorkbook.path & Application.PathSeparator & StagedXlamFile(file)
End Function

Public Sub PromoteStagedUpdate(file As String)
    FileCopy StagedXlamPath(file), XlamPath(file)
End Sub
