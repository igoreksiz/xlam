Attribute VB_Name = "UpdaterModule"
Option Explicit
Option Private Module

Private lastUpdateCheck As Date

Public Sub AutoUpdateCheck()
    If Not GetSetting("autoUpdate", True) Then Exit Sub
    ' Default to one auto-check per day, but allow
    ' interval to be specified in minutes (primarily
    ' for testing)
    Dim interval As Integer
    interval = CInt(GetSetting("autoUpdateMinutes", 24 * 60))
    If VBA.Now() - (interval / (24 * 60)) > lastUpdateCheck Then
        Call DownloadUpdates(blockEvents:=True)
    End If
End Sub

' Primarily used for testing staging transitions,
' this forces the latest version to be downloaded
' and staged.
Public Function ForceUpdate()
    ForceUpdate = DownloadUpdates(blockEvents:=True, force:=GetSetting("forceUpdate", False))
End Function

' Downloads and stages the latest release from github
' if not already up-to-date. Returns True if there are
' staged updates to be applied.
Public Function DownloadUpdates(Optional blockEvents As Boolean, Optional force As Boolean) As Boolean
    If HasUpdates And Not force Then
        DownloadUpdates = True
        Exit Function
    End If
    
    lastUpdateCheck = VBA.Now()
    
    Dim allowPrereleases As Boolean
    allowPrereleases = GetSetting("allowPrereleases", False)
    
    Dim latest As String, _
        current As String, _
        lReleased As String, _
        cReleased As String, _
        loaderUrl As String, _
        functionsUrl As String, _
        releaseUrl As String, _
        download As Integer, _
        functionsVersion As String, _
        autoSec As MsoAutomationSecurity, _
        lReleaseDate As Date
        
    Dim WebClient As New WebClient, _
        WebRequest As New WebRequest, _
        WebResponse As WebResponse, _
        asset As Object
    
    autoSec = Application.AutomationSecurity
    WebClient.BlockEventLoop = blockEvents
    
    ' Skip update check if AddInVersion is not set
    ' This probably indicates something is wrong
    ' with the current Excel session and it should
    ' be restarted.
    If AddInVersion = "" Then
        GoTo Finish
    End If
    
    functionsVersion = AddInVersion(AddInFunctionsFile)
    If functionsVersion = "" And HasAddInFunctions Then
        Dim functionsWb As Workbook
        Application.AutomationSecurity = msoAutomationSecurityForceDisable
        Set functionsWb = Workbooks.Open(LocalPath(AddInFunctionsFile))
        functionsVersion = AddInVersion(AddInFunctionsFile)
        functionsWb.Close
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
        lReleaseDate = CDate(VBA.DateValue(VBA.Mid(lReleased, 1, 10)) + VBA.TimeValue(VBA.Mid(lReleased, 12, 8)))
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

    ' The functions add-in isn't available, but we have
    ' the latest release, download just the functions
    ' component. This may happen during installation and
    ' manual upgrades.
    If functionsVersion = "" And cReleased = lReleased Then
        DownloadFile functionsUrl, StagingPath(AddInFunctionsFile)
        VBA.SetAttr StagingPath(AddInFunctionsFile), vbHidden
        functionsVersion = AddInVersion
    End If
    
    If functionsVersion <> AddInVersion Then
        ' For some reason the manager and function components
        ' are out of sync. Force a re-download of the latest
        download = vbYes
    ElseIf lReleased = "" Then
        ' We were unable to get the latest release from github.
        ' TODO: Error handling here
    ElseIf cReleased = "" And lReleaseDate > AddInReleaseDate Then
        ' User is running an unreleased version of the add-in.
        ' This may happen if we delete a release from github or
        ' if we send a hotfixed/beta version.
        '
        ' If the release was deleted from github, we probably
        ' want to downgrade to the current latest
        '
        ' If we sent this as a hotfix, we probably don't want
        ' to update unless the latest release was created after
        ' the hotfix version.
        '
        download = vbYes
    ElseIf cReleased < lReleased Then
        ' There is a newer version available
        download = vbYes
    End If

    If force Or download = vbYes Then
        DownloadFile loaderUrl, StagingPath(AddInInstalledFile)
        VBA.SetAttr StagingPath(AddInInstalledFile), vbHidden
        
        DownloadFile functionsUrl, StagingPath(AddInFunctionsFile)
        VBA.SetAttr StagingPath(AddInFunctionsFile), vbHidden
    End If
    
Finish:
    Application.AutomationSecurity = autoSec
    DownloadUpdates = HasUpdates
End Function

Public Function HasUpdates() As Boolean
    HasUpdates = IsStaged(AddInInstalledFile) Or IsStaged(AddInFunctionsFile)
End Function

Private Function IsStaged(file As String) As Boolean
    IsStaged = _
        Dir(StagingPath(file)) <> "" Or _
        Dir(StagingPath(file), vbHidden) <> ""
End Function
