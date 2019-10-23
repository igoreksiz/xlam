Attribute VB_Name = "FinboxioQuotaModule"
Option Explicit
Option Private Module

Public QuotaUsed As Long
Public QuotaTotal As Long

Public Sub UpdateQuota(used As Long, remaining As Long)
    QuotaUsed = used
    QuotaTotal = used + remaining
    If remaining > 0 Then ClearRateLimit
    InvalidateAppRibbon
End Sub

Public Function QuotaLabel() As String
    If QuotaTotal < 1 Then
        QuotaLabel = "Check Quota"
    ElseIf QuotaUsed >= QuotaTotal Then
        QuotaLabel = "Quota Exhausted"
    Else
        QuotaLabel = VBA.Round(QuotaUsed / QuotaTotal * 100) & "% Quota Usage"
    End If
End Function

Public Function QuotaImage() As String
    If QuotaTotal < 1 Then
        QuotaImage = "Piggy"
        Exit Function
    End If
    
    Dim pct As Integer
    pct = VBA.Round(QuotaUsed / QuotaTotal * 100)
    If pct < 70 Then
        QuotaImage = "HappyFace"
    ElseIf pct < 90 Then
        QuotaImage = "TraceError"
    Else
        QuotaImage = "HighImportance"
    End If
End Function

Public Sub CheckQuota(Optional blockEvents As Boolean)
    On Error GoTo ClearQuota
    
    Dim webClient As New webClient

    webClient.BlockEventLoop = blockEvents
    webClient.BaseUrl = TierUrl
    webClient.TimeoutMs = 5000

    ' Setup Basic Auth with API key as username and empty password
    Dim APIKey As String: APIKey = GetAPIKey()
    If APIKey <> "" Then
        Dim Auth As New HttpBasicAuthenticator
        Auth.Setup APIKey, ""
        Set webClient.Authenticator = Auth
    End If

    Dim webRequest As New webRequest
    webRequest.Method = WebMethod.HttpGet
    webRequest.ResponseFormat = WebFormat.Json
    webRequest.AddHeader "X-Finboxio-Addon", GetAPIHeader()

    Dim webResponse As webResponse
    Set webResponse = webClient.Execute(webRequest)

    Dim used As Long, remaining As Long, resets As String
    used = CLng(webResponse.Data.Item("data").Item("quota").Item("used"))
    remaining = CLng(webResponse.Data.Item("data").Item("quota").Item("remaining"))
    UpdateQuota used, remaining
    Exit Sub
    
ClearQuota:
    UpdateQuota 0, 0
End Sub
