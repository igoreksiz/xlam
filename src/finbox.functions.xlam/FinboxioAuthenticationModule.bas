Attribute VB_Name = "FinboxioAuthenticationModule"
Option Explicit
Option Private Module

Private APIKeyStore As APIKeyHandler
Private tier As String

Public Sub SetAPIKeyHandler(handler As APIKeyHandler)
    Set APIKeyStore = handler
End Sub

Public Function ShowLoginForm()
    If ExcelVersion = "Mac2011" Then
        Mac2011CredentialsForm.Show
    ElseIf ExcelVersion = "Mac2016" Or ExcelVersion = "Mac2019" Then
        Mac2016CredentialsForm.Show
    Else
        DefaultCredentialsForm.Show
    End If
End Function

Public Function Login(ByVal email As String, ByVal password As String) As Boolean
    Login = False
    
    If APIKeyStore Is Nothing Then
        Set APIKeyStore = New APIKeyHandler
    End If
    
    ' build json request and convert to postData string
    Dim jsonReqObj As Object
    Set jsonReqObj = ParseJson("{}")
    
    jsonReqObj.Item("email") = email
    jsonReqObj.Item("password") = password
    
    Dim postData As String
    postData = ConvertToJson(jsonReqObj)
   
     ' POST login request
    Dim webClient As New webClient
    webClient.BaseUrl = AUTH_URL
    
    Dim webRequest As New webRequest
    webRequest.Method = WebMethod.HttpPost
    webRequest.RequestFormat = WebFormat.Json
    webRequest.ResponseFormat = WebFormat.Json
    webRequest.Body = postData
    
    Dim webResponse As webResponse
    Set webResponse = webClient.Execute(webRequest)
    
    ' Process according to HTTP response code
    Dim APItier As String
    Dim APIKey As String
    
    Select Case webResponse.statusCode
        Case 401
            MsgBox _
                Title:="[Finbox] Login Error", _
                Prompt:="The provided credentials are invalid.", _
                Buttons:=vbCritical
        Case 200
            ' Extract api_tier and api_key from json response
            APItier = ""
            APIKey = ""
            On Error Resume Next
            APItier = webResponse.Data.Item("pro_status")
            APIKey = webResponse.Data.Item("api_key")
            On Error GoTo ErrorHandler
            
            LogMessage "Logged in as " & APItier & " user " & email
            
            ' Process api_tier and api_key
            If APItier = "inactive" Then
                MsgBox _
                    Title:="[Finbox] Login Error", _
                    Prompt:="You have not verified your email address yet." & vbCrLf & _
                            "To resend the verification email, visit https://finbox.com/profile.", _
                    Buttons:=vbCritical
            Else
                APIKeyStore.StoreApiKey (APIKey)
                Login = True
            End If
        Case Else
            GoTo ErrorHandler
    End Select
    
    tier = ""
    
    Set jsonReqObj = Nothing
    Set webClient = Nothing
    Set webRequest = Nothing
    Set webResponse = Nothing
    
    CheckQuota
    InvalidateAppRibbon
    Exit Function

ErrorHandler:
    tier = ""
    CheckQuota
    InvalidateAppRibbon
    Dim answer As Integer
    MsgBox _
        Title:="[Finbox] Login Error", _
        Prompt:="Unable to authenticate with finbox.com at this time. Please try again and contact support@finbox.com if this problem persists.", _
        Buttons:=vbCritical
End Function

Public Function GetTier()
    ' Only load tier once
    If Not tier = "" Then GoTo OnError
    
    On Error GoTo OnError
    tier = "anonymous"
    
    Dim webClient As New webClient
    webClient.BlockEventLoop = True
    webClient.BaseUrl = TierUrl
    
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
    webRequest.AddHeader "X-Finbox-Addon", GetAPIHeader()
    
    Dim webResponse As webResponse
    Set webResponse = webClient.Execute(webRequest)

    tier = webResponse.Data.Item("data").Item("tier")
OnError:
    GetTier = tier
End Function

Public Function StoreApiKey(key As String)
    If APIKeyStore Is Nothing Then
        Set APIKeyStore = New APIKeyHandler
    End If
    APIKeyStore.StoreApiKey (key)
    StoreApiKey = True
End Function

Public Sub Logout()
    If APIKeyStore Is Nothing Then
        Set APIKeyStore = New APIKeyHandler
    End If
    APIKeyStore.ClearAPIKey
    tier = ""
    CheckQuota
End Sub

Public Function IsLoggedIn()
    Dim key As String
    key = GetAPIKey()
    IsLoggedIn = key <> ""
End Function

Public Function IsLoggedOut()
    Dim key As String
    key = GetAPIKey()
    IsLoggedOut = key = ""
End Function

Public Function GetAPIKey() As String
    If APIKeyStore Is Nothing Then
        Set APIKeyStore = New APIKeyHandler
    End If
    GetAPIKey = APIKeyStore.GetAPIKey()
End Function


