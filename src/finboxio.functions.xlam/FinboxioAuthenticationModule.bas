Attribute VB_Name = "FinboxioAuthenticationModule"
Option Explicit
Option Private Module

Private APIKeyStore As APIKeyHandler
Private tier As String

Public Sub SetAPIKeyHandler(handler As APIKeyHandler)
    Set APIKeyStore = handler
End Sub

Public Function ShowLoginForm()
    If EXCEL_VERSION = "Mac2011" Then
        Mac2011CredentialsForm.Show
    ElseIf EXCEL_VERSION = "Mac2016" Then
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
            MsgBox "The provided credentials are invalid.", vbCritical
        Case 200
            ' Extract api_tier and api_key from json response
            APItier = ""
            APIKey = ""
            On Error Resume Next
            APItier = webResponse.data.Item("pro_status")
            APIKey = webResponse.data.Item("api_key")
            On Error GoTo ErrorHandler
            
            LogMessage "Logged in as " & APItier & " user " & email
            
            ' Process api_tier and api_key
            If APItier = "inactive" Then
                MsgBox "You have not verified your email address yet." & vbCrLf & _
                    "To resend the verification email, visit https://finbox.io/profile.", _
                    vbCritical
            Else
                APIKeyStore.StoreApiKey (APIKey)
                Login = True
            End If
        Case Else
            MsgBox "The finbox.io API returned http status code " & webResponse.statusCode & " = " & vbCr & _
                VBA.Trim(webResponse.StatusDescription), vbCritical
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
    answer = MsgBox("Failed to authenticate with finbox.io. Contact support@finbox.io if this problem persists.", vbCritical, "finbox.io Addin")
End Function

Public Function GetTier()
    ' Only load tier once
    If Not tier = "" Then GoTo OnError
    
    On Error GoTo OnError
    tier = "anonymous"
    
    Dim webClient As New webClient
    webClient.BaseUrl = TIER_URL
    
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

    tier = webResponse.data.Item("data").Item("tier")
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


