Attribute VB_Name = "FinboxioAuthenticationModule"
' finbox.io API Integration

Option Explicit

Private APIKeyStore As APIKeyHandler
Private tier As String

Public Sub SetAPIKeyHandler(handler As APIKeyHandler)
    Set APIKeyStore = handler
End Sub

Public Function ShowLoginForm()
    If EXCEL_VERSION = "Mac2011" Then
        MacCredentialsForm.Show
    ElseIf EXCEL_VERSION = "Mac2016" Then
        Mac2016CredentialsForm.Show
    Else
        CredentialsForm.Show
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
    
    jsonReqObj("email") = email
    jsonReqObj("password") = password
    
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
    Dim APIkey As String
    
    Select Case webResponse.statusCode
        Case 401
            MsgBox "The provided credentials are invalid.", vbCritical, AppTitle
        Case 200
            ' Extract api_tier and api_key from json response
            APItier = ""
            APIkey = ""
            On Error Resume Next
            APItier = webResponse.data("api_tier")
            APIkey = webResponse.data("api_key")
            On Error GoTo ErrorHandler
            
            ' Process api_tier and api_key
            If APItier = "inactive" Then
                MsgBox "You have not verified your email address yet." & vbCrLf & _
                    "To resend the verification email, visit https://finbox.io/profile.", _
                    vbCritical, AppTitle
            Else
                APIKeyStore.StoreApiKey (APIkey)
                Login = True
            End If
        Case Else
            MsgBox "The finbox.io API returned http status code " & webResponse.statusCode & " = " & vbCr & _
                VBA.Trim(webResponse.StatusDescription), vbCritical, AppTitle
    End Select
    
    tier = ""
    
    Set jsonReqObj = Nothing
    Set webClient = Nothing
    Set webRequest = Nothing
    Set webResponse = Nothing
    
    InvalidateAppRibbon
    Exit Function
ErrorHandler:
    tier = ""
    InvalidateAppRibbon
    Dim answer As Integer
    answer = MsgBox("Failed to authenticate with finbox.io. Contact support@finbox.io if this problem persists.", vbCritical, "finbox.io Addin")
End Function

Public Function GetTier()
    If Not tier = "" Then
        GoTo OnError
    End If
    
    On Error GoTo OnError
    
    Dim APIkey As String
    
    tier = "anonymous"
    APIkey = GetAPIKey()
    
    Dim webClient As New webClient
    
    webClient.BaseUrl = TIER_URL
    
    Dim Auth As New HttpBasicAuthenticator
    Auth.Setup APIkey, ""
    Set webClient.Authenticator = Auth
    
    Dim webRequest As New webRequest
    webRequest.Method = WebMethod.HttpGet
    webRequest.ResponseFormat = WebFormat.Json
    webRequest.AddHeader "X-Finboxio-Addon", GetAPIHeader()
    
    Dim webResponse As webResponse
    Set webResponse = webClient.Execute(webRequest)

    tier = webResponse.data("data")("tier")
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


