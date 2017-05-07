Attribute VB_Name = "FinboxioAuthenticationModule"
' finbox.io API Integration

Option Explicit

Private APIKeyStore As APIKeyHandler

Public Sub SetAPIKeyHandler(handler As APIKeyHandler)
    Set APIKeyStore = handler
End Sub

Public Function ShowLoginForm()
    If EXCEL_VERSION = "Mac2011" Then
        MacCredentialsForm.Show
    Else
        CredentialsForm.Show
    End If
End Function

Public Function Login(ByVal email As String, ByVal password As String) As Boolean
    Login = False
    
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
            APItier = webResponse.Data("api_tier")
            APIkey = webResponse.Data("api_key")
            On Error GoTo ErrorHandler
            
            ' Process api_tier and api_key
            If APItier = "inactive" Then
                MsgBox "You have not verified your email address yet." & vbCrLf & _
                    "To resend the verification email, visit https://finbox.io/profile.", _
                    vbCritical, AppTitle
            Else
                APIKeyStore.StoreAPIKey (APIkey)
                Login = True
            End If
        Case Else
            MsgBox "The finbox.io API returned http status code " & webResponse.statusCode & " = " & vbCr & _
                VBA.Trim(webResponse.StatusDescription), vbCritical, AppTitle
    End Select

    Set jsonReqObj = Nothing
    Set webClient = Nothing
    Set webRequest = Nothing
    Set webResponse = Nothing
    
    AppRibbon.Invalidate
    Exit Function
ErrorHandler:
    Dim answer As Integer
    answer = MsgBox("Failed to authenticate with finbox.io. Contact support@finbox.io if this problem persists.", vbCritical, "finbox.io Addin")
    AppRibbon.Invalidate
End Function

Public Sub Logout()
    APIKeyStore.ClearAPIKey
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
    GetAPIKey = APIKeyStore.GetAPIKey()
End Function
