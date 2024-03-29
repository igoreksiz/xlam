Attribute VB_Name = "FinboxioAPIModule"
Option Explicit
Option Private Module

Public Function GetErrorCode(name As String)
    If name = "InvalidKeyError" Then
        GetErrorCode = INVALID_KEY_ERROR
    ElseIf name = "InvalidPeriodError" Then
        GetErrorCode = INVALID_PERIOD_ERROR
    ElseIf name = "UnsupportedCompanyError" Then
        GetErrorCode = UNSUPPORTED_COMPANY_ERROR
    ElseIf name = "UnsupportedMetricError" Then
        GetErrorCode = UNSUPPORTED_METRIC_ERROR
    ElseIf name = "RestrictedCompanyError" Then
        GetErrorCode = RESTRICTED_COMPANY_ERROR
    ElseIf name = "RestrictedMetricError" Then
        GetErrorCode = RESTRICTED_METRIC_ERROR
    Else
        GetErrorCode = UNSPECIFIED_API_ERROR
    End If
End Function

Public Function RequestAndCacheKeys(ByRef keys() As String)
    Dim i As Integer, mock As String, k As String, ep As ErrorPoint, escaped As String, errs

    ' Remove duplicate keys
    Dim unique As New Dictionary
    For i = 1 To UBound(keys)
        ' If unique.Exists(keys(i)) Then LogMessage "Duplicate key " & keys(i)
        unique.Item(keys(i)) = 1
    Next

    LogMessage "Requesting " & NumElements(unique.keys) & " key(s)"

    ' Request all keys in batches smaller than MAX_BATCH_SIZE
    Dim batchStart As Long: batchStart = 0
    Do While batchStart <= NumElements(unique.keys)
        Dim jsonReqObj As Object
        Dim jsonDataObj As Object
        Dim batchKeys() As String
        Set jsonReqObj = ParseJson("{}")
        Set jsonDataObj = ParseJson("{}")

        ReDim batchKeys(0)
        For i = batchStart To Application.Min(NumElements(unique.keys) - 1, batchStart + MAX_BATCH_SIZE)
            k = "" & unique.keys(i)
            ' Allow mock status injection for testing
            If VBA.InStr(VBA.LCase(k), "x-mock-status") > 0 Then
                mock = VBA.Right(k, 3)
            Else
                escaped = EscapeQuotes(k)
                jsonDataObj.Item(escaped) = k
                Call InsertElementIntoArray(batchKeys, UBound(batchKeys) + 1, k)
            End If
            ' LogMessage "Requesting " & k
        Next
        batchStart = batchStart + MAX_BATCH_SIZE

        Set jsonReqObj.Item("data") = jsonDataObj

        Dim postData As String
        postData = ConvertToJson(jsonReqObj)

        Dim webClient As New webClient

        webClient.BaseUrl = BatchUrl
        webClient.TimeoutMs = 10000

        ' Setup Basic Auth with API key as username and empty password
        Dim APIKey As String: APIKey = GetAPIKey()
        If APIKey <> "" Then
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup APIKey, ""
            Set webClient.Authenticator = Auth
        End If

        Dim webRequest As New webRequest
        webRequest.Method = WebMethod.HttpPost
        webRequest.RequestFormat = WebFormat.Json
        webRequest.ResponseFormat = WebFormat.Json
        webRequest.Body = postData
        webRequest.AddHeader "X-Finbox-Addon", GetAPIHeader()
        
        If mock <> "" Then
            webRequest.AddHeader "X-Mock-Status", mock
        End If

        Dim webResponse As webResponse
        Set webResponse = webClient.Execute(webRequest)

        Dim QuotaUsed As Long, QuotaRemaining As Long
        For i = 1 To webResponse.Headers.count
            If VBA.LCase(webResponse.Headers(i).Item("Key")) = "x-quota-used" Then
                QuotaUsed = CLng(webResponse.Headers(i).Item("Value"))
            ElseIf VBA.LCase(webResponse.Headers(i).Item("Key")) = "x-quota-remaining" Then
                QuotaRemaining = CLng(webResponse.Headers(i).Item("Value"))
            End If
        Next i
        UpdateQuota QuotaUsed, QuotaRemaining

        ' Extract any error response
        Dim errStr As String
        If Not webResponse.Data Is Nothing Then
            errStr = ConvertToJson(webResponse.Data.Item("errors"), Whitespace:=2)
        End If

        ' If errStr <> "" Then LogMessage "errors: " & errStr

        If webResponse.statusCode = 429 Then
            If QuotaRemaining = 0 And QuotaUsed = 0 Then UpdateQuota 1, 0
            Err.Raise LIMIT_EXCEEDED_ERROR, "Data Limit Exceeded", "You must wait before making additional requests"
        ElseIf webResponse.statusCode >= 400 Or webResponse.Data Is Nothing Then
            Err.Raise UNSPECIFIED_API_ERROR, "API Response Error", "The API request returned " & webResponse.statusCode
        End If

        For i = 1 To UBound(batchKeys)
            k = batchKeys(i)
            Call SetCachedValue(k, ConvertValue(webResponse.Data.Item("data").Item(k)))
        Next
        
        If TypeName(webResponse.Data.Item("errors")) = "Collection" Then
            Set errs = webResponse.Data.Item("errors")
            For i = 1 To errs.count
                Set ep = New ErrorPoint
                ep.name = errs(i).Item("error")
                ep.code = GetErrorCode(ep.name)
                ep.description = errs(i).Item("description")
                Call SetCachedValue(errs(i).Item("id"), ep)
            Next
        End If
    Loop
End Function

Private Function ConvertValue(ByRef Data As Variant)
    If IsNull(Data) Then
        Dim nullValue As String
        nullValue = GetSetting("defaultNullValue", "0")
        If nullValue = "xlErrNull" Then
            Data = CVErr(xlErrNull)
            GoTo FinishConversion
        Else
            Data = nullValue
        End If
    End If
    
    If TypeName(Data) = "Collection" Then
        Dim i As Long, total As Long, converted As Variant
        total = Data.count
        For i = 1 To total
            converted = ConvertValue(Data(1))
            Data.Remove 1
            Data.Add converted
        Next
        Set ConvertValue = Data
        Exit Function
    ElseIf VBA.IsDate(Data) Then
        Data = CDate(Data)
    ElseIf TypeName(Data) = "String" Then
        Dim languageAdjusted As String
        languageAdjusted = AdjustForLanguage(CStr(Data))
        If IsNumeric(languageAdjusted) Then
            Data = CDbl(languageAdjusted)
        End If
    ElseIf TypeName(Data) = "Boolean" Then
        Data = Data
    ElseIf IsNumeric(Data) Then
        Data = CDbl(Data)
    Else
        Data = CVErr(xlErrValue)
    End If
    
FinishConversion:
    ConvertValue = Data
End Function

