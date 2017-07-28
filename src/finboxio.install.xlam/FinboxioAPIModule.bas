Attribute VB_Name = "FinboxioAPIModule"
Option Explicit

Public Function RequestAndCacheKeys(ByRef keys() As String)
    Dim i As Integer, k As String, escaped As String

    ' Remove duplicate keys
    Dim unique As New Dictionary
    For i = 1 To UBound(keys)
        unique(keys(i)) = 1
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
            escaped = EscapeQuotes(k)
            jsonDataObj(escaped) = k
            Call InsertElementIntoArray(batchKeys, UBound(batchKeys) + 1, k)
            ' LogMessage "Requesting " & k
        Next
        batchStart = batchStart + MAX_BATCH_SIZE

        Set jsonReqObj("data") = jsonDataObj

        Dim postData As String
        postData = ConvertToJson(jsonReqObj)

        Dim webClient As New webClient

        webClient.BaseUrl = BATCH_URL

        ' Setup Basic Auth with API key as username and empty password
        Dim Auth As New HttpBasicAuthenticator
        Auth.Setup GetAPIKey(), ""

        Set webClient.Authenticator = Auth

        Dim webRequest As New webRequest
        webRequest.Method = WebMethod.HttpPost
        webRequest.RequestFormat = WebFormat.Json
        webRequest.ResponseFormat = WebFormat.Json
        webRequest.Body = postData
        webRequest.AddHeader "X-Finboxio-Addon", GetAPIHeader()

        Dim webResponse As webResponse
        Set webResponse = webClient.Execute(webRequest)

        ' Extract any error response
        Dim errStr As String
        If Not webResponse.data Is Nothing Then
            errStr = ConvertToJson(webResponse.data("errors"), Whitespace:=2)
        End If

        If errStr <> "" Then LogMessage "errors: " & errStr

        If webResponse.statusCode = 429 Then
            Err.Raise LIMIT_EXCEEDED_ERROR, "Data Limit Exceeded", "You must wait before making additional requests"
        ElseIf webResponse.statusCode >= 400 Or webResponse.data Is Nothing Then
            Err.Raise UNSPECIFIED_API_ERROR, "API Response Error", "The API request returned " & webResponse.statusCode
        End If

        For i = 1 To UBound(batchKeys)
            k = batchKeys(i)
            Call SetCachedValue(k, ConvertValue(webResponse.data("data")(k)))
        Next
    Loop
End Function

Private Function ConvertValue(ByRef data As Variant)
    If IsNull(data) Then
        data = CVErr(xlErrNull)
    ElseIf TypeName(data) = "Collection" Then
        Dim i As Long, total As Long, converted As Variant
        total = data.count
        For i = 1 To total
            converted = ConvertValue(data(1))
            data.Remove 1
            data.Add converted
        Next
        Set ConvertValue = data
        Exit Function
    ElseIf IsDate(data) Then
        data = CDate(data)
    ElseIf TypeName(data) = "String" Then
        Dim numeric As String, char As String, pos As Long, languageAdjusted As String
        numeric = "1234567890-.,"
        languageAdjusted = ""
        For pos = 1 To VBA.Len(data)
            char = VBA.Mid(data, pos, 1)
            If VBA.InStr(numeric, char) = 0 Then
                languageAdjusted = "x"
                Exit For
            ElseIf char = "," Then
                languageAdjusted = languageAdjusted & Application.International(xlThousandsSeparator)
            ElseIf char = "." Then
                languageAdjusted = languageAdjusted & Application.International(xlDecimalSeparator)
            Else
                languageAdjusted = languageAdjusted & char
            End If
        Next
        If IsNumeric(languageAdjusted) Then
            data = CDbl(languageAdjusted)
        End If
    ElseIf TypeName(data) = "Boolean" Then
        data = data
    ElseIf IsNumeric(data) Then
        data = CDbl(data)
    Else
        data = CVErr(xlErrValue)
    End If
    ConvertValue = data
End Function
