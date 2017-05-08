Attribute VB_Name = "FinboxioFNBXModule"
' finbox.io API Integration

Option Explicit

Public RedisplayDataLimit

Public Sub AddUDFCategoryDescription()
    #If Mac Then
        'Excel for Mac does not support the property .MacroOptions
        Exit Sub
    #End If
    Application.MacroOptions Macro:="FNBX", Category:="finbox.io", _
        Description:="Returns a datapoint representing a selected company metric at a given point in time."
End Sub

Public Function FNBX(ByRef ticker As String, ByRef metric As String, Optional ByRef period = "") As Variant
    On Error GoTo Error_Handler
    
    If IsReplacingLinks Then
        FNBX = CVErr(xlErrName)
        Exit Function
    End If

    Dim cell As String
    cell = CurrentCaller()
    
    ' check for null arguments
    If IsEmpty(ticker) Or IsEmpty(metric) Then
       FNBX = CVErr(xlErrNum) ' return #NUM!
       LogMessage "ticker.metric mal-formed.", ticker & "." & metric
       Exit Function
    End If
        
    ' build key from arguments
    Dim key As String
    
    key = ticker & "." & metric
    
    Dim pType As String: pType = TypeName(period)
    If pType = "Range" Then
        period = period.value
        pType = TypeName(period)
    End If
    
    If pType = "Double" Then
        period = ""
    ElseIf pType = "Date" Then
        period = "Y" & Year(period) & ".M" & Month(period) & ".D" & Day(period)
    End If
    
    If period <> "" Then
        key = key & "[""" & period & """]"
    End If
    
    ' check if (recent) key value is available in cache
    If IsCached(key) Then
        FNBX = GetCachedValue(key)
        Exit Function
    End If
    
    Dim loggedIn As Boolean
    If Not IsLoggedIn() Then
        ShowLoginForm
    End If
    
    ' check if user was recently notified of limit overage
    If TypeName(RedisplayDataLimit) = "Date" Then
        If RedisplayDataLimit > Now() Then
            FNBX = CVErr(xlErrNA)
            Exit Function
        Else
            RedisplayDataLimit = True
        End If
    End If
    
    ' check for null API key
    Dim APIkey As String
    APIkey = GetAPIKey()
    
    ' Add all uncached keys to request
    Dim i As Integer
    Dim k As String
    Dim escaped As String
    Dim allKeys() As String
    Dim requestedKeys() As String
    Dim added As Boolean
    
    ReDim requestedKeys(0)
    allKeys = FindAllKeys()
    
    For i = 1 To UBound(allKeys)
        k = allKeys(i)
        If Not IsCached(k) Then
            added = InsertElementIntoArray(requestedKeys, UBound(requestedKeys) + 1, k)
        End If
    Next

    Debug.Print "Building batch request for " & NumElements(requestedKeys) & " keys"
    
    Dim batchStart As Long: batchStart = 1
    Do While batchStart < NumElements(requestedKeys)
        ' build json request
        Dim jsonReqObj As Object
        Dim jsonDataObj As Object
        Dim batchKeys() As String
        Set jsonReqObj = ParseJson("{}")
        Set jsonDataObj = ParseJson("{}")
        
        ReDim batchKeys(0)
        For i = batchStart To Application.Min(NumElements(requestedKeys) - 1, batchStart + MAX_BATCH_SIZE)
            k = requestedKeys(i)
            escaped = EscapeQuotes(k)
            jsonDataObj(escaped) = k
            added = InsertElementIntoArray(batchKeys, UBound(batchKeys) + 1, k)
        Next
        batchStart = batchStart + MAX_BATCH_SIZE
        
        Set jsonReqObj("data") = jsonDataObj

        Dim postData As String
        postData = ConvertToJson(jsonReqObj)
    
        ' request json from web
        Dim webClient As New webClient
        
        webClient.BaseUrl = BATCH_URL
        
        Dim Auth As New HttpBasicAuthenticator
        Auth.Setup APIkey, "" ' api_key, password(not used)
    
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
        If Not webResponse.Data Is Nothing Then
            errStr = ConvertToJson(webResponse.Data("errors"), Whitespace:=2)
        End If
        
        If errStr <> "" Then LogMessage "errors: " & errStr
        
        If webResponse.statusCode = 429 Then
            DisplayDataLimit
            LogMessage "Finbox.io Data Limit Reached"
            FNBX = CVErr(xlErrNA)
            GoTo Exit_Function
        ' Return error if HTTP response code not 200
        ElseIf webResponse.statusCode >= 400 Or webResponse.Data Is Nothing Then
            LogMessage "The finbox.io API returned http status code " & webResponse.statusCode & " = " & _
                    VBA.Trim(webResponse.StatusDescription), key
    
            FNBX = CVErr(xlErrNA) ' return #N/A
            GoTo Exit_Function
        End If
        
        RedisplayDataLimit = True
        
        Dim resStr As String
        resStr = ConvertToJson(webResponse.Data("data"), Whitespace:=2)
        
        ' Extract data value from json response
        Dim dataVal As Variant
        
        For i = 1 To UBound(batchKeys)
            k = batchKeys(i)
            escaped = EscapeQuotes(k)
            If IsNull(webResponse.Data("data")(k)) Then
                Call SetCachedValue(k, CVErr(xlErrNull))
            Else
                If TypeName(webResponse.Data("data")(k)) = "Collection" Then
                    dataVal = CollectionToString(webResponse.Data("data")(k))
                Else
                    dataVal = webResponse.Data("data")(k)
                    If IsDate(dataVal) Then
                        dataVal = CDate(dataVal)
                    Else
                        If IsNumeric(dataVal) Then dataVal = CDbl(dataVal)
                    End If
                End If
                   
                Call SetCachedValue(k, dataVal)
            End If
        Next
    Loop
    
    ' key should now be cached
    If IsCached(key) Then
        FNBX = GetCachedValue(key)
    Else
        Debug.Print "key failed to cache " & key
        FNBX = CVErr(xlErrNull)
    End If
    
    GoTo Exit_Function
    
Error_Handler:
    FNBX = CVErr(xlErrValue) ' return #VALUE!
    
    LogMessage "VBA code error " & Err.Number & " [" & Err.Description & "]", key
    
Exit_Function:
    ' Clean up and exit
    Set jsonReqObj = Nothing
    Set jsonDataObj = Nothing
    Set webClient = Nothing
    Set webRequest = Nothing
    Set Auth = Nothing
    Set webResponse = Nothing
    
    On Error GoTo 0
End Function

Public Function FindAllKeys() As String()
    Dim fnd As String, FirstFound As String
    Dim FoundCell As Range, rng As Range
    Dim myRange As Range, LastCell As Range
    Dim formula As String
    Dim allKeys() As String
    Dim book As Workbook
    
    ReDim allKeys(0)
    
    Dim sheet As Worksheet
    Dim curSheet As String
    
    curSheet = ActiveSheet.name
    
    For Each book In Workbooks
        For Each sheet In book.Worksheets
            fnd = "FNBX("
            Set myRange = sheet.UsedRange
            #If Mac Then
                Dim cell As Range
                On Error Resume Next
                For Each cell In myRange
                    If cell.HasFormula Then
                        ParseKeys cell.formula, sheet, allKeys
                    End If
                Next cell
            #Else
                Set LastCell = myRange.Cells(myRange.Cells.count)
                Set FoundCell = myRange.Find(What:=fnd, LookIn:=xlFormulas, LookAt:=xlPart, After:=LastCell)
                If Not FoundCell Is Nothing Then
                    FirstFound = FoundCell.address
                    Set rng = FoundCell
                    On Error Resume Next
                    Do Until FoundCell Is Nothing
                        Set FoundCell = myRange.Find(What:=fnd, LookIn:=xlFormulas, LookAt:=xlPart, After:=FoundCell)
                        If cell.HasFormula Then
                            formula = FoundCell.formula
                            ParseKeys formula, sheet, allKeys
                        End If
                        If FoundCell.address = FirstFound Then Exit Do
                    Loop
                End If
            #End If
        Next sheet
    Next book
    Sheets(curSheet).Select
    Application.Run "ResetFindReplace"
    
    FindAllKeys = allKeys()
End Function

Sub ParseKeys(formula As String, sheet As Worksheet, ByRef keys)
    Dim argIndex As String: argIndex = VBA.InStr(formula, "(")
    If argIndex = 0 Then Exit Sub
    
    Dim name As String: name = VBA.Left(formula, argIndex - 1)
    Dim args() As String: args = GetParameters(formula)
    Dim argsCount As Long: argsCount = NumElements(args)
    
    If name = "FNBX" Or name = "=FNBX" Then
        Dim success As Boolean
        Dim ticker As String
        Dim metric As String
        Dim period
        
        ticker = EvalArgument(args(0), sheet)
        metric = EvalArgument(args(1), sheet)
        period = ""
        
        If argsCount > 2 Then
            period = EvalArgument(args(2), sheet)
            Dim pType As String: pType = TypeName(period)
            If pType = "Double" Then
                period = ""
            ElseIf pType = "Date" Then
                period = "Y" & Year(period) & ".M" & Month(period) & ".D" & Day(period)
            End If
        End If
        
        Dim key As String
        key = ticker & "." & metric
        If period <> "" Then
            key = key & "[""" & period & """]"
        End If
        
        success = InsertElementIntoArray(keys, UBound(keys) + 1, key)
    Else

    End If
End Sub

Function EvalArgument(arg As String, sheet As Worksheet)
    Dim value
    Dim address As String
    If ValidAddress(arg) Then
        address = sheet.Range(arg).address(External:=True)
        value = Range(address).value
        EvalArgument = value
    Else
        value = Application.Evaluate(arg)
        EvalArgument = value
    End If
End Function

Function GetParameters(func As String) As String()
    Dim args() As String
    Dim safeArgs As String
    Dim c As String
    Dim i As Long, pdepth As Long

    func = VBA.Trim(func)
    i = VBA.InStr(func, "(")
    func = VBA.Mid(func, i + 1)
    func = VBA.Mid(func, 1, VBA.Len(func) - 1)

    For i = 1 To VBA.Len(func)
        c = VBA.Mid(func, i, 1)
        If c = "(" Then
            pdepth = pdepth + 1
        ElseIf c = ")" Then
            pdepth = pdepth - 1
        ElseIf c = "," And pdepth = 0 Then
            c = "[[,]]"
        End If
        safeArgs = safeArgs & c
    Next i
    args = Split(safeArgs, "[[,]]")
    GetParameters = args
End Function
