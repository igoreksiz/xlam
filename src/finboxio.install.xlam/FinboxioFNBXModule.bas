Attribute VB_Name = "FinboxioFNBXModule"
' finbox.io API Integration

Option Explicit

Public RedisplayDataLimit
Public CheckedForUpdates As Boolean

Public Sub AddUDFCategoryDescription()
    #If Mac Then
        ' Excel for Mac does not support the property .MacroOptions
        Exit Sub
    #End If
    Application.MacroOptions Macro:="FNBX", Category:="finbox.io", _
        Description:="Returns a datapoint representing a selected company metric at a given point in time."
End Sub

' TODO: Value casting and language adjustment is duplicated a few times. Consolidate this into a function or module.
' TODO: Move api request logic (including batch-splitting) into a separate module to make this function easier to follow.
Public Function FNBX(ByRef ticker As String, ByRef metric As String, Optional ByRef period = "") As Variant
    ' Must be marked volatile to enable recalculation on refresh
    Application.Volatile
    
    On Error GoTo Error_Handler
    
    ' Check for updates on first use
    If Not CheckedForUpdates Then
        CheckedForUpdates = True
        CheckUpdates
    End If

    ' Dont try to calculate during a link replacement.
    ' Link replacement converts FNBX references in
    ' external workbooks to local references.
    ' e.g. 'C:\...\finboxio.xlam'!FNBX => FNBX
    If IsReplacingLinks Then
        ' TODO: is it possible to just reuse existing
        ' value instead of putting error in cell?
        FNBX = CVErr(xlErrName)
        Exit Function
    End If

    Dim val As Variant
    Dim count As Integer
    Dim cell As String
    Dim languageAdjusted As String
    Dim pos As Long
    Dim char As String
    Dim numeric As String
    
    ' Get the address of the cell this was called from
    cell = CurrentCaller()

    ' Check for null arguments
    If IsEmpty(ticker) Or IsEmpty(metric) Then
       FNBX = CVErr(xlErrNum)
       LogMessage "ticker.metric mal-formed.", ticker & "." & metric
       Exit Function
    End If

    ' Build finql key from arguments
    Dim key As String
    Dim index As Integer

    key = ticker & "." & metric

    Dim pType As String: pType = TypeName(period)
    If pType = "Range" Then
        period = period.value
        pType = TypeName(period)
    End If

    If pType = "Double" Then
        index = CInt(period)
        period = ""
    ElseIf pType = "Date" Then
        period = "Y" & Year(period) & ".M" & Month(period) & ".D" & Day(period)
    End If

    If period <> "" Then
        key = key & "[""" & period & """]"
    End If

    ' Check if key value is available in cache
    If IsCached(key) Then
        If TypeName(GetCachedValue(key)) = "Collection" Then
            Set val = GetCachedValue(key)
        Else
            val = GetCachedValue(key)
        End If
        If pType = "Double" And TypeName(val) = "Collection" Then
            count = val.count
            If count < index Then
                val = CVErr(xlErrNull)
            Else
                val = val(index)
                If IsDate(val) Then
                    val = CDate(val)
                ElseIf TypeName(val) = "String" Then
                    numeric = "1234567890-.,"
                    languageAdjusted = ""
                    For pos = 1 To VBA.Len(val)
                        char = VBA.Mid(val, pos, 1)
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
                        val = CDbl(languageAdjusted)
                    End If
                End If
            End If
        ElseIf TypeName(val) = "Collection" Then
            val = CollectionToString(val)
        End If
        FNBX = val
        Exit Function
    End If

    ' Check if user recently hit limit overage
    ' and refuse to request data if so
    If TypeName(RedisplayDataLimit) = "Date" Then
        If RedisplayDataLimit > Now() Then
            FNBX = CVErr(xlErrNA)
            Exit Function
        Else
            RedisplayDataLimit = True
        End If
    End If

    ' Check if user is logged in and prompt if not
    Dim APIkey As String
    If Not IsLoggedIn() Then
        ShowLoginForm
    End If
    APIkey = GetAPIKey()

    ' Collect all uncached keys to request
    Dim i As Integer
    Dim k As String
    Dim escaped As String
    Dim allKeys() As String
    Dim requestedKeys() As String
    Dim added As Boolean

    Dim book As Workbook
    Dim cellRange As range
    Set cellRange = range(cell)
    Set book = cellRange.Worksheet.Parent
    
    ReDim requestedKeys(0)
    allKeys = FindAllKeys(book)
    added = InsertElementIntoArray(allKeys, UBound(allKeys) + 1, key)

    For i = 1 To UBound(allKeys)
        k = allKeys(i)
        If Not IsCached(k) Then
            added = InsertElementIntoArray(requestedKeys, UBound(requestedKeys) + 1, k)
        End If
    Next

    If (NumElements(requestedKeys) - 1) = 1 Then
        Debug.Print "Building batch request for " & requestedKeys(1) & " (" & cell & ")"
    Else
        Debug.Print "Building batch request for " & (NumElements(requestedKeys) - 1) & " keys (" & cell & ")"
    End If
    
    ' Request all keys in batches smaller than MAX_BATCH_SIZE
    Dim batchStart As Long: batchStart = 1
    Do While batchStart < NumElements(requestedKeys)
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

        Dim webClient As New webClient

        webClient.BaseUrl = BATCH_URL

        ' Setup Basic Auth with API key as username and empty password
        Dim Auth As New HttpBasicAuthenticator
        Auth.Setup APIkey, ""

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
            ' Notify user that they have hit their data limit and return #N/A error
            DisplayDataLimit
            LogMessage "Finbox.io Data Limit Reached"
            FNBX = CVErr(xlErrNA)
            GoTo Exit_Function
        ElseIf webResponse.statusCode >= 400 Or webResponse.Data Is Nothing Then
            ' Log unspecified errors and return #N/A
            LogMessage "The finbox.io API returned http status code " & webResponse.statusCode & " = " & VBA.Trim(webResponse.StatusDescription), key
            FNBX = CVErr(xlErrNA)
            GoTo Exit_Function
        End If

        ' Clear data-limit block on successful request
        RedisplayDataLimit = True

        Dim dataVal As Variant
        
        For i = 1 To UBound(batchKeys)
            k = batchKeys(i)
            If IsNull(webResponse.Data("data")(k)) Then
                ' Cache null value
                Call SetCachedValue(k, CVErr(xlErrNull))
            Else
                ' Cast value to appropriate type and cache
                If TypeName(webResponse.Data("data")(k)) = "Collection" Then
                    Set dataVal = webResponse.Data("data")(k)
                Else
                    dataVal = webResponse.Data("data")(k)
                End If
                
                If IsDate(dataVal) Then
                    dataVal = CDate(dataVal)
                ElseIf TypeName(dataVal) = "String" Then
                    numeric = "1234567890-.,"
                    languageAdjusted = ""
                    For pos = 1 To VBA.Len(dataVal)
                        char = VBA.Mid(dataVal, pos, 1)
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
                        dataVal = CDbl(languageAdjusted)
                    End If
                End If
                Call SetCachedValue(k, dataVal)
            End If
        Next
    Loop

    ' Key should now be cached, so just lookup and return
    If IsCached(key) Then
        If TypeName(GetCachedValue(key)) = "Collection" Then
            Set val = GetCachedValue(key)
        Else
            val = GetCachedValue(key)
        End If
        
        If pType = "Double" And TypeName(val) = "Collection" Then
            ' For formulas that request list item at specific index [=FNBX("AAPL","benchmarks",1)]
            ' Lists are indexed starting at position 1! (Not 0)
            count = val.count
            If count < index Then
                val = CVErr(xlErrNull)
            Else
                ' Cast individual list items to proper type
                val = val(index)
                If IsDate(val) Then
                    val = CDate(val)
                ElseIf TypeName(val) = "String" Then
                    numeric = "1234567890-.,"
                    languageAdjusted = ""
                    For pos = 1 To VBA.Len(val)
                        char = VBA.Mid(val, pos, 1)
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
                        val = CDbl(languageAdjusted)
                    End If
                End If
            End If
        ElseIf TypeName(val) = "Collection" Then
            ' For formulas that request a list without specifying an index [=FNBX("AAPL","benchmarks")]
            val = CollectionToString(val)
        End If
        FNBX = val
    Else
        ' For some reason this value was not found in the cache.
        ' Generally, this indicates some problem since even
        ' unsupported keys should get cached as null.
        FNBX = CVErr(xlErrNull)
    End If

    GoTo Exit_Function

Error_Handler:
    ' Log unspecified errors and return #VALUE!
    FNBX = CVErr(xlErrValue)
    LogMessage "VBA error from cell " & cell & ": " & Err.Number & " [" & Err.Description & "]", key

Exit_Function:
    ' Do any cleanup here
End Function

' Return a list of all finql keys required by a workbook
Public Function FindAllKeys(ByRef book As Workbook) As String()
    Dim fnd As String, range As range, cell As range, formula As String
    Dim allKeys() As String

    ReDim allKeys(0)

    Dim sheet As Worksheet
    If Not book Is Nothing Then
        For Each sheet In book.Worksheets
            fnd = "FNBX("
            Set range = sheet.UsedRange
            #If Mac Then
                ' VBA on Mac does not allow us to use Find while running in the context
                ' of a UDF. So we have to iterate all the cells and check for FNBX. A
                ' couple of optimizations are important for making this usable with very
                ' large sheets. First, load the 2D array of formulas from the entire range
                ' instead of individual cells. Second, use SpecialCells to reduce the range
                ' to only include cells with formulas.
                Dim formulas As Variant
                On Error Resume Next
                formulas = range.SpecialCells(xlCellTypeFormulas).formula
                Dim i As Long, j As Long
                For i = LBound(formulas, 1) To UBound(formulas, 1)
                    For j = LBound(formulas, 2) To UBound(formulas, 2)
                        If Not formulas(i, j) = "" Then
                            If VBA.InStr(formulas(i, j), fnd) > 0 Then
                                formula = formulas(i, j)
                                Set cell = range.Cells(i, j)
                                ParseFormula formula, cell, sheet, allKeys
                            End If
                        End If
                    Next j
                Next i
            #Else
                ' On Windows, we can do a search for all FNBX cells. Generally, this gives
                ' better performance because we don't have to iterate through all cells in
                ' the workbook. However, we are accessing each formula on individual cells
                ' which could be more expensive than loading all formulas for a range in a
                ' single call. So there may be a need to optimize this call for very large
                ' workbooks with a high ratio of FNBXs-to-cells (TODO: test performance
                ' trade-off for workbooks with increasing FNBX count)
                Dim FirstFound As String, LastCell As range, FoundCell As range
                Set LastCell = range.Cells(range.Cells.count)
                Set FoundCell = range.Find(What:=fnd, LookIn:=xlFormulas, LookAt:=xlPart, After:=LastCell)
                If Not FoundCell Is Nothing Then
                    FirstFound = FoundCell.address
                    On Error Resume Next
                    Do Until FoundCell Is Nothing
                        Set FoundCell = range.Find(What:=fnd, LookIn:=xlFormulas, LookAt:=xlPart, After:=FoundCell)
                        If FoundCell.HasFormula Then
                            formula = FoundCell.formula
                            ParseFormula formula, FoundCell, sheet, allKeys
                        End If
                        If FoundCell.address = FirstFound Then Exit Do
                    Loop
                End If
                
                ' Reset the Find/Replace dialog after Find (not 100% sure this is necessary)
                Application.Run "ResetFindReplace"
            #End If
        Next sheet
    End If

    FindAllKeys = allKeys()
End Function
