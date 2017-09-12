Attribute VB_Name = "FinboxioFNBXModule"
Option Explicit

Public CheckedForUpdates As Boolean
Public RequestedLogin As Boolean

Public Sub AddUDFCategoryDescription()
    On Error Resume Next
    #If Mac Then
        ' Excel for Mac does not support the property .MacroOptions
        Exit Sub
    #ElseIf VBA7 Then
        Application.MacroOptions _
            Macro:="FNBX", _
            Category:=1, _
            description:="Returns a datapoint representing a selected company metric at a given point in time.", _
            StatusBar:="Downloading data from finbox.io", _
            HelpFile:=HELP_URL, _
            ArgumentDescriptions:=Array( _
                "The company's ticker or finbox.io ID (e.g. AAPL or FINBOX:AAPL). Visit finbox.io to see a complete list of supported companies.", _
                "The metric ID for the data you wish to retrieve (e.g revenue or roic). See the finbox.io data explorer for a complete list of supported metrics.", _
                "The period of the data you want to retrieve. Accepts a date or finbox.io period string (e.g. 2017-03-01 or FY-1). This parameter is optional and may not be supported for some metrics." _
            )
    #Else
        Application.MacroOptions _
            Macro:="FNBX", _
            Category:=1, _
            description:="Returns a datapoint representing a selected company metric at a given point in time.", _
            StatusBar:="Downloading data from finbox.io", _
            HelpFile:=HELP_URL
    #End If
End Sub

Public Function FNBX(ByRef ticker As String, ByRef metric As String, Optional ByRef period = "") As Variant
Attribute FNBX.VB_Description = "Returns a datapoint representing a selected company metric at a given point in time."
Attribute FNBX.VB_ProcData.VB_Invoke_Func = " \n1"
    ' Must be marked volatile to enable recalculation on refresh
    Application.Volatile
    
    On Error GoTo HandleErrors

    ' Get the address of the cell this was called from
    Dim address As String, cell As Range
    address = CurrentCaller()
    Set cell = Range(address)

    ' Dont try to calculate during a link replacement.
    ' Link replacement converts FNBX references in
    ' external workbooks to local references.
    ' e.g. 'C:\...\finboxio.xlam'!FNBX => FNBX
    If IsReplacingLinks Then
        FNBX = CVErr(xlErrName)
        GoTo Finish
    End If

    ' Promote any staged manager update
    PromoteStagedUpdate

    ' Check for null arguments
    If IsEmpty(ticker) Or ticker = "" Or IsEmpty(metric) Or metric = "" Then
       Err.Raise INVALID_ARGS_ERROR, "Invalid Arguments Error", "The FNBX function requires a ticker and a metric"
    End If

    ' Resolve period argument and set list index if provided.
    '
    ' Note:
    ' Determining if the index param is a date period or index
    ' is somewhat dubious because formulas like TODAY() pass in
    ' numbers to the formula. For now, we assume that if a number
    ' is passed in that represents a date in the past 50 years, it
    ' should be treated like a Date. Otherwise it's a list index.
    ' This assumption should be valid unless we start supporting
    ' more than 50 years of data or design metrics that might
    ' return lists longer than 20k+ items
    '
    ' P.S.
    ' This code is (sort of) duplicated in ParseKeys, so
    ' if you change this, check that function as well
    
    Dim index As Integer
    If TypeName(period) = "Range" Then
        period = Application.Evaluate(period.address(External:=True))
    End If
    If TypeName(period) = "Double" And period < (Now() - (365 * 50)) Then
        index = CInt(period)
        period = ""
    ElseIf TypeName(period) = "Double" Or TypeName(period) = "Date" Then
        period = "Y" & Year(period) & ".M" & Month(period) & ".D" & Day(period)
    ElseIf TypeName(period) = "String" And IsDateString(CStr(period)) Then
        period = DateStringToPeriod(CStr(period))
    End If

    ' Build finql key from arguments
    Dim key As String
    key = VBA.UCase(ticker) & "." & VBA.LCase(metric)
    If period <> "" Then key = key & "[""" & VBA.UCase(period) & """]"
    
    ' If key is already cached, just return it
    If IsCached(key) Then
        FNBX = CachedToFNBX(key, index)
        GoTo Finish
    End If

    ' In some versions of excel, formulas may be called with incomplete
    ' arguments when the workbook is first loading. For example, =FNBX(A1,A2,A3)
    ' may be called with an empty period if A3 is a formula that hasn't
    ' been calculated yet. This causes unnecessary API requests and can
    ' significantly hurt load performance. So if we get a key that doesn't
    ' have a period, we check here to see if the formula in this cell
    ' actually does include a period. If so, we simply return an error since
    ' this cell will be recalculated once all arguments are resolved.
    If period = "" Then
        Dim cellKeys() As String, ik As Integer, sameKey As Boolean
        ReDim cellKeys(0)
        Call ParseFormula(cell.Formula, cell, cell.Worksheet, cellKeys)
        For ik = 1 To UBound(cellKeys)
            If key = cellKeys(ik) Then sameKey = True
        Next ik
        If Not sameKey Then
            FNBX = CVErr(xlErrValue)
            GoTo Finish
        End If
    End If

    ' Check if user is logged in and show form if not
    If Not IsLoggedIn() And Not RequestedLogin Then ShowLoginForm
    RequestedLogin = True

    ' Check if user recently hit limit overage
    ' and refuse to request data if so
    If IsRateLimited Then
        FNBX = CVErr(xlErrNA)
        GoTo Finish
    End If

    ' Collect all uncached keys to request
    Dim book As Workbook: Set book = cell.Worksheet.Parent
    ' LogMessage "Parsing Keys"
    Dim uncached() As String: uncached = FindUncachedKeys(book)
    ' LogMessage "Parsed Keys"
    Call InsertElementIntoArray(uncached, UBound(uncached) + 1, key)

    ' Request and cache keys
    Call RequestAndCacheKeys(uncached)

    ' Key should now be cached, so just lookup and return
    If IsCached(key) Then
        FNBX = CachedToFNBX(key, index)
    Else
        ' For some reason this value was not found in the cache.
        ' Generally, this indicates some problem since even
        ' unsupported/empty keys should get cached as null or error
        Err.Raise MISSING_VALUE_ERROR, "Missing Value Error", "Could not find " & key & " in the cache"
    End If

    GoTo Finish

HandleErrors:
    If Err.Number = LIMIT_EXCEEDED_ERROR Then
        ShowRateLimitWarning
    End If

    If Err.Number = MISSING_VALUE_ERROR Then
        FNBX = CVErr(xlErrNull)
    ElseIf Err.Number = INVALID_ARGS_ERROR Then
        FNBX = CVErr(xlErrValue)
    ElseIf Err.Number = INVALID_KEY_ERROR Then
        FNBX = CVErr(xlErrValue)
    ElseIf Err.Number = INVALID_PERIOD_ERROR Then
        FNBX = CVErr(xlErrValue)
    ElseIf Err.Number = UNSUPPORTED_METRIC_ERROR Then
        FNBX = CVErr(xlErrValue)
    ElseIf Err.Number = UNSUPPORTED_COMPANY_ERROR Then
        FNBX = CVErr(xlErrValue)
    ElseIf Err.Number = RESTRICTED_COMPANY_ERROR Then
        FNBX = CVErr(xlErrNA)
    ElseIf Err.Number = RESTRICTED_METRIC_ERROR Then
        FNBX = CVErr(xlErrNA)
    Else
        FNBX = CVErr(xlErrNA)
    End If
    
    LogMessage "VBA error code " & Err.Number & " [" & Err.description & "]", address
    
Finish:

End Function

' Return a list of all uncached finql keys required by a workbook
Private Function FindUncachedKeys(ByRef book As Workbook) As String()
    Dim keys() As String, uncached() As String
    ReDim keys(0)
    ReDim uncached(0)
    Dim i As Long, j As Long
    If Not book Is Nothing Then
        Dim fnd As String, rng As Range, cell As Range, Formula As String
        Dim sheet As Worksheet
        For Each sheet In book.Worksheets
            fnd = "FNBX("
            Set rng = sheet.UsedRange
            #If Mac Then
                ' VBA on Mac does not allow us to use Find while running in the context
                ' of a UDF. So we have to iterate all the cells and check for FNBX. A
                ' couple of optimizations are important for making this usable with very
                ' large sheets. First, load the 2D array of formulas from the entire range
                ' instead of individual cells. Second, use SpecialCells to reduce the range
                ' to only include cells with formulas.
                Dim formulas As Variant
                On Error Resume Next
                formulas = rng.SpecialCells(xlCellTypeFormulas).Formula
                For i = LBound(formulas, 1) To UBound(formulas, 1)
                    For j = LBound(formulas, 2) To UBound(formulas, 2)
                        If Not formulas(i, j) = "" Then
                            If VBA.InStr(VBA.UCase(formulas(i, j)), fnd) > 0 Then
                                Formula = formulas(i, j)
                                Set cell = rng.Cells(i, j)
                                Call ParseFormula(Formula, cell, sheet, keys)
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
                Dim FirstFound As String, LastCell As Range, FoundCell As Range
                Set LastCell = rng.Cells(rng.Cells.count)
                Set FoundCell = rng.Find(What:=fnd, LookIn:=xlFormulas, LookAt:=xlPart, After:=LastCell, MatchCase:=False)
                If Not FoundCell Is Nothing Then
                    FirstFound = FoundCell.address
                    On Error Resume Next
                    Do Until FoundCell Is Nothing
                        Set FoundCell = rng.Find(What:=fnd, LookIn:=xlFormulas, LookAt:=xlPart, After:=FoundCell, MatchCase:=False)
                        If FoundCell.HasFormula Then
                            Formula = FoundCell.Formula
                            Call ParseFormula(Formula, FoundCell, sheet, keys)
                        End If
                        If FoundCell.address = FirstFound Then Exit Do
                    Loop
                End If
                
                ' Reset the Find/Replace dialog after Find (not 100% sure this is necessary)
                ResetFindReplace
            #End If
        Next sheet
    End If

    For i = 1 To UBound(keys)
        If Not IsCached(keys(i)) Then
            Call InsertElementIntoArray(uncached, UBound(uncached) + 1, keys(i))
        End If
    Next

    FindUncachedKeys = uncached()
End Function

