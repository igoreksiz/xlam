Attribute VB_Name = "FinboxioFNBXModule"
Option Explicit

Public CheckedForUpdates As Boolean
Public RequestedLogin As Boolean

Public Sub AddUDFCategoryDescription()
    #If Mac Then
        ' Excel for Mac does not support the property .MacroOptions
        Exit Sub
    #End If
    Application.MacroOptions Macro:="FNBX", Category:="finbox.io", _
        description:="Returns a datapoint representing a selected company metric at a given point in time."
End Sub

Public Function FNBX(ByRef ticker As String, ByRef metric As String, Optional ByRef period = "") As Variant
Attribute FNBX.VB_Description = "Returns a datapoint representing a selected company metric at a given point in time."
Attribute FNBX.VB_ProcData.VB_Invoke_Func = " \n19"
    ' Must be marked volatile to enable recalculation on refresh
    Application.Volatile
    
    On Error GoTo HandleErrors

    ' Get the address of the cell this was called from
    Dim address As String, cell As range
    address = CurrentCaller()
    Set cell = range(address)

    ' Dont try to calculate during a link replacement.
    ' Link replacement converts FNBX references in
    ' external workbooks to local references.
    ' e.g. 'C:\...\finboxio.xlam'!FNBX => FNBX
    If IsReplacingLinks Then
        FNBX = CVErr(xlErrName)
        GoTo Finish
    End If

    ' Check for updates on first use
    ' If Not CheckedForUpdates Then CheckUpdates
    ' CheckedForUpdates = True

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
        Call ParseFormula(cell.formula, cell, cell.Worksheet, cellKeys)
        For ik = 1 To UBound(cellKeys)
            If key = cellKeys(ik) Then sameKey = True
        Next ik
        If Not sameKey Then
            FNBX = CVErr(xlErrValue)
            GoTo Finish
        End If
    End If

    ' Check if user is logged in and prompt if not
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
        Dim fnd As String, range As range, cell As range, formula As String
        Dim sheet As Worksheet
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
                For i = LBound(formulas, 1) To UBound(formulas, 1)
                    For j = LBound(formulas, 2) To UBound(formulas, 2)
                        If Not formulas(i, j) = "" Then
                            If VBA.InStr(VBA.UCase(formulas(i, j)), fnd) > 0 Then
                                formula = formulas(i, j)
                                Set cell = range.Cells(i, j)
                                Call ParseFormula(formula, cell, sheet, keys)
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
                Set FoundCell = range.Find(What:=fnd, LookIn:=xlFormulas, LookAt:=xlPart, After:=LastCell, MatchCase:=False)
                If Not FoundCell Is Nothing Then
                    FirstFound = FoundCell.address
                    On Error Resume Next
                    Do Until FoundCell Is Nothing
                        Set FoundCell = range.Find(What:=fnd, LookIn:=xlFormulas, LookAt:=xlPart, After:=FoundCell, MatchCase:=False)
                        If FoundCell.HasFormula Then
                            formula = FoundCell.formula
                            Call ParseFormula(formula, FoundCell, sheet, keys)
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

