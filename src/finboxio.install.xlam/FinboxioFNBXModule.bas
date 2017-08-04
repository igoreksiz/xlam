Attribute VB_Name = "FinboxioFNBXModule"
' finbox.io API Integration

Option Explicit

Public CheckedForUpdates As Boolean

Public Sub AddUDFCategoryDescription()
    #If Mac Then
        ' Excel for Mac does not support the property .MacroOptions
        Exit Sub
    #End If
    Application.MacroOptions Macro:="FNBX", Category:="finbox.io", _
        Description:="Returns a datapoint representing a selected company metric at a given point in time."
End Sub

Public Function FNBX(ByRef ticker As String, ByRef metric As String, Optional ByRef period = "") As Variant
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
        FNBX = cell.value
        GoTo Finish
    End If

    ' Check for updates on first use
    If Not CheckedForUpdates Then CheckUpdates
    CheckedForUpdates = True

    ' Check if user recently hit limit overage
    ' and refuse to request data if so
    If IsRateLimited Then
        FNBX = cell.value
        GoTo Finish
    End If

    ' Check for null arguments
    If IsEmpty(ticker) Or IsEmpty(metric) Then
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
        period = Application.Evaluate(period.address)
    End If
    If TypeName(period) = "Double" And period < (Now() - (365 * 50)) Then
        index = CInt(period)
        period = ""
    ElseIf TypeName(period) = "Double" Or TypeName(period) = "Date" Then
        period = "Y" & Year(period) & ".M" & Month(period) & ".D" & Day(period)
    End If

    ' Build finql key from arguments
    Dim key As String
    key = ticker & "." & metric
    If period <> "" Then key = key & "[""" & period & """]"
    
    ' If key is already cached, just return it
    If IsCached(key) Then
        FNBX = CachedToFNBX(key, index)
        GoTo Finish
    End If

    ' Check if user is logged in and prompt if not
    If Not IsLoggedIn() Then ShowLoginForm

    ' Collect all uncached keys to request
    Dim book As Workbook: Set book = cell.Worksheet.Parent
    Dim uncached() As String: uncached = FindUncachedKeys(book)
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
    
    If Err.Number = INVALID_ARGS_ERROR Then
        FNBX = CVErr(xlErrNum)
    ElseIf Err.Number = MISSING_VALUE_ERROR Then
        FNBX = CVErr(xlErrNull)
    Else
        FNBX = CVErr(xlErrValue)
    End If
    
    LogMessage "VBA error code " & Err.Number & " [" & Err.Description & "]", address
    
Finish:
    
End Function

' Return a list of all uncached finql keys required by a workbook
Public Function FindUncachedKeys(ByRef book As Workbook) As String()
    Dim keys() As String, uncached() As String
    ReDim keys(0)
    ReDim uncached(0)
    Dim i As Long, j As Long
    If Not book Is Nothing Then
        Dim fnd As String, range As range, cell As range, formula As String
        Dim Sheet As Worksheet
        For Each Sheet In book.Worksheets
            fnd = "FNBX("
            Set range = Sheet.UsedRange
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
                            If VBA.InStr(formulas(i, j), fnd) > 0 Then
                                formula = formulas(i, j)
                                Set cell = range.Cells(i, j)
                                ParseFormula formula, cell, Sheet, keys
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
                            ParseFormula formula, FoundCell, Sheet, keys
                        End If
                        If FoundCell.address = FirstFound Then Exit Do
                    Loop
                End If
                
                ' Reset the Find/Replace dialog after Find (not 100% sure this is necessary)
                ResetFindReplace
            #End If
        Next Sheet
    End If

    For i = 1 To UBound(keys)
        If Not IsCached(keys(i)) Then
            Call InsertElementIntoArray(uncached, UBound(uncached) + 1, keys(i))
        End If
    Next

    FindUncachedKeys = uncached()
End Function

Public Function CachedToFNBX(key As String, Optional index As Integer)
    If TypeName(GetCachedValue(key)) = "Collection" Then
        Dim list As Collection
        Set list = GetCachedValue(key)
        If TypeName(index) = "Empty" Then
            CachedToFNBX = CollectionToString(list)
        ElseIf list.count < index Then
            CachedToFNBX = CVErr(xlErrNull)
        Else
            CachedToFNBX = Val(index)
        End If
    Else
        CachedToFNBX = GetCachedValue(key)
    End If
End Function




