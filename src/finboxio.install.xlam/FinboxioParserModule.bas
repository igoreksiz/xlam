Attribute VB_Name = "FinboxioParserModule"
Option Explicit

' Locate all FNBX formulas in a string and evaluate required keys for each
Sub ParseFormula(formula As String, cell As range, Sheet As Worksheet, ByRef keys)
    Dim fn As String: fn = ""
    Dim quotes As Boolean: quotes = False
    Dim inFNBX As Long: inFNBX = 0
    Dim parens As Long: parens = 0
    Dim i As Long
    For i = 1 To VBA.Len(formula)
        Dim char As String
        char = VBA.Mid(formula, i, 1)
        If char = """" Then
            quotes = Not quotes
            If VBA.Len(fn) > 0 Then
                fn = fn & char
            End If
        ElseIf quotes Then
            If VBA.Len(fn) > 0 Then
                fn = fn & char
            End If
        ElseIf inFNBX = 4 And char = "(" Then
            parens = parens + 1
            fn = fn & char
        ElseIf inFNBX = 4 And char = ")" Then
            parens = parens - 1
            fn = fn & char
            If parens = 0 Then
                ParseKeys fn, cell, Sheet, keys
                fn = ""
                inFNBX = 0
            End If
        ElseIf parens = 0 And inFNBX = 0 And char = "F" Then
            fn = fn & char
            inFNBX = 1
        ElseIf parens = 0 And inFNBX = 1 And char = "N" Then
            fn = fn & char
            inFNBX = 2
        ElseIf parens = 0 And inFNBX = 2 And char = "B" Then
            fn = fn & char
            inFNBX = 3
        ElseIf parens = 0 And inFNBX = 3 And char = "X" Then
            fn = fn & char
            inFNBX = 4
        ElseIf inFNBX = 4 And parens > 0 Then
            fn = fn & char
        Else
            fn = ""
            inFNBX = 0
        End If
    Next i
End Sub

' Determine all finql keys required by a FNBX formula
Sub ParseKeys(formula As String, cell As range, Sheet As Worksheet, ByRef keys)
    Dim argIndex As String: argIndex = VBA.InStr(formula, "(")
    If argIndex = 0 Then Exit Sub

    Dim name As String: name = VBA.Left(formula, argIndex - 1)
    Dim args() As String: args = ParseArguments(formula)
    Dim argsCount As Long: argsCount = NumElements(args)

    If name = "FNBX" Or name = "=FNBX" Or name = "=-FNBX" Then
        Dim success As Boolean
        Dim ticker As String
        Dim metric As String
        Dim activated As Boolean
        Dim nested As Boolean
        Dim period
        
        ' Test each argument for nested FNBX formulas
        ' and parse only the nested formulas since
        ' these must be resolved before we can determine
        ' the key for the current formula
        
        nested = False
        If argsCount > 0 Then
            If VBA.InStr(args(0), "FNBX(") > 0 Then
                ParseFormula args(0), cell, Sheet, keys
                nested = True
            End If
        End If
        
        If argsCount > 1 Then
            If VBA.InStr(args(1), "FNBX(") > 0 Then
                ParseFormula args(1), cell, Sheet, keys
                nested = True
            End If
        End If
        
        If argsCount > 2 Then
            If VBA.InStr(args(2), "FNBX(") > 0 Then
                ParseFormula args(2), cell, Sheet, keys
                nested = True
            End If
        End If
        
        If argsCount > 3 Then
            If VBA.InStr(args(3), "FNBX(") > 0 Then
                ParseFormula args(3), cell, Sheet, keys
                nested = True
            End If
        End If
        
        If nested Then Exit Sub
        
        ' Build the finql key required by the formula.
        ' This code is (sort of) duplicated in FNBX, so
        ' if you change this, check that function as well
        
        ticker = EvalArgument(args(0), cell, Sheet)
        metric = EvalArgument(args(1), cell, Sheet)
        period = ""

        If argsCount > 2 Then
            period = EvalArgument(args(2), cell, Sheet)
            If TypeName(period) = "Double" And period < (Now() - (365 * 50)) Then
                period = ""
            ElseIf TypeName(period) = "Double" Or TypeName(period) = "Date" Then
                period = "Y" & Year(period) & ".M" & Month(period) & ".D" & Day(period)
            End If
        End If

        Dim key As String
        key = ticker & "." & metric
        If period <> "" Then
            key = key & "[""" & period & """]"
        End If

        ' Add resolved key to list of keys to request
        success = InsertElementIntoArray(keys, UBound(keys) + 1, key)
    End If
End Sub

' Parse a list of argument strings given an excel formula
Function ParseArguments(func As String) As String()
    Dim args() As String
    Dim safeArgs As String
    Dim c As String
    Dim i As Long, pdepth As Long
    Dim quoted As Boolean

    quoted = False
    func = VBA.Trim(func)
    i = VBA.InStr(func, "(")
    func = VBA.Mid(func, i + 1)
    func = VBA.Mid(func, 1, VBA.Len(func) - 1)

    ' Escape any commas in nested formulas or quotes
    For i = 1 To VBA.Len(func)
        c = VBA.Mid(func, i, 1)
        If c = "(" Then
            pdepth = pdepth + 1
        ElseIf c = ")" Then
            pdepth = pdepth - 1
        ElseIf c = """" Then
            quoted = Not quoted
        ElseIf c = Application.International(xlListSeparator) And pdepth = 0 And Not quoted Then
            c = "[[,]]"
        End If
        safeArgs = safeArgs & c
    Next i
    args = Split(safeArgs, "[[,]]")
    ParseArguments = args
End Function

' Evaluate the value of an argument that may include
' formulas or cell references
Function EvalArgument(arg As String, cell As range, Sheet As Worksheet)
    Dim value
    Dim address As String
    If IsCellAddress(arg) Then
        ' Evaluate reference to another sheet/cell
        Dim parts
        Dim sheetName As String
        Dim cellAddr As String
        
        parts = VBA.Split(arg, "!")
        If (NumElements(parts) > 1) Then
            sheetName = parts(0)
            cellAddr = parts(1)
        Else
            sheetName = Sheet.name
            cellAddr = parts(0)
        End If
            
        address = Sheet.Parent.Sheets(sheetName).range(cellAddr).address(External:=True)
        value = range(address).value
        EvalArgument = value
    ElseIf IsTableAddress(arg) Then
        ' Evaluate reference to a table cell
        value = EvalTableAddress(arg, cell)
        EvalArgument = value
    Else
        ' Evaluate nested formula arg or return constant value arg
        value = Application.Evaluate(arg)
        EvalArgument = value
    End If
End Function

' Determine if string represents a valid excel cell address
Public Function IsCellAddress(strAddress As String) As Boolean
    Dim r As range
    On Error Resume Next
    Set r = range(strAddress)
    If Not r Is Nothing Then IsCellAddress = True
End Function

' Determine if argument represents a valid reference to a table address
Function IsTableAddress(arg As String) As Boolean
    Dim i As Long
    Dim c As String
    Dim result As Boolean
    result = False
    
    For i = 1 To VBA.Len(arg)
        c = VBA.Mid(arg, i, 1)
        If c = """" Then
            result = False
            i = VBA.Len(arg)
        ElseIf c = "[" Then
            result = True
            i = VBA.Len(arg)
        End If
    Next i
    
    IsTableAddress = result
End Function

' Evaluate the value at a particular table address
' relative to the given cell
Function EvalTableAddress(arg As String, cell As range)
    Dim i As Long
    Dim j As Long
    Dim c As String
    Dim table As ListObject
    
    i = VBA.InStr(arg, "[")
    If (i = 1) Then
        Set table = cell.ListObject
    Else
        Dim name As String: name = VBA.Left(arg, i)
        Set table = cell.Worksheet.ListObjects(name)
    End If
    
    j = VBA.InStr(i, arg, "]")
    Dim header As String
    header = VBA.Mid(arg, i + 1, j - i - 1)
    header = VBA.Replace(header, "@", "")
    
    Dim row As Long
    Dim first As range
    Set first = table.DataBodyRange.Cells(1, 1)
    row = first.row - 1
    
    Dim row2 As Long
    row2 = cell.row
    
    Dim focus
    EvalTableAddress = table.DataBodyRange(row2 - row, table.ListColumns(header).index)
End Function

