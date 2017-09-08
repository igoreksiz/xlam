Attribute VB_Name = "FinboxioUtilitiesModule"
Option Explicit
Option Private Module

Public Function CollectionToString(ByVal dataCol As Variant) As String
    Dim i As Integer, val As String, sep As String
    sep = Application.International(xlListSeparator)
    For i = 1 To dataCol.count
        If CollectionToString <> "" Then
            CollectionToString = CollectionToString & sep
        End If
        val = CStr(dataCol(i))
        If VBA.InStr(val, sep) > 0 Then val = """" & EscapeQuotes(val) & """"
        CollectionToString = CollectionToString & val
    Next i
End Function

Public Sub ResetFindReplace()
   'Resets the find/replace dialog box options
   Dim r As Range

   On Error Resume Next

   Set r = Cells.Find(What:="", _
   LookIn:=xlFormulas, _
   SearchOrder:=xlRows, _
   LookAt:=xlPart, _
   MatchCase:=False)

   On Error GoTo 0

   'Reset the defaults

   On Error Resume Next

   Set r = Cells.Find(What:="", _
   LookIn:=xlFormulas, _
   SearchOrder:=xlRows, _
   LookAt:=xlPart, _
   MatchCase:=False)

   On Error GoTo 0
End Sub

Public Function EscapeQuotes(str As String) As String
    EscapeQuotes = Replace(str, """", "\""")
End Function

Public Function DescapeQuotes(str As String) As String
    DescapeQuotes = Replace(str, "\""", """")
End Function

Public Function CurrentCaller() As String
    If TypeOf Application.Caller Is Range Then
        Dim rng As Range
        Set rng = Application.Caller

        CurrentCaller = rng.address(External:=True)
    Else
        CurrentCaller = CStr(Application.Caller)
    End If
End Function

Public Function IsDateString(period As String)
    IsDateString = VBA.IsDate(period)
End Function

Public Function DateStringToPeriod(period As String)
    Dim d As Date: d = CDate(period)
    DateStringToPeriod = "Y" & VBA.Year(d) & ".M" & VBA.Month(d) & ".D" & VBA.Day(d)
End Function

Public Function GetAPIHeader() As String
    GetAPIHeader = "Excel - " & ExcelVersion & " - v" & AddInVersion(AddInFunctionsFile)
End Function

Public Function ExcelVersion() As String
    #If Mac Then
        #If MAC_OFFICE_VERSION = 14 Then
            ExcelVersion = "Mac2011"
        #ElseIf MAC_OFFICE_VERSION = 15 Then
            ExcelVersion = "Mac2016"
        #Else
            ExcelVersion = "Unsupported"
        #End If
    #Else
        Dim version As Integer
        version = MSOfficeVersion
        If version = 12 Then
            ExcelVersion = "Win2007"
        ElseIf version = 14 Then
            ExcelVersion = "Win2010"
        ElseIf version = 15 Then
            ExcelVersion = "Win2013"
        ElseIf version = 16 Then
            ExcelVersion = "Win2016"
        Else
            ExcelVersion = "Unsupported"
        End If
    #End If
End Function

' Returns the version of MS Office being run
'    9 = Office 2000
'   10 = Office XP / 2002
'   11 = Office 2003 & LibreOffice 3.5.2
'   12 = Office 2007
'   14 = Office 2010 or Office 2011 for Mac
'   15 = Office 2013 or Office 2016 for Mac
Public Function MSOfficeVersion() As Integer
    Dim verStr As String
    Dim startPos As Integer
    MSOfficeVersion = 0
    verStr = Application.version
    startPos = VBA.InStr(verStr, ".")
    On Error Resume Next
    If startPos > 0 Then
        MSOfficeVersion = CInt(VBA.Left(verStr, startPos - 1))
    Else
        MSOfficeVersion = CInt(verStr)
    End If
End Function
