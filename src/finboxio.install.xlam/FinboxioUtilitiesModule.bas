Attribute VB_Name = "FinboxioUtilitiesModule"
' finbox.io API Integration

' Written by Michael Chambers, April 2017
' michael@mrchambers.f9.co.uk

' Upwork Contract Id 17916950

Option Explicit

Public Function CollectionToString(ByVal dataCol As Variant) As String
    Dim i As Integer
    For i = 1 To dataCol.Count
        If CollectionToString <> "" Then CollectionToString = CollectionToString & ", "
        CollectionToString = CollectionToString & dataCol(i)
    Next i
End Function

Public Function MSOffVer() As Integer
' Function returns version of MS Office being run
'    9 = Office 2000
'   10 = Office XP / 2002
'   11 = Office 2003 & LibreOffice 3.5.2
'   12 = Office 2007
'   14 = Office 2010 or Office 2011 for Mac
'   15 = Office 2013 or Office 2016 for Mac
    
    Dim verStr As String
    Dim startPos As Integer
    MSOffVer = 0
        
    verStr = Application.Version
    startPos = InStr(verStr, ".")
        
        On Error Resume Next
    If startPos > 0 Then
        MSOffVer = CInt(VBA.Left(verStr, startPos - 1))
    Else
        MSOffVer = CInt(verStr)
    End If
        On Error GoTo 0

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

Public Function ValidAddress(strAddress As String) As Boolean
    Dim r As Range
    On Error Resume Next
    Set r = Range(strAddress)
    If Not r Is Nothing Then ValidAddress = True
End Function

Public Function EscapeQuotes(str As String) As String
    EscapeQuotes = Replace(str, """", "\""")
End Function

Public Function DescapeQuotes(str As String) As String
    EscapeQuotes = Replace(str, "\""", """")
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

Public Function GetAPIHeader()
    Dim APIHeader As String
    
    #If Mac Then
        APIHeader = "Excel_Mac_"
    #Else
        APIHeader = "Excel_Win_"
    #End If
    
    APIHeader = APIHeader & MSOffVer() & "-" & AppVersion
    
    GetAPIHeader = APIHeader
End Function

Public Function FixAddinLinks()
    Dim Sht As Worksheet
    Application.ScreenUpdating = False
    For Each Sht In Worksheets
        Sht.Cells.Replace _
            What:="'*finboxio.install.xlam'!", _
            Replacement:="", _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            MatchCase:=False
    Next Sht
    For Each Sht In Worksheets
        Sht.Cells.Replace _
            What:="'*finboxio.xlam'!", _
            Replacement:="", _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            MatchCase:=False
    Next Sht
    Application.Run "ResetFindReplace"
    Application.ScreenUpdating = True
End Function
