Attribute VB_Name = "FinboxioUnlinkModule"
Option Explicit

' No need to reinvent the wheel...
'
' https://github.com/intrinio/intrinio-excel/blob/0668ac3d31b9e79832eaf1483c0600a571827ee5/src/IntrinioUtilities.bas#L51-L101
'
' The MIT License (MIT)
' =====================
'
' Copyright (c) `2014-2016` `Tribunat LLC, dba Intrinio`
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of
' this software and associated documentation files (the "Software"), to deal in
' the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
' of the Software, and to permit persons to whom the Software is furnished to do
' so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'

Public Sub UnlinkFormulas()
    Dim ws As Worksheet
    Dim Ans As Variant
    Dim fileSaveName As Variant
    Dim wbName As String
    Dim Msg As String
    Dim FileName As String
    Dim i As Integer
    Dim r As range
    Dim calcType
    
    calcType = Application.Calculation
    
    On Error GoTo RestoreSettings
    
    Application.Calculation = xlCalculationManual
    
    wbName = ActiveWorkbook.name
    wbName = Replace(wbName, ".xlsm", "")
    wbName = Replace(wbName, ".xlsx", "")
    wbName = Replace(wbName, ".xls", "")

    Msg = "This will save the current workbook and create a copy with all finbox.io formulas replaced by their current values. Do you wish to continue?"

    Ans = MsgBox(Msg, vbYesNo, "Save and unlink?")
     
    Select Case Ans
              
    Case vbYes
        #If Mac Then
            fileSaveName = Application.GetSaveAsFilename( _
                InitialFileName:=wbName & " - UNLINKED")
        #Else
            fileSaveName = Application.GetSaveAsFilename( _
                InitialFileName:=wbName & " - UNLINKED", _
                fileFilter:="Excel Workbook (*.xlsx), *.xlsx")
        #End If

        If TypeName(fileSaveName) <> "Boolean" Then
            Application.DisplayAlerts = False
            ActiveWorkbook.Save
            
            For i = 1 To Sheets.count
                On Error Resume Next
                For Each r In Sheets(i).UsedRange.SpecialCells(xlCellTypeFormulas)
                    If r.formula Like "*FNBX*" Then r.value = r.value
                Next r
            Next i

            ActiveWorkbook.SaveAs FileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
            Application.DisplayAlerts = True
        End If
    End Select
    
RestoreSettings:
    Application.DisplayAlerts = True
    Application.Calculation = calcType
End Sub
