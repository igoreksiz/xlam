Attribute VB_Name = "FinboxioUnlinkModule"
Option Explicit

Public Sub UnlinkFormulas()
    On Error GoTo ShowWarning
    
    Dim wbName As String
    wbName = ActiveWorkbook.name
    wbName = Replace(wbName, ".xlsm", "")
    wbName = Replace(wbName, ".xlsx", "")
    wbName = Replace(wbName, ".xls", "")

    Dim msg As String, choice As Variant
    msg = "This will save any changes you have made to the current workbook and create a copy with all finbox.io formulas replaced by their current values. Do you wish to continue?"
    choice = MsgBox(msg, vbYesNo, "Save and unlink?")
    Select Case choice
        Case vbYes
            Dim fileSaveName As Variant
            #If Mac Then
                fileSaveName = Application.GetSaveAsFilename( _
                    InitialFileName:=wbName & " - unlinked")
            #Else
                fileSaveName = Application.GetSaveAsFilename( _
                    InitialFileName:=wbName & " - unlinked", _
                    fileFilter:="Excel Workbook (*.xlsx), *.xlsx")
            #End If
    
            If TypeName(fileSaveName) <> "Boolean" Then
                Application.DisplayAlerts = False
                ActiveWorkbook.Save
                
                Dim calcType: calcType = Application.Calculation
                Application.Calculation = xlCalculationManual
                Dim r As range, i As Long
                For i = 1 To Sheets.count
                    On Error Resume Next
                    For Each r In Sheets(i).UsedRange.SpecialCells(xlCellTypeFormulas)
                        If r.formula Like "*FNBX*" Then r.value = r.value
                    Next r
                Next i
                Application.Calculation = calcType
                
                ActiveWorkbook.SaveAs FileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
                Application.DisplayAlerts = True
            End If
    End Select
    Exit Sub
    
ShowWarning:
    MsgBox ("This workbook cannot be unlinked")
End Sub
