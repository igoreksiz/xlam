Attribute VB_Name = "FinboxioUpdateModule"
Option Explicit
Option Private Module

Public Const AddInLoaderFile = "finboxio"

Public Sub CheckUpdates(Optional explicit As Boolean = False, Optional wb As Workbook)
    If Dir(StagedXlamPath(AddInLoaderFile)) <> "" Then
        Dim i As Integer
        For i = 1 To Application.AddIns.count
            On Error Resume Next
            If Application.AddIns.Item(i).FullName = XlamPath(AddInLoaderFile) Then Exit For
        Next i
        If i <= Application.AddIns.count Then Application.AddIns.Item(i).Installed = False
        PromoteStagedFile AddInLoaderFile
        If i <= Application.AddIns.count Then Application.AddIns.Item(i).Installed = True
    End If
End Sub

Public Function StagedXlamPath(file As String) As String
    StagedXlamPath = XlamPath(file & ".staged")
End Function

Public Function XlamPath(file As String) As String
    XlamPath = ThisWorkbook.path & Application.PathSeparator & file & ".xlam"
End Function

Public Sub PromoteStagedFile(file As String)
    SetAttr XlamPath(file), vbNormal
    Kill XlamPath(file)
    Name StagedXlamPath(file) As XlamPath(file)
End Sub


