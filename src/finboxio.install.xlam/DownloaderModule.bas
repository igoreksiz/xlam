Attribute VB_Name = "DownloaderModule"
Option Explicit
Option Private Module

Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Sub NativeDownload(url As String, file As String)
    Dim result As Long
    result = URLDownloadToFile(0, url, file, 0, 0)
End Sub

Public Sub DownloadFile(url As String, file As String)
    NativeDownload url, file
End Sub
