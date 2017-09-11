Attribute VB_Name = "ConfigModule"
Option Explicit
Option Private Module

Public Const RELEASES_URL = "https://api.github.com/repos/finboxio/xlam/releases"
Public Const DOWNLOADS_URL = "https://github.com/finboxio/xlam/releases/download"

Public Const AddInInstalledFile = "finboxio.xlam"
Public Const AddInInstallerFile = "finboxio.install.xlam"
Public Const AddInFunctionsFile = "finboxio.functions.xlam"
Public Const AddInKeyFile = "finboxio.key"
Public Const AddInSettingsFile = "finboxio.cfg"
Public Const AddInLogFile = "finboxio.log"

' These will be loaded on Workbook_Open
Public AddInInstalled As Boolean
Public cwd As String

Public Function AddInManagerFile() As String
    AddInManagerFile = ThisWorkbook.name
End Function

Public Function StagingFile(file As String) As String
    StagingFile = VBA.Left(file, VBA.InStrRev(file, ".")) & "staged" & VBA.Mid(file, InStrRev(file, "."))
End Function

Public Function StagingPath(file As String) As String
    StagingPath = LocalPath(StagingFile(file))
End Function

Public Sub cd(path As String)
    If VBA.Right(path, 1) = Application.PathSeparator Then
        cwd = VBA.Left(path, VBA.Len(path) - 1)
    Else
        cwd = path
    End If
End Sub

Public Function LocalPath(file As String) As String
    If cwd = "" Then cwd = ThisWorkbook.path
    LocalPath = cwd & Application.PathSeparator & file
End Function

Public Function AddInVersion(Optional file As String) As String
    If file = "" Then file = ThisWorkbook.name
    On Error Resume Next
    AddInVersion = Workbooks(file).Sheets("finboxio").Range("AppVersion").value
End Function

Public Function AddInReleaseDate(Optional file As String) As Date
    If file = "" Then file = ThisWorkbook.name
    AddInReleaseDate = VBA.Now()
    On Error Resume Next
    AddInReleaseDate = Workbooks(file).Sheets("finboxio").Range("ReleaseDate").value
End Function

Public Function AddInLocation(Optional file As String) As String
    If file = "" Then file = ThisWorkbook.name
    On Error Resume Next
    AddInLocation = Workbooks(file).FullName
End Function

Public Function SafeDir(file As String, Optional attributes As VbFileAttribute) As String
    On Error Resume Next
    SafeDir = VBA.Dir(file, attributes)
End Function

Public Function ExcelVersion() As String
    Dim version As Integer: version = MSOfficeVersion
    ExcelVersion = "Unsupported"
    
    #If Mac Then
        If version = 14 Then
            ExcelVersion = "Mac2011"
        ElseIf version = 15 Then
            ExcelVersion = "Mac2016"
        End If
    #Else
        If version = 12 Then
            ExcelVersion = "Win2007"
        ElseIf version = 14 Then
            ExcelVersion = "Win2010"
        ElseIf version = 15 Then
            ExcelVersion = "Win2013"
        ElseIf version = 16 Then
            ExcelVersion = "Win2016"
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

Sub auto_add()
End Sub
Sub auto_remove()
End Sub
