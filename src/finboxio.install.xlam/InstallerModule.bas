Attribute VB_Name = "InstallerModule"
Option Explicit
Option Private Module

Const DebugMessages = False

Dim AlreadyRun As Boolean

Public Function InstallAddIn(self As Workbook)
    ' Don't run if add-in is already installed
    ' Don't run if called more than once (excel reopens file after installation)
    InstallAddIn = (self.name = XlamFile(AddInLoaderFile))
    If InstallAddIn Or AlreadyRun Then GoTo Finish
    
    Dim i As addIn, installed As addIn
    For Each i In self.Application.AddIns
        If i.name = XlamFile(AddInLoaderFile) Then
            installed = i
            Exit For
        End If
    Next i
    
    Dim prompt As String
    If installed Then
        prompt = "This will upgrade your finbox.io add-in to version " & AddInVersion & "."
    Else
        prompt = "This will install version " & AddInVersion & " of the finbox.io add-in."
    End If
    
    Dim continue As Integer
    continue = MsgBox( _
        Title:="finbox.io", _
        prompt:=prompt & " Do you wish to continue?", _
        Buttons:=vbYesNo & vbQuestion)

    If continue = vbYes Then
        Dim installPath As String
        ' TODO: Need to fully uninstall add-on if it's in wrong place
        If installed Then
            installPath = StagedXlamPath(AddInLoaderFile)
            SaveCopy self, installPath
        Else
            installPath = XlamPath(AddInLoaderFile)
            SaveCopy self, installPath
            Set installed = Application.AddIns.Add(installPath, True)
            installed.installed = True
        End If
    End If
Finish:
    AlreadyRun = True
End Function

Function SavePath()
    #If Mac Then
        If EXCEL_VERSION = "Mac2016" Then
            SavePath = MacScript("return POSIX path of (path to desktop folder) as string")
            SavePath = Replace(SavePath, "/Desktop", "") & "Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/AppData/Microsoft/Office/16.0/"
            SavePath = SavePath & "Add-Ins/"
        Else
            SavePath = Application.LibraryPath
        End If
    #Else
        SavePath = Application.UserLibraryPath
    #End If
End Function

Sub SaveCopy(wb, name As String)
    If EXCEL_VERSION = "Mac2016" Then
        SaveCopyAsExcel2016 wb, name
    Else
        Dim path As String
        path = SavePath()
        wb.SaveCopyAs path & name
    End If
End Sub

Sub DebugBox(sText As String)
    If DebugMessages Then MsgBox (sText)
End Sub

Function CreateFolderinMacOffice2016() As String
    'Function to create folder if it not exists in the Microsoft Office Folder
    Dim PathToFolder As String
    Dim TestStr As String

    PathToFolder = SavePath()

    On Error Resume Next
    TestStr = Dir(PathToFolder, vbDirectory)
    On Error GoTo 0
    If TestStr = vbNullString Then
        MkDir PathToFolder
        'You can use this msgbox line for testing if you want
        'MsgBox "You find the new folder in this location :" & PathToFolder
    End If
    CreateFolderinMacOffice2016 = PathToFolder
End Function

Sub SaveCopyAsExcel2016(wb, name As String)
    'Save a copy of the file with a Date/time stamp in a sub folder
    'in the Microsoft Office Folder
    'This macro use the custom function named : CreateFolderinMacOffice2016
    Dim Folderstring As String
    Dim StrFilePath As String
    Dim StrFileName As String
    Dim FileExtStr As String

    If VBA.InStr(name, Application.PathSeparator) Then
        wb.SaveCopyAs name
    Else
        'Create folder if it not exists in the Microsoft Office Folder
        Folderstring = CreateFolderinMacOffice2016()
    
        StrFilePath = Folderstring
        StrFileName = name
    
        wb.SaveCopyAs StrFilePath & StrFileName
    End If
End Sub



