Attribute VB_Name = "FinboxioInstallModule"
' (c) Willy Roche (willy.roche(at)centraliens.net)
' Install procedure of XLAM (library of functions)
' This procedure will install a file name .install.xlam in the proper excel directory
' The install package will be name
' During install you may be prompt to enable macros (accept it)
' You can accept to install or refuse (which let you modify the XLAM file macros or install procedure

Option Explicit

' Taken from https://stackoverflow.com/questions/9745469/automatically-install-excel-vba-add-in

Const bVerboseMessages = False ' Set it to True to be able to Debug install mechanism
Dim bAlreadyRun As Boolean ' Will be use to verify if the procedure has already been run

Public Sub InstallAddin(self)
    ' This sub will automatically start when xlam file is opened (both install version and installed version)
    Dim oAddIn As Object, oXLApp As Object, oWorkbook As Workbook
    Dim i As Integer
    Dim iAddIn As Integer
    Dim bAlreadyInstalled As Boolean, bLegacyInstalled As Boolean
    Dim sAddInName As String, sAddInFileName As String, sCurrentPath As String, sStandardPath As String
    
    sCurrentPath = self.path & Application.PathSeparator
    sStandardPath = SavePath()
    
    If Not self.name = "finboxio.xlam" Then
        sAddInFileName = "finboxio.xlam"
        
        ' Avoid the re-entry of script after activating the addin
        If Not (bAlreadyRun) Then
            bAlreadyRun = True ' Ensure we wont install it multiple times (because Excel reopen files after an XLAM installation)
            If MsgBox("Do you want to install the finbox.io excel add-in? This will overwrite any previously installed versions.", vbYesNo) = vbYes Then
                ' Create a workbook otherwise, we get into troubles as Application.AddIns may not exist
                Set oXLApp = Application
                Set oWorkbook = oXLApp.Workbooks.Add
                ' Test if AddIn already installed
                For i = 1 To self.Application.AddIns.count
                    On Error Resume Next
                    If self.Application.AddIns.Item(i).FullName = sStandardPath & sAddInFileName Then
                        bAlreadyInstalled = True
                        iAddIn = i
                    ElseIf self.Application.AddIns.Item(i).name = "finboxio.xlam" Then
                        bLegacyInstalled = True
                        iAddIn = i
                    End If
                Next i
                If bAlreadyInstalled Then
                    ' Already installed
                    DebugBox ("Called from:'" & sCurrentPath & "' Already installed")
                    If self.Application.AddIns.Item(iAddIn).Installed Then
                        ' Deactivate the add-in to be able to overwrite the file
                        self.Application.AddIns.Item(iAddIn).Installed = False
                        SaveCopy self, sAddInFileName
                        self.Application.AddIns.Item(iAddIn).Installed = True
                    Else
                        SaveCopy self, sAddInFileName
                        self.Application.AddIns.Item(iAddIn).Installed = True
                    End If
                ElseIf bLegacyInstalled Then
                    ' Installed in an old location
                    DebugBox ("Called from:'" & sCurrentPath & "' Legacy installed")
                    self.Application.AddIns.Item(iAddIn).Installed = False
                    SaveCopy self, self.Application.AddIns.Item(iAddIn).FullName
                    self.Application.AddIns.Item(iAddIn).Installed = True
                    DebugBox ("Legacy overwritten")
                Else
                    ' Not yet installed
                    DebugBox ("Called from:'" & sCurrentPath & "' Not installed")
                    SaveCopy self, sAddInFileName
                    Set oAddIn = oXLApp.AddIns.Add(sStandardPath & sAddInFileName, True)
                    oAddIn.Installed = True
                End If
                Dim restart As Integer
                restart = MsgBox("The finbox.io add-in was successfully installed! You must quit and restart Excel to activate it. Would you like to quit Excel now?", vbYesNo)
                If restart = vbYes Then
                    oWorkbook.Close (False) ' Close the workbook opened by the install script
                    oXLApp.Quit ' Close the app opened by the install script
                    Set oWorkbook = Nothing ' Free memory
                    Set oXLApp = Nothing ' Free memory
                    self.Close (False)
                End If
            End If
        Else
            DebugBox ("Called from:'" & sCurrentPath & "' Already Run")
            ' Already run, so nothing to do
        End If
    Else
        DebugBox ("Called from:'" & sCurrentPath & "' in place")
        ' Already in right place, so nothing to do
    End If
End Sub

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
    If bVerboseMessages Then MsgBox (sText)
End Sub

Function CreateFolderinMacOffice2016() As String
    'Function to create folder if it not exists in the Microsoft Office Folder
    'Ron de Bruin : 8-Jan-2016
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
        With wb
            .SaveCopyAs name
        End With
    Else
        'Create folder if it not exists in the Microsoft Office Folder
        Folderstring = CreateFolderinMacOffice2016()
    
        StrFilePath = Folderstring
        StrFileName = name
    
        With wb
            .SaveCopyAs StrFilePath & StrFileName
        End With
    End If
End Sub


