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
    Dim bAlreadyInstalled As Boolean
    Dim sAddInName As String, sAddInFileName As String, sCurrentPath As String, sStandardPath As String
    
    sCurrentPath = self.Path & Application.PathSeparator
    #If Mac Then
        sStandardPath = Application.LibraryPath
    #Else
        sStandardPath = Application.UserLibraryPath
    #End If
    
    If VBA.InStr(1, self.name, ".install.xlam", vbTextCompare) Then
        ' This is an install version, so let’s pick the proper AddIn name
        sAddInName = VBA.Left(self.name, VBA.InStr(1, self.name, ".install.xlam", vbTextCompare) - 1)
        sAddInFileName = sAddInName & ".xlam"
        
        ' Avoid the re-entry of script after activating the addin
        If Not (bAlreadyRun) Then
            bAlreadyRun = True ' Ensure we won’t install it multiple times (because Excel reopen files after an XLAM installation)
            If MsgBox("Do you want to install the finbox.io excel add-in? This will overwrite any previously installed versions.", vbYesNo) = vbYes Then
                ' Create a workbook otherwise, we get into troubles as Application.AddIns may not exist
                Set oXLApp = Application
                Set oWorkbook = oXLApp.Workbooks.Add
                ' Test if AddIn already installed
                For i = 1 To self.Application.AddIns.count
                    If self.Application.AddIns.Item(i).FullName = sStandardPath & sAddInFileName Then
                        bAlreadyInstalled = True
                        iAddIn = i
                    End If
                Next i
                If bAlreadyInstalled Then
                    ' Already installed
                    DebugBox ("Called from:'" & sCurrentPath & "' Already installed")
                    If self.Application.AddIns.Item(iAddIn).Installed Then
                        ' Deactivate the add-in to be able to overwrite the file
                        self.Application.AddIns.Item(iAddIn).Installed = False
                        self.SaveCopyAs sStandardPath & sAddInFileName
                        self.Application.AddIns.Item(iAddIn).Installed = True
                    Else
                        self.SaveCopyAs sStandardPath & sAddInFileName
                        self.Application.AddIns.Item(iAddIn).Installed = True
                    End If
                Else
                    ' Not yet installed
                    DebugBox ("Called from:'" & sCurrentPath & "' Not installed")
                    self.SaveCopyAs sStandardPath & sAddInFileName
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

Sub DebugBox(sText As String)
    If bVerboseMessages Then MsgBox (sText)
End Sub
