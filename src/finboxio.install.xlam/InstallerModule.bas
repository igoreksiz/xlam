Attribute VB_Name = "InstallerModule"
Option Explicit
Option Private Module

' Install this workbook as an add-in in the default
' Excel add-in location. This simplifies deployment
' and management across different platforms and
' ensures valid trust settings for the add-in.
'
' This function returns `True` if it is called from
' an already-installed instance of the add-in,
' otherwise it will return `False`.

Public Function InstallAddIn(self As Workbook) As Boolean
    On Error GoTo HandleError
    
    ' Don't run if add-in is already installed
    InstallAddIn = (self.name = AddInInstalledFile)
    If InstallAddIn Then GoTo Finish
    
    Dim i As addIn, installed As addIn
    For Each i In Application.AddIns
        If i.name = AddInInstalledFile Then
            Set installed = i
            Exit For
        End If
    Next i
    
    Dim prompt As String, UpgradeVersion As String, CurrentVersion As String
    UpgradeVersion = AddInVersion(ThisWorkbook.name)
    If Not installed Is Nothing Then
        CurrentVersion = AddInVersion(installed.name)
        prompt = "This will upgrade your finbox.io add-in from v" & CurrentVersion & " to v" & UpgradeVersion & "."
    Else
        prompt = "This will install version " & UpgradeVersion & " of the finbox.io add-in."
    End If
    
    Dim continue As Integer
    continue = MsgBox( _
        Title:="finbox.io", _
        prompt:=prompt & " Do you wish to continue?", _
        Buttons:=vbYesNo Or vbQuestion)

    If continue = vbYes Then
        Dim installPath As String
        installPath = SavePath & AddInInstalledFile
        
        ' TODO: Need to fully uninstall add-on if it's in wrong place
        
        ' Uninstall the existing add-in since
        ' we need to overwrite the workbook
        If Not installed Is Nothing Then
            installed.installed = False
        End If
        
        ' Copy the workbook into the default
        ' add-in location
        SaveCopy self, installPath
        
        ' If there is a local version of the
        ' finboxio.functions.xlam add-in, we
        ' stage that. This is primarily for
        ' convenient installation of dev
        ' (e.g. non-released) add-in versions
        If HasAddInFunctions Then
            VBA.FileCopy _
                LocalPath(AddInFunctionsFile), _
                SavePath & StagingFile(AddInFunctionsFile)
            VBA.SetAttr SavePath & StagingFile(AddInFunctionsFile), vbHidden
        End If
            
        ' Add the workbook as an add-in
        ' if this is a new installation
        If installed Is Nothing Then
            Dim wb As Workbook
            
            ' AddIns.Add will fail unless a workbook is open
            ' so we create a hidden one here and clean up after
            If Application.Workbooks.count = 0 Then
                Application.ScreenUpdating = False
                Set wb = Application.Workbooks.Add
            End If
            
            Set installed = Application.AddIns.Add(installPath, True)
            
            If Not wb Is Nothing Then wb.Close
        End If
        
        ' Activate the installed add-in
        installed.installed = True
        
        ' Our work is done! Close the installer
        ' workbook since the in-place add-in is
        ' now running
        Application.ScreenUpdating = True
        self.Close
    ElseIf VBA.Dir(ThisWorkbook.path & Application.PathSeparator & ".git", vbDirectory Or vbHidden) <> "" Then
        ' If we're running this from a development directory,
        ' close the installed add-ins and continue
        If Not installed Is Nothing Then
            UnloadAddInFunctions
            Workbooks(installed.name).Close
        End If
    Else
        ' This add-in shouldn't be run outside
        ' of the installation directory
        self.Close
    End If
    
    GoTo Finish
    
HandleError:
    MsgBox "Got Error: " & Err.Description
    
Finish:
    Application.ScreenUpdating = True
End Function

Function SavePath()
    #If Mac Then
        If ExcelVersion = "Mac2016" Then
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

Sub SaveCopy(wb, path As String)
    If ExcelVersion = "Mac2016" Then
        SaveCopyAsExcel2016 wb, path
    Else
        wb.SaveCopyAs path
    End If
End Sub

Sub SaveCopyAsExcel2016(wb, path As String)
    Dim folder As String
    folder = Left(path, InStrRev(path, "/"))
    If VBA.Dir(folder, vbDirectory) = vbNullString Then VBA.MkDir folder
    wb.SaveCopyAs path
End Sub
