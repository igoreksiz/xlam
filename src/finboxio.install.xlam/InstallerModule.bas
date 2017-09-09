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
    
    Dim msg As String, UpgradeVersion As String, CurrentVersion As String
    UpgradeVersion = AddInVersion(ThisWorkbook.name)
    If Not installed Is Nothing Then
        CurrentVersion = AddInVersion(installed.name)
        If CurrentVersion = "" Then
            msg = "This will install version " & UpgradeVersion & " of the finbox.io add-in."
        Else
            msg = "This will upgrade your finbox.io add-in from v" & CurrentVersion & " to v" & UpgradeVersion & "."
        End If
    Else
        msg = "This will install version " & UpgradeVersion & " of the finbox.io add-in."
    End If
    
    Dim continue As Integer
    continue = MsgBox( _
        Title:="[finbox.io] Add-in Installation", _
        Prompt:=msg & " Do you wish to continue?", _
        Buttons:=vbYesNo Or vbQuestion)

    If continue = vbYes Then
        Dim installPath As String
        installPath = SavePath(AddInInstalledFile)
        
        ' TODO: Need to fully uninstall add-on if it's in wrong place
        
        ' Uninstall the existing add-in since
        ' we need to overwrite the workbook
        If Not installed Is Nothing Then
            installed.installed = False
        End If
        
        ' Copy the workbook into the default add-in location
        ' and remove any existing functions component. The
        ' corresponding functions component will be installed
        ' automatically
        SaveCopy self, installPath
        RemoveAddInFunctions
        
        ' If there is a local version of the
        ' finboxio.functions.xlam add-in, we
        ' install that. This is primarily for
        ' convenient installation of dev
        ' (e.g. non-released) add-in versions
        If HasAddInFunctions Then
            VBA.FileCopy LocalPath(AddInFunctionsFile), SavePath(AddInFunctionsFile)
            VBA.SetAttr SavePath(AddInFunctionsFile), vbHidden
        Else
            InstallAddInFunctions
        End If
            
        ' Add the workbook as an add-in
        ' if this is a new installation
        If installed Is Nothing Then
            Dim Wb As Workbook
            
            ' AddIns.Add will fail unless a workbook is open
            ' so we create a hidden one here and clean up after
            If Application.Workbooks.count = 0 Then
                Application.ScreenUpdating = False
                Set Wb = Application.Workbooks.Add
            End If
            
            Set installed = Application.AddIns.Add(installPath, True)
            
            If Not Wb Is Nothing Then Wb.Close
        End If
        
        ' Activate the installed add-in
        installed.installed = True
        
        ' Our work is done! Close the installer
        ' workbook since the in-place add-in is
        ' now running
        Application.ScreenUpdating = True
        MsgBox _
            Title:="[finbox.io] Installation Succeeded", _
            Prompt:="The finbox.io add-in is now installed and ready to use! Enjoy!", _
            Buttons:=vbInformation
        self.Close
    ElseIf SafeDir(ThisWorkbook.path & Application.PathSeparator & ".git", vbDirectory Or vbHidden) <> "" Then
        ' If we're running this from a development directory,
        ' close the installed add-ins and continue
        If Not installed Is Nothing Then
            ' Originally wanted to use AddIn.IsOpen here, but that
            ' seems to not be available on Mac so we have to just
            ' try to close the workbook directly and ignore errors
            On Error Resume Next
            Workbooks(installed.name).Close
            UnloadAddInFunctions
            LoadAddInFunctions
        End If
    Else
        ' This add-in shouldn't be run outside
        ' of the installation directory
        self.Close
    End If
    
    GoTo Finish
    
HandleError:
    MsgBox _
        Title:="[finbox.io] Add-in Error", _
        Prompt:="Unable to install the finbox.io add-on. Please try again and contact support@finbox.io if this problem persists.", _
        Buttons:=vbCritical
    
Finish:
    Application.ScreenUpdating = True
End Function

Public Sub InstallAddInFunctions()
    cd SavePath
    
    On Error GoTo HandleError
    DownloadFile DOWNLOADS_URL & "/v" & AddInVersion & "/" & AddInFunctionsFile, StagingPath(AddInFunctionsFile)
    VBA.SetAttr StagingPath(AddInFunctionsFile), vbHidden
    PromoteStagedUpdate
    Exit Sub
HandleError:
    On Error Resume Next
    MsgBox _
        Title:="[finbox.io] Installation Failed", _
        Prompt:="The add-in functions could not be installed at this time. Please try again and contact support@finbox.io if this problem persists.", _
        Buttons:=vbCritical
    RemoveAddInFunctions
    
    cd ThisWorkbook.path
End Sub

Public Sub RemoveAddInFunctions()
    cd SavePath
    
    On Error Resume Next
    UninstallAddInFunctions
    UnloadAddInFunctions
    
    SetAttr LocalPath(AddInFunctionsFile), vbNormal
    Kill LocalPath(AddInFunctionsFile)
    
    SetAttr StagingPath(AddInFunctionsFile), vbNormal
    Kill StagingPath(AddInFunctionsFile)
    
    cd ThisWorkbook.path
End Sub

Function SavePath(Optional file As String)
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
    If file <> "" Then SavePath = SavePath & file
End Function

Sub SaveCopy(Wb, path As String)
    If ExcelVersion = "Mac2016" Then
        SaveCopyAsExcel2016 Wb, path
    Else
        Wb.SaveCopyAs path
    End If
End Sub

Sub SaveCopyAsExcel2016(Wb, path As String)
    Dim folder As String
    folder = Left(path, InStrRev(path, "/"))
    If SafeDir(folder, vbDirectory) = vbNullString Then VBA.MkDir folder
    Wb.SaveCopyAs path
End Sub
