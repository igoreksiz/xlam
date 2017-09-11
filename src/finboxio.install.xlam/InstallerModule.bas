Attribute VB_Name = "InstallerModule"
Option Explicit
Option Private Module

Public installing As Boolean
Public uninstalling As Boolean

Public Function IsInstalling() As Boolean
    IsInstalling = installing Or uninstalling
End Function

' Install this workbook as an add-in in the default
' Excel add-in location. This simplifies deployment
' and management across different platforms and
' ensures valid trust settings for the add-in.
'
' This function returns `True` if it is called from
' an already-installed instance of the add-in,
' otherwise it will return `False`.

Public Function InstallAddIn(self As Workbook) As Boolean
    ' Don't run if add-in is already installed
    InstallAddIn = (self.name = AddInInstalledFile)
    #If Mac Then
        If Not InstallAddIn Then MacInstallPrompt.Show
    #Else
        If Not InstallAddIn Then InstallPrompt.Show
    #End If
End Function

Public Sub FinishInstalling()
    On Error GoTo HandleError
    
    Dim i As addin, installed As addin
    For Each i In Application.AddIns
        If i.name = AddInInstalledFile Then
            Set installed = i
            Exit For
        End If
    Next i

    installing = True
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
    SaveCopy ThisWorkbook, installPath
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
    installing = False
    MsgBox _
        Title:="[finbox.io] Installation Succeeded", _
        Prompt:="The finbox.io add-in is now installed and ready to use! Enjoy!", _
        Buttons:=vbInformation
    
    On Error Resume Next
    ThisWorkbook.Close
    Exit Sub
    
HandleError:
    installing = False
    Application.ScreenUpdating = True
    MsgBox _
        Title:="[finbox.io] Add-in Error", _
        Prompt:="Unable to install the finbox.io add-on. Please try again and contact support@finbox.io if this problem persists.", _
        Buttons:=vbCritical
End Sub

Public Sub CancelInstall()
    Dim i As addin, installed As addin
    For Each i In Application.AddIns
        If i.name = AddInInstalledFile Then
            Set installed = i
            Exit For
        End If
    Next i
    
    If SafeDir(ThisWorkbook.path & Application.PathSeparator & ".git", vbDirectory Or vbHidden) <> "" Then
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
        ThisWorkbook.Close
    End If
End Sub

Public Sub UninstallAddIn()
    uninstalling = True
    
    On Error Resume Next
        
    ' Uninstall and delete installed add-in files
    Dim i As addin, installed As addin
    For Each i In Application.AddIns
        If VBA.InStr(i.name, "finbox") > 0 Then
            Workbooks(i.name).Close
            i.installed = False
            If SafeDir(i.FullName) <> "" Then Kill i.FullName
            If SafeDir(i.FullName, vbHidden) <> "" Then
                SetAttr i.FullName, vbNormal
                Kill i.FullName
            End If
        End If
    Next i
    
    cd SavePath
    
    ' Second check to make sure the add-in manager is removed
    Workbooks(AddInInstalledFile).Close
    If SafeDir(LocalPath(AddInInstalledFile)) <> "" Then Kill LocalPath(AddInInstalledFile)
    If SafeDir(LocalPath(AddInInstalledFile), vbHidden) <> "" Then
        SetAttr LocalPath(AddInInstalledFile), vbNormal
        Kill LocalPath(AddInInstalledFile)
    End If
    
    If SafeDir(StagingPath(AddInInstalledFile)) <> "" Then Kill StagingPath(AddInInstalledFile)
    If SafeDir(StagingPath(AddInInstalledFile), vbHidden) <> "" Then
        SetAttr StagingPath(AddInInstalledFile), vbNormal
        Kill StagingPath(AddInInstalledFile)
    End If
    
    ' Second check to make sure the add-in functions are removed
    Workbooks(AddInFunctionsFile).Close
    If SafeDir(LocalPath(AddInFunctionsFile)) <> "" Then Kill LocalPath(AddInFunctionsFile)
    If SafeDir(LocalPath(AddInFunctionsFile), vbHidden) <> "" Then
        SetAttr LocalPath(AddInFunctionsFile), vbNormal
        Kill LocalPath(AddInFunctionsFile)
    End If
    
    If SafeDir(StagingPath(AddInFunctionsFile)) <> "" Then Kill StagingPath(AddInFunctionsFile)
    If SafeDir(StagingPath(AddInFunctionsFile), vbHidden) <> "" Then
        SetAttr StagingPath(AddInFunctionsFile), vbNormal
        Kill StagingPath(AddInFunctionsFile)
    End If
    
    ' Delete the api key file
    If SafeDir(LocalPath(AddInKeyFile)) <> "" Then Kill LocalPath(AddInKeyFile)
    If SafeDir(LocalPath(AddInKeyFile), vbHidden) <> "" Then
        SetAttr LocalPath(AddInKeyFile), vbNormal
        Kill LocalPath(AddInKeyFile)
    End If
    
    ' Delete the config file
    If SafeDir(LocalPath(AddInSettingsFile)) <> "" Then Kill LocalPath(AddInSettingsFile)
    If SafeDir(LocalPath(AddInSettingsFile), vbHidden) <> "" Then
        SetAttr LocalPath(AddInSettingsFile), vbNormal
        Kill LocalPath(AddInSettingsFile)
    End If
    
    cd ThisWorkbook.path
    
    uninstalling = False
    
    MsgBox _
        Title:="[finbox.io] Add-In Removed", _
        Prompt:="The finbox.io add-in has been successfully removed. Hope to see you back soon!", _
        Buttons:=vbInformation
    
    ThisWorkbook.Close
End Sub

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

Public Sub CloseInstaller()
    If ThisWorkbook.name = AddInInstallerFile Then Exit Sub
    On Error GoTo Closed
    Dim opened As String
    opened = Workbooks(AddInInstallerFile).name
    If Not Application.Run(AddInInstallerFile & "!IsInstalling") Then
        Workbooks(AddInInstallerFile).Close
    End If
Closed:
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
    If SafeDir(path) <> "" Then Kill path
    If SafeDir(path, vbHidden) <> "" Then
        SetAttr path, vbNormal
        Kill path
    End If
    
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
