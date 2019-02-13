Attribute VB_Name = "FinboxioConfigModule"
Option Explicit
Option Private Module

Public Const CACHE_TIMEOUT_MINUTES = 60
Public Const MAX_BATCH_SIZE = 99

Public Const PROFILE_URL = "https://finbox.io/profile"
Public Const WATCHLIST_URL = "https://finbox.io/watchlist"
Public Const SCREENER_URL = "https://finbox.io/screener"
Public Const TEMPLATES_URL = "https://finbox.io/templates"
Public Const SIGNUP_URL = "https://finbox.io/signup"
Public Const HELP_URL = "https://finbox.io/blog/using-the-excel-add-in/"
Public Const UPGRADE_URL = "https://finbox.io/premium"

Public Const AUTH_URL = "https://api.finbox.io/v2/tokens"
Public Const API_URL_BETA = "https://api.finbox.io/beta"
Public Const API_URL_V3 = "https://public-api.finbox.io/v3"

Public Const LIMIT_EXCEEDED_ERROR = 20400
Public Const INVALID_AUTH_ERROR = 20401
Public Const INVALID_ARGS_ERROR = 20402
Public Const INVALID_KEY_ERROR = 20403
Public Const INVALID_PERIOD_ERROR = 20404
Public Const UNSUPPORTED_COMPANY_ERROR = 20405
Public Const UNSUPPORTED_METRIC_ERROR = 20406
Public Const RESTRICTED_COMPANY_ERROR = 20407
Public Const RESTRICTED_METRIC_ERROR = 20408
Public Const MISSING_VALUE_ERROR = 20409
Public Const UNSPECIFIED_API_ERROR = 20500

Public Const AddInInstalledFile = "finboxio.xlam"
Public Const AddInInstallerFile = "finboxio.install.xlam"
Public Const AddInFunctionsFile = "finboxio.functions.xlam"
Public Const AddInKeyFile = "finboxio.key"
Public Const AddInLogFile = "finboxio.log"
Public Const AddInSettingsFile = "finboxio.cfg"

Public Function AddInManagerFile() As String
    On Error Resume Next
    AddInManagerFile = Workbooks(AddInInstalledFile).name
    AddInManagerFile = Workbooks(AddInInstallerFile).name
End Function

Public Function StagingPath(file As String) As String
    StagingPath = LocalPath(VBA.Left(file, VBA.InStrRev(file, ".")) & "staged" & VBA.Mid(file, InStrRev(file, ".")))
End Function

Public Function LocalPath(file As String) As String
    LocalPath = ThisWorkbook.path & Application.PathSeparator & file
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

Public Function ApiUrl()
    ApiUrl = GetSetting("finboxioApiUrl", API_URL_BETA)
End Function

Public Function TierUrl()
    TierUrl = ApiUrl & "/usage"
End Function

Public Function BatchUrl()
    BatchUrl = ApiUrl & "/data/batch"
End Function

Sub auto_add()
End Sub
Sub auto_remove()
End Sub
