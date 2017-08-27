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
Public Const USAGE_URL = "https://finbox.io/profile/api"
Public Const UPGRADE_URL = "https://finbox.io/premium"
Public Const UPDATE_URL = "https://finbox.io/integrations/excel?dl=1"

Public Const AUTH_URL = "https://api.finbox.io/v2/tokens"
Public Const RELEASES_URL = "https://api.github.com/repos/finboxio/xlam/releases"
Public Const INSTALLER_URL = "https://github.com/finboxio/xlam/releases/download/v"

Public Const TIER_URL = "https://api.finbox.io/beta/usage"
Public Const BATCH_URL = "https://api.finbox.io/beta/data/batch"
Public Const DOWNLOAD_URL = "https://api.finbox.io/v2/add-ons/excel"

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

Public Function AddInVersion(file As String) As String
    On Error Resume Next
    AddInVersion = Workbooks(file).Sheets("finboxio").range("AppVersion").value
End Function

Public Function AddInReleaseDate(file As String) As Date
    AddInReleaseDate = VBA.Now()
    On Error Resume Next
    AddInReleaseDate = Workbooks(file).Sheets("finboxio").range("ReleaseDate").value
End Function

Public Function AddInLocation(file As String) As String
    On Error Resume Next
    AddInLocation = Workbooks(file).FullName
End Function

