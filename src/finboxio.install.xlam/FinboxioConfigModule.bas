Attribute VB_Name = "FinboxioConfigModule"
' finbox.io API Integration

Option Explicit

Public Const AppVersion = "v0.6"
Public Const AppTitle = "finbox.io Add-In " & AppVersion

Public Const CACHE_TIMEOUT_MINUTES = 60
Public Const MAX_BATCH_SIZE = 99

Public Const SIGNUP_URL = "https://finbox.io/signup"
Public Const HELP_URL = "https://finbox.io/how-to/getting-started/using-excel-add-on"

Public Const AUTH_URL = "https://api.finbox.io/v2/tokens"
Public Const UPDATES_URL = "https://api.staging.finbox.io/v2/add-ons/excel/latest"

Public Const BATCH_URL = "https://api.finbox.io/beta/data/batch"

#If Mac Then
    #If MAC_OFFICE_VERSION < 15 Then
        Public Const EXCEL_VERSION = "Mac2011"
    #End If
#End If
