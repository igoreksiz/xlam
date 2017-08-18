Attribute VB_Name = "FinboxioConfigModule"
Option Explicit
Option Private Module

Public Const AppVersion = "v0.16"
Public Const AppTitle = "finbox.io Add-In " & AppVersion

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
Public Const UPDATES_URL = "https://api.finbox.io/v2/add-ons/excel/latest"

Public Const TIER_URL = "https://api.finbox.io/beta/usage"
Public Const BATCH_URL = "https://api.finbox.io/beta/data/batch"

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

#If Mac Then
    #If MAC_OFFICE_VERSION < 15 Then
        Public Const EXCEL_VERSION = "Mac2011"
    #Else
        Public Const EXCEL_VERSION = "Mac2016"
    #End If
#Else
    Public Const EXCEL_VERSION = "Win"
#End If

