Attribute VB_Name = "FinboxioConfigModule"
' finbox.io API Integration

Option Explicit

Public Const AppVersion = "v0.15"
Public Const AppTitle = "finbox.io Add-In " & AppVersion

Public Const CACHE_TIMEOUT_MINUTES = 60
Public Const MAX_BATCH_SIZE = 99

Public Const PROFILE_URL = "https://finbox.io/profile"
Public Const WATCHLIST_URL = "https://finbox.io/watchlist"
Public Const SCREENER_URL = "https://finbox.io/screener"
Public Const TEMPLATES_URL = "https://finbox.io/templates"
Public Const SIGNUP_URL = "https://finbox.io/signup"
Public Const HELP_URL = "https://finbox.io/how-to/getting-started/using-excel-add-on"
Public Const USAGE_URL = "https://finbox.io/profile/api"
Public Const UPGRADE_URL = "https://finbox.io/premium"
Public Const UPDATE_URL = "https://finbox.io/integrations/excel?dl=1"

Public Const AUTH_URL = "https://api.finbox.io/v2/tokens"
Public Const UPDATES_URL = "https://api.finbox.io/v2/add-ons/excel/latest"

Public Const TIER_URL = "https://api.finbox.io/beta/usage"
Public Const BATCH_URL = "https://api.finbox.io/beta/data/batch"

#If Mac Then
    #If MAC_OFFICE_VERSION < 15 Then
        Public Const EXCEL_VERSION = "Mac2011"
    #Else
        Public Const EXCEL_VERSION = "Mac2016"
    #End If
#Else
    Public Const EXCEL_VERSION = "Win"
#End If
