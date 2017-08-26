Attribute VB_Name = "ConfigModule"
Option Explicit
Option Private Module

' These will be loaded on Workbook_Open
Public AddInVersion As String
Public AddInDate As Date
Public AddInDetail As String
Public AddInInstalled As Boolean

Public Const AddInLoaderFile = "finboxio"
Public Const AddInFunctionsFile = "finboxio.functions"

Public Const RELEASES_URL = "https://api.github.com/repos/finboxio/xlam/releases"

#If Mac Then
    #If MAC_OFFICE_VERSION < 15 Then
        Public Const EXCEL_VERSION = "Mac2011"
    #Else
        Public Const EXCEL_VERSION = "Mac2016"
    #End If
#Else
    Public Const EXCEL_VERSION = "Win"
#End If
