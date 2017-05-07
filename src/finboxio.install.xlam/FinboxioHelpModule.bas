Attribute VB_Name = "FinboxioHelpModule"
Option Explicit

Public Function LoadHelp()
    ActiveWorkbook.FollowHyperlink HELP_URL
End Function

