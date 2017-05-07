Attribute VB_Name = "FinboxioHelpModule"
Option Explicit

Public Function FinboxioHelp(Optional control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink HELP_URL
End Function

