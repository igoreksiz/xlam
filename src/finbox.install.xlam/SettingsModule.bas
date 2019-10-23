Attribute VB_Name = "SettingsModule"
Option Explicit
Option Private Module

' Sample config file:
'
' ---------
' autoUpdate=True
' updateOnLaunch=True
' allowPrereleases=True
'

Private settings As New Dictionary
Private hasReadSettings As Boolean

Public Function GetSetting(key As String, Optional default)
    If Not hasReadSettings Then ReadSettings
    GetSetting = default
    If settings.Exists(key) Then GetSetting = settings.Item(key)
End Function

Public Sub ReadSettings()
    On Error GoTo Finish
    Dim file As String, line As String, key As String, value As String
    file = LocalPath(AddInSettingsFile)
    Dim ipt As Integer
    ipt = FreeFile
    Open file For Input As ipt
        While Not EOF(ipt)
            Line Input #ipt, line
            line = VBA.Trim(Application.Clean(line))
            key = VBA.Left(line, VBA.InStr(line, "=") - 1)
            value = VBA.Mid(line, VBA.InStr(line, "=") + 1)
            If VBA.LCase(value) = "true" Then value = True
            If VBA.LCase(value) = "false" Then value = False
            If settings.Exists(key) Then settings.Remove (key)
            Call settings.Add(key, value)
        Wend
    Close #ipt
Finish:
    hasReadSettings = True
End Sub

'Public Function SetSetting(key As String, value)
'    If settings.Exists(key) Then Call settings.Remove(key)
'    Call settings.Add(key, value)
'    WriteSettings
'End Function
'
'Public Sub WriteSettings()
'    Dim file As String
'    file = LocalPath(AddInSettingsFile)
'    Open file For Output As #1
'        Dim key
'        For Each key In settings.keys
'            Print #1, key & "=" & settings.Item(key)
'        Next key
'    Close #1
'End Sub

