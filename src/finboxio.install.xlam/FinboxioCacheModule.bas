Attribute VB_Name = "FinboxioCacheModule"
' finbox.io API Integration

Option Explicit

Private CachedValues As New Dictionary
Private CachedTimestamp As New Dictionary

Public Function ClearCache()
    CachedValues.RemoveAll
    CachedTimestamp.RemoveAll
End Function

Public Function IsCached(ByVal key As String, Optional skip As Boolean = False) As Boolean

    ' Return boolean true if cached value for key is available within the cache timeout
    IsCached = False
    
    If skip Then
        Exit Function
    End If
    
    If CachedTimestamp.Exists(key) Then
        If CachedTimestamp(key) + (CACHE_TIMEOUT_MINUTES / 60 / 24) >= Now() Then IsCached = True
    End If
    
End Function

Public Sub SetCachedValue(ByVal key As String, ByVal dataValue As Variant)


    ' Set cached value and timestamp for key
    CachedValues(key) = dataValue
    CachedTimestamp(key) = Now()

End Sub

Public Function GetCachedValue(ByVal key As String) As Variant

    ' Retrieve cached value for key
    If CachedValues.Exists(key) Then
        GetCachedValue = CachedValues(key)
    Else
        GetCachedValue = CVErr(xlErrNA) ' return #NA
    End If
    
End Function

