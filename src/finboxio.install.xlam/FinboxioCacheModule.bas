Attribute VB_Name = "FinboxioCacheModule"
' finbox.io API Integration

Option Explicit

Private CachedValues As New Dictionary
Private CachedTimestamp As New Dictionary
Private RecachedValues As New Dictionary
Private RecachedTimestamp As New Dictionary
Private Recaching As Boolean

Public Function ClearCache()
    CachedValues.RemoveAll
    CachedTimestamp.RemoveAll
    RecachedValues.RemoveAll
    RecachedTimestamp.RemoveAll
    Recaching = False
End Function

Public Function StartRecache()
    Recaching = True
End Function

Public Function StopRecache()
    Recaching = False
    Dim key As Variant
    For Each key In RecachedValues.keys
        If TypeName(RecachedValues(key)) = "Collection" Then
            Set CachedValues(key) = RecachedValues(key)
        Else
            CachedValues(key) = RecachedValues(key)
        End If
        CachedTimestamp(key) = RecachedTimestamp(key)
    Next
    RecachedValues.RemoveAll
    RecachedTimestamp.RemoveAll
End Function

Public Function IsCached(ByVal key As String, Optional skip As Boolean = False) As Boolean
    ' Return boolean true if cached value for key is available within the cache timeout
    IsCached = False
    
    If skip Then
        Exit Function
    End If
    
    If Not Recaching Then
        If CachedTimestamp.Exists(key) Then
            If CachedTimestamp(key) + (CACHE_TIMEOUT_MINUTES / 60 / 24) >= Now() Then IsCached = True
        End If
    Else
        If RecachedTimestamp.Exists(key) Then
            If RecachedTimestamp(key) + (CACHE_TIMEOUT_MINUTES / 60 / 24) >= Now() Then IsCached = True
        End If
    End If
End Function

Public Sub SetCachedValue(ByVal key As String, ByVal dataValue As Variant)
    ' Set cached value and timestamp for key
    If Not Recaching Then
        If TypeName(dataValue) = "Collection" Then
            Set CachedValues(key) = dataValue
        Else
            CachedValues(key) = dataValue
        End If
        CachedTimestamp(key) = Now()
    Else
        If TypeName(dataValue) = "Collection" Then
            Set RecachedValues(key) = dataValue
        Else
            RecachedValues(key) = dataValue
        End If
        RecachedTimestamp(key) = Now()
    End If
End Sub

Public Function GetCachedValue(ByVal key As String) As Variant
    ' Retrieve cached value for key
    If Not Recaching Then
        If CachedValues.Exists(key) Then
            If TypeName(CachedValues(key)) = "Collection" Then
                Set GetCachedValue = CachedValues(key)
            Else
                GetCachedValue = CachedValues(key)
            End If
        Else
            GetCachedValue = CVErr(xlErrNA) ' return #NA
        End If
    Else
        If RecachedValues.Exists(key) Then
            If TypeName(RecachedValues(key)) = "Collection" Then
                Set GetCachedValue = RecachedValues(key)
            Else
                GetCachedValue = RecachedValues(key)
            End If
        Else
            GetCachedValue = CVErr(xlErrNA) ' return #NA
        End If
    End If
End Function


