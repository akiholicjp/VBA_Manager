#If Win64 Then
Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" _
   Alias "WNetGetConnectionA" _
   (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
#Else
Declare Function WNetGetConnection Lib "mpr.dll" _
   Alias "WNetGetConnectionA" _
   (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
#End If

' =================== VBA: TEST: Begin ===================

Function WNetGetConnectionWrap(ByVal strTarget As String) As String
   Const C_MAX_LENGTH = 1023
   Dim s As String
   s = String$(C_MAX_LENGTH + 1, vbNullChar)
   If WNetGetConnection(strTarget, s, C_MAX_LENGTH) = 0 Then
      WNetGetConnectionWrap = Left(s, InStr(s, vbNullChar) - 1)
   Else
      WNetGetConnectionWrap = ""
   End If
End Function
