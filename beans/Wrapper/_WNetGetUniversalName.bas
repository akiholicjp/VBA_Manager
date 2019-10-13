#If Win64 Then
Declare PtrSafe Function WNetGetUniversalName Lib "mpr.dll" _
   Alias "WNetGetUniversalNameA" _
   (ByVal lpLocalPath As String, ByVal dwInfoLevel As Long, lpBuffer As Any, lpBufferSize As Long) As Long
#Else
Declare Function WNetGetUniversalName Lib "mpr.dll" _
   Alias "WNetGetUniversalNameA" _
   (ByVal lpLocalPath As String, ByVal dwInfoLevel As Long, lpBuffer As Any, lpBufferSize As Long) As Long
#End If
