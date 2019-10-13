Function WSH_Popup(Text As String, Optional ByVal SecondsToWait As Long, Optional ByVal Title As String, Optional ByVal iType As Long = vbOKOnly) As Long
   Dim WSH As Object
   Set WSH = CreateObject("WScript.Shell")
   WSH_Popup = WSH.Popup(Text, SecondsToWait, Title, iType)
   Set WSH = Nothing
End Function

