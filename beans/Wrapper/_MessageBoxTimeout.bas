#If VBA7 And Win64 Then
Declare PtrSafe Function MessageBoxTimeoutA Lib "User32" ( _
   ByVal Hwnd As Long, _
   ByVal lpText As String, _
   ByVal lpCaption As String, _
   ByVal uType As VbMsgBoxStyle, _
   ByVal wLanguageID As Long, _
   ByVal dwMilliseconds As Long) As Long
#Else
Declare Function MessageBoxTimeoutA Lib "User32"( _
   ByVal Hwnd As Long, _
   ByVal lpText As String, _
   ByVal lpCaption As String, _
   ByVal uType As VbMsgBoxStyle, _
   ByVal wLanguageID As Long, _
   ByVal dwMilliseconds As Long) As Long
#End If
