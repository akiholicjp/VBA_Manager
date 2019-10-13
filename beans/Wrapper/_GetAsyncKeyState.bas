#If Win64 Then
Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long
#Else
Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long
#End If
