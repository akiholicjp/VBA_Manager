#If Win64 Then
Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
Declare Function GetTickCount Lib "kernel32" () As Long
#End If
