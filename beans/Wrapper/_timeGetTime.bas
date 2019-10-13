#If Win64 Then
Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If
