#If Win64 Then
Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, _
         ByVal wParam As Long, lParam As Any) As Long
#Else
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, _
         ByVal wParam As Long, lParam As Any) As Long
#End If
