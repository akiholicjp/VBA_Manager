#If Win64 Then
Declare PtrSafe Function GetDlgCtrlID Lib "user32.dll" ( _
                                                ByVal hWnd As Long _
                                                ) As Long
#Else
Declare Function GetDlgCtrlID Lib "user32.dll" ( _
                                                ByVal hWnd As Long _
                                                ) As Long
#End If
