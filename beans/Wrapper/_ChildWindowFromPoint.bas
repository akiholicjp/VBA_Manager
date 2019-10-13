#If Win64 Then
Declare PtrSafe Function ChildWindowFromPoint Lib "user32.dll" ( _
                                                        ByVal hwndParent As Long, _
                                                        ByVal x As Long, _
                                                        ByVal y As Long _
                                                        ) As Long
#Else
Declare Function ChildWindowFromPoint Lib "user32.dll" ( _
                                                        ByVal hwndParent As Long, _
                                                        ByVal x As Long, _
                                                        ByVal y As Long _
                                                        ) As Long
#End If
