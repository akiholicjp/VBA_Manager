#If Win64 Then
Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" ( _
                                                                  ByVal lpClassName As String, _
                                                                  ByVal lpWindowName As String _
                                                                  ) As Long

#Else
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" ( _
                                                                  ByVal lpClassName As String, _
                                                                  ByVal lpWindowName As String _
                                                                  ) As Long
#End If
