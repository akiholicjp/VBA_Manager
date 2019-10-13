#If Win64 Then
Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
                                                                      ByVal hwndParent As Long, _
                                                                      ByVal hwndChildAfter As Long, _
                                                                      ByVal lpszClass As String, _
                                                                      ByVal lpszWindow As String _
                                                                      ) As Long
#Else
Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
                                                                      ByVal hwndParent As Long, _
                                                                      ByVal hwndChildAfter As Long, _
                                                                      ByVal lpszClass As String, _
                                                                      ByVal lpszWindow As String _
                                                                      ) As Long
#End If
