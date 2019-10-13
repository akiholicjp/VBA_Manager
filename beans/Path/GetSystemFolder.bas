' VBA: Import ../GetFSO.bas

Function GetSystemFolder() As String
   GetSystemFolder = GetFSO().GetSpecialFolder(1).Path
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetSystemFolder()
   assert StrConv("C:\Windows\System32", vbUpperCase), StrConv(GetSystemFolder(), vbUpperCase)
End Sub
