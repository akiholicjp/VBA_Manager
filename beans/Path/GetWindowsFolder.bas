' VBA: Import ../GetFSO.bas

Function GetWindowsFolder() As String
   GetWindowsFolder = GetFSO().GetSpecialFolder(0).Path
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetWindowsFolder()
   Dim oWSH As Object, oEnv As Object
   Set oWSH = CreateObject("WScript.Shell")
   assert StrConv(oWSH.ExpandEnvironmentStrings("%windir%"), vbUpperCase), StrConv(GetWindowsFolder(), vbUpperCase)
End Sub
