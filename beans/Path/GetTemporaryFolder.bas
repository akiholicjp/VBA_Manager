' VBA: Import ../GetFSO.bas

Function GetTemporaryFolder() As String
   GetTemporaryFolder = GetFSO().GetSpecialFolder(2).Path
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetTemporaryFolder()
   Dim oWSH As Object, oEnv As Object
   Set oWSH = CreateObject("WScript.Shell")
   assert StrConv(oWSH.ExpandEnvironmentStrings("%USERPROFILE%\AppData\Local\Temp"), vbUpperCase), StrConv(GetTemporaryFolder(), vbUpperCase)
End Sub
