Function GetFolderDesktop() As String
   Dim oWSH As Object
   Set oWSH = CreateObject("WScript.Shell")
   GetFolderDesktop = oWSH.SpecialFolders("Desktop")
   Set oWSH = Nothing
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetFolderDesktop()
   Dim oWSH As Object, oEnv As Object
   Set oWSH = CreateObject("WScript.Shell")
   assert StrConv(oWSH.ExpandEnvironmentStrings("%USERPROFILE%\Desktop"), vbUpperCase), StrConv(GetFolderDesktop(), vbUpperCase)
   Set oWSH = Nothing
End Sub
