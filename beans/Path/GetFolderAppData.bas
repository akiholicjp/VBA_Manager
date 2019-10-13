Function GetFolderAppData() As String
   Dim oWSH As Object
   Set oWSH = CreateObject("WScript.Shell")
   GetFolderAppData = oWSH.SpecialFolders("AppData")
   Set oWSH = Nothing
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetFolderAppData()
   Dim oWSH As Object, oEnv As Object
   Set oWSH = CreateObject("WScript.Shell")
   assert StrConv(oWSH.ExpandEnvironmentStrings("%USERPROFILE%\AppData\Roaming"), vbUpperCase), StrConv(GetFolderAppData(), vbUpperCase)
   Set oWSH = Nothing
End Sub
