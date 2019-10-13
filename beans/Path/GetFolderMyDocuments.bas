Function GetFolderMyDocuments() As String
   Dim oWSH As Object
   Set oWSH = CreateObject("WScript.Shell")
   GetFolderMyDocuments = oWSH.SpecialFolders("MyDocuments")
   Set oWSH = Nothing
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetFolderMyDocuments()
   Dim oWSH As Object, oEnv As Object
   Set oWSH = CreateObject("WScript.Shell")
   assert StrConv(oWSH.ExpandEnvironmentStrings("%USERPROFILE%\Documents"), vbUpperCase), StrConv(GetFolderMyDocuments(), vbUpperCase)
   Set oWSH = Nothing
End Sub
