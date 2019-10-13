Function GetFolderAllUsersDesktop() As String
   Dim oWSH As Object
   Set oWSH = CreateObject("WScript.Shell")
   GetFolderAllUsersDesktop = oWSH.SpecialFolders("AllUsersDesktop")
   Set oWSH = Nothing
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetFolderAllUsersDesktop()
   assert StrConv("C:\Users\Public\Desktop", vbUpperCase), StrConv(GetFolderAllUsersDesktop(), vbUpperCase)
End Sub
