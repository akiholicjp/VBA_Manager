Function ExpandEnvironmentStringsWhole(ByVal str As String) As String
   Dim oWSH As Object
   Dim s As String
   Set oWSH = CreateObject("WScript.Shell")
   Do While True
      s = oWSH.ExpandEnvironmentStrings(str)
      If s = str Then Exit Do
      str = s
   Loop
   ExpandEnvironmentStringsWhole = s
   Set oWSH = Nothing
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ExpandEnvironmentStringsWhole()
   Dim oWSH As Object, oEnv As Object
   Set oWSH = CreateObject("WScript.Shell")
   Set oEnv = oWSH.Environment("Process")

   oEnv("BBB") = "zzz"
   oEnv("AAA") = "xxx%BBB%yyy"

   assert "xxx%BBB%yyy", oEnv("AAA")
   assert "xxxzzzyyy", ExpandEnvironmentStringsWhole("xxx%BBB%yyy")
   assert "zzz", ExpandEnvironmentStringsWhole("%BBB%")
   assert "xxxzzzyyy", ExpandEnvironmentStringsWhole("%AAA%")
End Sub
