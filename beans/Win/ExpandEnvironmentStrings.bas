Function ExpandEnvironmentStrings(ByVal str As String) As String
   Dim oWSH As Object
   Set oWSH = CreateObject("WScript.Shell")
   ExpandEnvironmentStrings = oWSH.ExpandEnvironmentStrings(str)
   Set oWSH = Nothing
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ExpandEnvironmentStrings()
   Dim oWSH As Object, oEnv As Object
   Set oWSH = CreateObject("WScript.Shell")
   Set oEnv = oWSH.Environment("Process")

   oEnv("BBB") = "zzz"
   oEnv("AAA") = "xxx%BBB%yyy"

   assert "xxx%BBB%yyy", oEnv("AAA")
   assert "xxxzzzyyy", ExpandEnvironmentStrings("xxx%BBB%yyy")
   assert "zzz", ExpandEnvironmentStrings("%BBB%")
   assert "xxx%BBB%yyy", ExpandEnvironmentStrings("%AAA%")
   assert "xxxzzzyyy", ExpandEnvironmentStrings(ExpandEnvironmentStrings("%AAA%"))
End Sub
