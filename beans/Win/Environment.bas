Function Environment(ByVal sType As String) As Object
   Dim oWSH As Object
   Set oWSH = CreateObject("WScript.Shell")
   Set Environment = oWSH.Environment(sType)
   Set oWSH = Nothing
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_Environment()
   Dim o As Object

   Set o = Environment("Process")
   o("TEST") = "AAA"
   assert "AAA", o("TEST")

   Set o = Environment("Volatile")
   o("TEST") = "BBB"
   assert "BBB", o("TEST")

   Set o = Environment("User")
   o("TEST") = "CCC"
   assert "CCC", o("TEST")

   o.Remove("TEST")

   Set o = Environment("Process")
   assert "AAA", o("TEST")

   o.Remove("TEST")

   assert "", o("TEST")

   Set o = Environment("Volatile")
   assert "BBB", o("TEST")

   o.Remove("TEST")

   assert "", o("TEST")

   Set o = Environment("System")
End Sub
