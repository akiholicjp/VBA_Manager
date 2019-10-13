' VBA: Import ../GetShell.bas
Function GetCurrentPath() As String
   GetCurrentPath = GetShell().CurrentDirectory
End Function

' =================== VBA: TEST: Begin ===================

' VBA: Import ../GetShell.bas

Public Sub xUnitTest_beans_GetCurrentPath()
   assert GetShell().CurrentDirectory, GetCurrentPath()
End Sub
