Function GetOwnPath() As String
#If ACCESS_VBA <> 1 Then
   GetOwnPath = ThisWorkbook.Path
#Else
   GetOwnPath = Application.CurrentProject.Path
#End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetOwnPath()
   assert StrConv(ThisWorkbook.Path, vbUpperCase), StrConv(GetOWnPath(), vbUpperCase)
End Sub
