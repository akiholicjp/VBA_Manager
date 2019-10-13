' VBA: Import ../GetFSO.bas

Function IsAbsolutePath(ByVal s As String) As Boolean
   Dim c As String
   s = Trim(s)
   If s = "" Then
      IsAbsolutePath = False
   ElseIf s Like "*:/" Or s Like "*:\" Then
      IsAbsolutePath = True
   ElseIf s Like "/*" OR s Like "\*" Then
      IsAbsolutePath = True
   Else
      IsAbsolutePath = (StrConv(s, vbUpperCase) = StrConv(GetFSO().GetAbsolutePathName(s), vbUpperCase))
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_IsAbsolutePath()
   assert True, IsAbsolutePath("C:\")
   assert True, IsAbsolutePath("/.")
   assert False, IsAbsolutePath(".")
   assert False, IsAbsolutePath("")
End Sub
