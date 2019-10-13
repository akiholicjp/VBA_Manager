' VBA: Import ../GetFSO.bas
' VBA: Import IsAbsolutePath.bas

Function GetAbsolutePath(ByVal pathFile As String, Optional ByVal sBaseDir As String = "") As String
   With GetFSO()
      If IsAbsolutePath(pathFile) Then
         GetAbsolutePath = pathFile
         Exit Function
      End If
      If sBaseDir <> "" Then
         pathFile = .BuildPath(sBaseDir, pathFile)
      End If
      GetAbsolutePath = .GetAbsolutePathName(pathFile)
   End With
End Function

' =================== VBA: TEST: Begin ===================

' VBA: Import GetOwnPath.bas
' VBA: Import SetCurrentPath.bas

Public Sub xUnitTest_beans_GetAbsolutePath()
   assert StrConv(GetOwnPath(), vbUpperCase), StrConv(GetAbsolutePath(GetOwnPath()), vbUpperCase)
   SetCurrentPath(GetOwnPath())
   assert StrConv(GetOwnPath(), vbUpperCase), StrConv(GetAbsolutePath("."), vbUpperCase)
End Sub
