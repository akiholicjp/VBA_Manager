' VBA: Import ../GetFSO.bas

Function GetBaseFileName(ByVal sPath As String) As String
   With GetFSO()
      GetBaseFileName = .GetBaseName(sPath) & "." & .GetExtensionName(sPath)
   End With
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetBaseFileName()
   assert "test.bas", GetBaseFileName("c:\Test\test.bas")
   assert "test.bas", GetBaseFileName("..\test.bas")
   assert "test.bas", GetBaseFileName("./test.bas")
   assert "hogehoge.dat", GetBaseFileName("./test.bas/hogehoge.dat")
   assert "hogehoge.dat", GetBaseFileName("~/hogehoge.dat")
End Sub
