' VBA: Import ../GetFSO.bas
' VBA: Import ../UniRegExp.bas

Function GetCollectionGrepPath(ByVal sBasePath As String, Optional ByVal sReg As String = "", Optional ByVal bFullPath As Boolean = False, Optional ByVal bRecursive As Boolean = True) As Collection
   Dim s As String
   Dim oCol As New Collection

   If sReg <> "" Then
      Call GetCollectionGrepPathSub(sBasePath, ".", oCol, bFullPath, bRecursive, UniRegExp(sReg))
   Else
      Call GetCollectionGrepPathSub(sBasePath, ".", oCol, bFullPath, bRecursive, Nothing)
   End If

   Set GetCollectionGrepPath = oCol
End Function

Private Sub GetCollectionGrepPathSub(ByVal sBasePath As String, ByVal sDir As String, ByRef oCol As Object, ByVal bFullPath As Boolean, ByVal bRecursive As Boolean, ByVal oReg As Object)
   Dim v As Variant
   Dim sPath As String
   With GetFSO()
      sPath = .BuildPath(sBasePath, sDir)
      If sDir = "." Then sDir = ""

      If bRecursive Then
         For Each v In .GetFolder(sPath).SubFolders
            Call GetCollectionGrepPathSub(sBasePath, .BuildPath(sDir, v.Name), oCol, bFullPath, bRecursive, oReg)
         Next v
      End If
      For Each v In .GetFolder(sPath).Files
         If bFullPath Then
            sPath = v.Path
         Else
            sPath = .BuildPath(sDir, v.Name)
         End If
         If oReg Is Nothing Then
            oCol.Add sPath
         Else
            If oReg.Test(sPath) Then oCol.Add sPath
         End If
      Next v
   End With
End Sub

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetCollectionGrepPath()
   assert "[""test_UTF8_3.txt"",""test2.bin"",""test2.txt"",""test_UTF8_2.txt"",""TestBook.xlsm"",""test12.bin"",""test12.txt"",""test.ini"",""test11.bin"",""test11.txt"",""test.xlsx"",""test_EUC_JP.txt"",""VBA_Manager.ini"",""~$TestBook.xlsm"",""test_SJIS2.txt"",""ezy_log.log"",""test_UTF8.txt"",""test_SJIS.txt""]", Dump(GetCollectionGrepPath(ThisWorkbook.Path, bRecursive:=False))
   assert "[""src\GlobalTest.bas"",""src\Z_BEANS.bas"",""test_UTF8_3.txt"",""test2.bin"",""test2.txt"",""test_UTF8_2.txt"",""TestBook.xlsm"",""test12.bin"",""test12.txt"",""test.ini"",""test11.bin"",""test11.txt"",""test.xlsx"",""test_EUC_JP.txt"",""VBA_Manager.ini"",""~$TestBook.xlsm"",""test_SJIS2.txt"",""ezy_log.log"",""test_UTF8.txt"",""test_SJIS.txt""]", Dump(GetCollectionGrepPath(ThisWorkbook.Path, bRecursive:=True))
   assert "[""test_UTF8_3.txt"",""test2.txt"",""test_UTF8_2.txt"",""test12.txt"",""test11.txt"",""test_EUC_JP.txt"",""test_SJIS2.txt"",""test_UTF8.txt"",""test_SJIS.txt""]", Dump(GetCollectionGrepPath(ThisWorkbook.Path, sReg:=".*\.txt", bRecursive:=True))
   assert "[]", Dump(GetCollectionGrepPath(ThisWorkbook.Path, bFullPath:=True, sReg:=".*x.*\.txt", bRecursive:=True))
   'Debug.Print Dump(GetCollectionGrepPath(ThisWorkbook.Path, bFullPath:=True, sReg:=".*t.*\.txt", bRecursive:=True))
End Sub
