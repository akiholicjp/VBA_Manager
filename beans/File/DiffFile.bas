Function DiffFile(ByVal sFromFile As String, ByVal sToFile As String, Optional ByVal bBinary As Boolean = False) As Boolean
   Dim fp1 As Long, fp2 As Long
   Dim i1 As Long, i2 As Long, i As Long
   Dim s1 As String, s2 As String
   Dim b1() As Byte, b2() As Byte
   Dim bRet As Boolean
   fp1 = 0
   fp2 = 0
   On Error GoTo ErrExit
   If bBinary Then
      fp1 = FreeFile(): Open sFromFile For Binary Access Read As fp1
      fp2 = FreeFile(): Open sToFile For Binary Access Read As fp2
      ReDim b1(LOF(fp1))
      ReDim b2(LOF(fp2))
   Else
      fp1 = FreeFile(): Open sFromFile For Input As fp1
      fp2 = FreeFile(): Open sToFile For Input As fp2
   End If
   If bBinary Then
      Get #fp1, , b1
      Get #fp2, , b2
      bRet = (CStr(b1) = CStr(b2))
   Else
      bRet = True
      Do Until EOF(fp1) Or EOF(fp2)
         Line Input #fp1, s1
         Line Input #fp2, s2
         If s1 = s2 Then
            i1 = i1 + 1
            i2 = i2 + 1
         Else
            bRet = False
            Exit Do
         End If
      Loop
      If Not (EOF(fp1) And EOF(fp2)) Then bRet = False
   End If
ExitProc:
   If fp1 > 0 Then Close fp1
   If fp2 > 0 Then Close fp2
   DiffFile = bRet
   Exit Function
ErrExit:
   bRet = False
   Resume ExitProc
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_DiffFile()
   assert True, DiffFile(ThisWorkbook.Path & "/test11.txt", ThisWorkbook.Path & "/test11.txt")
   assert True, DiffFile(ThisWorkbook.Path & "/test11.txt", ThisWorkbook.Path & "/test12.txt")
   assert False, DiffFile(ThisWorkbook.Path & "/test11.txt", ThisWorkbook.Path & "/test2.txt")
   assert True, DiffFile(ThisWorkbook.Path & "/test11.bin", ThisWorkbook.Path & "/test11.bin", bBinary:=True)
   assert True, DiffFile(ThisWorkbook.Path & "/test11.bin", ThisWorkbook.Path & "/test12.bin", bBinary:=True)
   assert False, DiffFile(ThisWorkbook.Path & "/test11.bin", ThisWorkbook.Path & "/test2.bin", bBinary:=True)
   assert False, DiffFile(ThisWorkbook.Path & "/test_SJIS.txt", ThisWorkbook.Path & "/test_UTF8.txt")
   assert False, DiffFile(ThisWorkbook.Path & "/test_SJIS.txt", ThisWorkbook.Path & "/test_EUC_JP.txt")
   assert False, DiffFile(ThisWorkbook.Path & "/test_SJIS.txt", ThisWorkbook.Path & "/test_UTF8.txt")
End Sub
