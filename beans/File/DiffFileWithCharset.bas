Function DiffFileWithCharset(ByVal sFile1 As String, sCharset1 As String, ByVal sFile2 As String, sCharset2 As String) As Boolean
   Dim streamRead  As Object
   Dim s1 As String, s2 As String
   On Error GoTo ErrExit

   Set streamRead = CreateObject("ADODB.Stream")
   streamRead.Type = 2 ' adTypeText(2)
   streamRead.Charset = sCharset1
   Call streamRead.Open()
   Call streamRead.LoadFromFile(sFile1)
   s1 = streamRead.ReadText()
   s1 = Replace(s1, vbLf, vbCrLf)
   s1 = Replace(s1, vbCr & vbCr, vbCr)
   Call streamRead.Close()
   Set streamRead = Nothing

   Set streamRead = CreateObject("ADODB.Stream")
   streamRead.Type = 2 ' adTypeText(2)
   streamRead.Charset = sCharset2
   Call streamRead.Open()
   Call streamRead.LoadFromFile(sFile2)
   s2 = streamRead.ReadText()
   s2 = Replace(s2, vbLf, vbCrLf)
   s2 = Replace(s2, vbCr & vbCr, vbCr)
   Call streamRead.Close()
   Set streamRead = Nothing

   DiffFileWithCharset = (s1 = s2)
ExitProc:
   Exit Function
ErrExit:
   If Not streamRead Is Nothing Then Call streamRead.Close()
   DiffFileWithCharset = False
   Resume ExitProc
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_DiffFileWithCharset()
   assert True, DiffFileWithCharset(ThisWorkbook.Path & "/test11.txt", "Shift-JIS", ThisWorkbook.Path & "/test11.txt", "Shift-JIS")
   assert True, DiffFileWithCharset(ThisWorkbook.Path & "/test11.txt", "Shift-JIS", ThisWorkbook.Path & "/test12.txt", "Shift-JIS")
   assert False, DiffFileWithCharset(ThisWorkbook.Path & "/test11.txt", "Shift-JIS", ThisWorkbook.Path & "/test2.txt", "Shift-JIS")
   assert True, DiffFileWithCharset(ThisWorkbook.Path & "/test_SJIS.txt", "Shift-JIS", ThisWorkbook.Path & "/test_UTF8.txt", "UTF-8")
   assert True, DiffFileWithCharset(ThisWorkbook.Path & "/test_SJIS.txt", "Shift-JIS", ThisWorkbook.Path & "/test_EUC_JP.txt", "EUC-JP")
   assert False, DiffFileWithCharset(ThisWorkbook.Path & "/test_SJIS.txt", "Shift-JIS", ThisWorkbook.Path & "/test_UTF8.txt", "Shift-JIS")
End Sub
