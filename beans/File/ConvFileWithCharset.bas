Sub ConvFileWithCharset(ByVal sFromFile As String, ByVal sToFile As String, sFromCharset As String, Optional sToCharset As String = "Shift-JIS", Optional bCrLf As Boolean = True)
   Dim stream  As Object
   Dim sText As String
   Dim bin As Variant
   Dim bUTF8 As Boolean

   Set stream = CreateObject("ADODB.Stream")
   stream.Type = 2 ' adTypeText(2)
   stream.Charset = sFromCharset
   Call stream.Open()
   Call stream.LoadFromFile(sFromFile)
   sText = stream.ReadText
   Call stream.Close()

   If bCrLf Then
      sText = Replace(sText, vbLf, vbCrLf)
      sText = Replace(sText, vbCr & vbCr, vbCr)
   End If

   bUTF8 = (sToCharset = "UTF-8")
   Set stream = CreateObject("ADODB.Stream")
   stream.Type = 2 ' adTypeText(2)
   stream.Charset = sToCharset
   Call stream.Open()
   Call stream.WriteText(sText)

   If bUTF8 Then ' UTF-8, then remove BOM
      stream.Position = 0
      stream.Type = 1 ' adTypeBinary
      stream.Position = 3
      bin = stream.Read()
      Call stream.Close()
      Set stream = CreateObject("ADODB.Stream")
      stream.Type = 1 ' dTypeBinary
      Call stream.Open()
      Call stream.Write(bin)
   End If

   Call stream.SaveToFile(sToFile, 2) ' adSaveCreateOverWrite(2)
   Call stream.Close()
End Sub

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ConvFileWithCharset()
   Call ConvFileWithCharset(ThisWorkbook.Path & "/test_SJIS.txt", ThisWorkbook.Path & "/test_UTF8_2.txt", "Shift-JIS", "UTF-8", bCrLf:=True)
   assert False, DiffFile(ThisWorkbook.Path & "/test_SJIS.txt", ThisWorkbook.Path & "/test_UTF8_2.txt")
   assert True, DiffFile(ThisWorkbook.Path & "/test_UTF8.txt", ThisWorkbook.Path & "/test_UTF8_2.txt")
   Call ConvFileWithCharset(ThisWorkbook.Path & "/test_SJIS.txt", ThisWorkbook.Path & "/test_UTF8_3.txt", "Shift-JIS", "UTF-8")
   assert True, DiffFile(ThisWorkbook.Path & "/test_UTF8.txt", ThisWorkbook.Path & "/test_UTF8_3.txt")
End Sub
