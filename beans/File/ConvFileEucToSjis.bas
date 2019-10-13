' VBA: Import ConvFileWithCharset.bas
' VBA: Import DiffFile.bas

Sub ConvFileEucToSjis(ByVal sFrom As String, ByVal sTo As String)
   Call ConvFileWithCharset(sFrom, sTo, "EUC-JP", "Shift-JIS", bCrLf:=True)
End Sub

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ConvFileEucToSjis()
   Call ConvFileEucToSjis(ThisWorkbook.Path & "/test_EUC_JP.txt", ThisWorkbook.Path & "/test_SJIS2.txt")
   assert False, DiffFile(ThisWorkbook.Path & "/test_SJIS.txt", ThisWorkbook.Path & "/test_EUC_JP.txt")
   assert True, DiffFile(ThisWorkbook.Path & "/test_SJIS.txt", ThisWorkbook.Path & "/test_SJIS2.txt")
End Sub
