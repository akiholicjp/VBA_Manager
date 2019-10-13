Function XorHex(ByVal str As String, Optional ByVal iLen As String = 2) As String
   Dim ret As Long
   Dim tmp As String
   ret = 0
   Do While Len(str) > iLen * 2
      tmp = Left(str, iLen * 2)
      str = Mid(str, iLen * 2 + 1)
      ret = ret Xor CLng("&H" & tmp)
   Loop
   ret = ret Xor CLng("&H" & str)
   XorHex = Hex(ret)
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_XorHex()
   assert "0", XorHex("12341234", 2)
   assert "1000", XorHex("12340234", 2)
End Sub
