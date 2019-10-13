Function IsHexStr(ByVal a_inStr As String) As Boolean
   Dim i As Long
   If a_inStr = "" Then
      IsHexStr = False
      Exit Function
   End If
   IsHexStr = True
   a_inStr = StrConv(a_inStr, vbUpperCase)
   For i = 1 To Len(a_inStr) Step 16
      IsHexStr = IsHexStr And IsNumeric("&H" & Mid(a_inStr, i, 16))
   Next i
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_IsHexStr()
   assert True, IsHexStr("10101010")
   assert True, IsHexStr("1010101010101010")
   assert True, IsHexStr("10101010101010101010101010101010")
   assert True, IsHexStr("1010101010101010101010101010101010101010101010101010101010101010")
   assert True, IsHexStr("10101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010")
   assert True, IsHexStr("0123456789")
   assert True, IsHexStr("012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789")
   assert True, IsHexStr("0123456789abcdef")
   assert True, IsHexStr("0123456789ABCDEF")
   assert True, IsHexStr("0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF")
   assert False, IsHexStr("x")
   assert False, IsHexStr("")
End Sub
