Function IsOctStr(ByVal a_inStr As String) As Boolean
   Dim i As Long
   If a_inStr = "" Then
      IsOctStr = False
      Exit Function
   End If
   IsOctStr = True
   For i = 1 To Len(a_inStr) Step 16
      IsOctStr = IsOctStr And IsNumeric("&O" & Mid(a_inStr, i, 16))
   Next i
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_IsOctStr()
   assert True, IsOctStr("10101010")
   assert True, IsOctStr("1010101010101010")
   assert True, IsOctStr("10101010101010101010101010101010")
   assert True, IsOctStr("1010101010101010101010101010101010101010101010101010101010101010")
   assert True, IsOctStr("10101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010")
   assert True, IsOctStr("01234567")
   assert True, IsOctStr("012345670123456701234567012345670123456701234567012345670123456701234567")
   assert False, IsOctStr("0123456789")
   assert False, IsOctStr("012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789")
   assert False, IsOctStr("0123456789abcdef")
   assert False, IsOctStr("0123456789ABCDEF")
   assert False, IsOctStr("0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF")
   assert False, IsOctStr("x")
   assert False, IsOctStr("")
End Sub
