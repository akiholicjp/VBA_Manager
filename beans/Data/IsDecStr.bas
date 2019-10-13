Function IsDecStr(ByVal a_inStr As String) As Boolean
   Dim i As Long
   If a_inStr = "" Then
      IsDecStr = False
      Exit Function
   End If
   IsDecStr = True
   For i = 1 To Len(a_inStr) Step 16
      IsDecStr = IsDecStr And IsNumeric(Mid(a_inStr, i, 16))
   Next i
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_IsDecStr()
   assert True, IsDecStr("10101010")
   assert True, IsDecStr("1010101010101010")
   assert True, IsDecStr("10101010101010101010101010101010")
   assert True, IsDecStr("1010101010101010101010101010101010101010101010101010101010101010")
   assert True, IsDecStr("10101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010")
   assert True, IsDecStr("0123456789")
   assert True, IsDecStr("012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789")
   assert False, IsDecStr("0123456789abcdef")
   assert False, IsDecStr("0123456789ABCDEF")
   assert False, IsDecStr("x")
   assert False, IsDecStr("")
End Sub
