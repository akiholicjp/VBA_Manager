Function BinToHex(ByVal sBin As String) As String
   Dim s4Bits As String
   Dim iPos As Long
   Dim iBitPos As Long
   Dim iVal As Long
   sBin = Replace(Replace(sBin, " ", ""), vbLf, "")
   sBin = StrReverse(sBin)
   BinToHex = ""
   For iPos = 1 To Len(sBin) Step 4
      s4Bits = Mid(sBin, iPos, 4)
      iVal = 0
      For iBitPos = 1 To Len(s4Bits)
         If Mid(s4Bits, iBitPos, 1) = "1" Then
         iVal = iVal + 2 ^ (iBitPos - 1)
         End If
      Next iBitPos
      BinToHex = Mid("0123456789ABCDEF", iVal + 1, 1) & BinToHex
   Next iPos
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_BinToHex()
   assert "C", BinToHex("1100")
   assert "0C", BinToHex("01100")
   assert "4", BinToHex("0 1 0 0")
   assert "2", BinToHex("10")
   assert "10", BinToHex("10000")
   assert "100", BinToHex("100000000")
   assert "7FFFFFFF", BinToHex("01111111 11111111 11111111 11111111")
End Sub
