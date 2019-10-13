Function BinToDec(ByVal sBin As String) As Long
   Dim retVal As Double
   Dim i As Long
   sBin = Replace(Replace(sBin, " ", ""), vbLf, "")
   If (Not Len(sBin) > 1023) Then
      If (sBin <> "") And (Replace(Replace(sBin, "1", ""), "0", "") = "") Then
         sBin = StrReverse(sBin)
         For i = 0 To Len(sBin)
            If (Mid(sBin, i + 1, 1) = 1) Then
               retVal = retVal + (2 ^ (i))
            End If
         Next i
      End If
      BinToDec = retVal
   Else
      MsgBox "Overflow would occur"
      Exit Function
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_BinToDec()
   assert &H0C, BinToDec("1100")
   assert &H4, BinToDec("0 1 0 0")
   assert &H2, BinToDec("10")
   assert &H10, BinToDec("10000")
   assert &H100, BinToDec("100000000")
   assert &H7FFFFFFF, BinToDec("01111111 11111111 11111111 11111111")
End Sub
