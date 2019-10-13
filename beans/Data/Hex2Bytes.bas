Function Hex2Bytes(ByVal str As String, Optional ByVal bZeroPadding As Boolean = True) As Byte()
   Dim n As Long
   Dim ret() As Byte

   If Len(str) Mod 2 <> 0 Then
      If bZeroPadding Then
         str = "0" & str
      Else
         Hex2Bytes = WorksheetFunction.NA
         Exit Function
      End If
   End If
   n = Len(str) / 2

   ReDim ret(n - 1) As Byte
   Do While n <> 0
      ret(n - 1) = CByte("&H" & Mid(str, n * 2 - 1, 2))
      n = n - 1
   Loop
   Hex2Bytes = ret
End Function
