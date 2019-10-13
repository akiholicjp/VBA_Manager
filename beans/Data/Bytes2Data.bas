Function Bytes2Data(data As Variant, ByVal iLen As Integer, Optional ByVal f2Comp As Boolean = False) As Double
   Dim vRet As Double, fSgn As Boolean
   Dim i As Integer
   If f2Comp Then fSgn = (data(0) And 2 ^ ((iLen - 1) Mod 8)) <> 0
   For i = 0 To Int((iLen - 1) / 8)
      If fSgn Then
         If i = 0 Then
            vRet = vRet * &H100 + (data(i) Xor (2 ^ (((iLen - 1) Mod 8) + 1) - 1))
         Else
            vRet = vRet * &H100 + (data(i) Xor &HFF)
         End If
      Else
         vRet = vRet * &H100 + data(i)
      End If
   Next i
   If fSgn Then
      Bytes2Data = -(vRet + 1)
   Else
      Bytes2Data = vRet
   End If
End Function
