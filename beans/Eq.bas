Function Eq(ByVal x As Variant, ByVal y As Variant) As Boolean
   If IsObject(x) Xor IsObject(y) Then
      Eq = False ' Empty
   ElseIf IsObject(x) And IsObject(y) Then
      Eq = (x Is y)
   Else
      Eq = (x = y)
   End If
End Function
