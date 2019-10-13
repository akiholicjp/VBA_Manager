' VBA: Import _TypeSng.bas

Function Bytes2Sng(a() As Byte) As Single
   Dim tmp As MySngBytes
   Dim tmp2 As MySngSingle
   tmp.b1 = a(3)
   tmp.b2 = a(2)
   tmp.b3 = a(1)
   tmp.b4 = a(0)
   LSet tmp2 = tmp
   Bytes2Sng = tmp2.s
End Function
