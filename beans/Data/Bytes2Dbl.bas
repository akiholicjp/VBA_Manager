' VBA: Import _TypeDbl.bas

Function Bytes2Dbl(a() As Byte) As Double
   Dim tmp As MyDblBytes
   Dim tmp2 As MyDblDouble
   tmp.b1 = a(7)
   tmp.b2 = a(6)
   tmp.b3 = a(5)
   tmp.b4 = a(4)
   tmp.b5 = a(3)
   tmp.b6 = a(2)
   tmp.b7 = a(1)
   tmp.b8 = a(0)
   LSet tmp2 = tmp
   Bytes2Dbl = tmp2.d
End Function
