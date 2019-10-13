Function Bytes2Lng(a() As Byte) As Double
   Bytes2Lng = ((a(0) * 256#  + a(1)) * 256# + a(2)) * 256# + a(3)
End Function
