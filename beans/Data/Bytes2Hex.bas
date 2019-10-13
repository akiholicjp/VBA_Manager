Function Bytes2Hex(bytes As Variant, Optional ByVal sDelim As String = "") As String
   Dim tmp As Variant
   For Each tmp In bytes
      If Bytes2Hex = "" Then
         Bytes2Hex = Right("0" & Hex(tmp), 2)
      Else
         Bytes2Hex = Bytes2Hex & sDelim & Right("0" & Hex(tmp), 2)
      End If
   Next tmp
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_Bytes2Hex()
   Dim data() As Byte

   ReDim data(1) As Byte
   data(0) = &H5
   data(1) = &HA5
   assert "05A5", Bytes2Hex(data)
   assert "05 A5", Bytes2Hex(data, " ")
   assert "05 A5", Bytes2Hex(data, sDelim:=" ")
   assert "05_A5", Bytes2Hex(data, sDelim:="_")

   ReDim data(2) As Byte
   data(0) = &H5
   data(1) = &HA5
   data(2) = &H9

   assert "05A509", Bytes2Hex(data)
   assert "05 A5 09", Bytes2Hex(data, " ")
   assert "05 A5 09", Bytes2Hex(data, sDelim:=" ")
   assert "05_A5_09", Bytes2Hex(data, sDelim:="_")

   data(0) = &H0
   data(1) = &HA5
   data(2) = &H0
   assert "00A500", Bytes2Hex(data)
   assert "00 A5 00", Bytes2Hex(data, " ")
   assert "00 A5 00", Bytes2Hex(data, sDelim:=" ")
   assert "00_A5_00", Bytes2Hex(data, sDelim:="_")
End Sub
