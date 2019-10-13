' VBA: Import _TypeDbl.bas

Function Dbl2Hex(ByVal d As Double) As String
   Dim Lng As MyDblDLong, Dbl As MyDblDouble
   Dbl.d = d
   LSet Lng = Dbl
   Dbl2Hex = Right("00000000" & Hex(Lng.L2), 8) & Right("00000000" & Hex(Lng.L1), 8)
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_Dbl2Hex()
   assert "3FF0000000000000", Dbl2Hex(1)
   assert "3FF1F9A6B50B0F28", Dbl2Hex(1.12345)
   assert "54806FB414ADEF8E", Dbl2Hex(1.12345E+99)
   assert "D4B20FE0BCD31B4B", Dbl2Hex(-9.87654321E+99)
End Sub
