' VBA: Import _TypeDbl.bas

Function Hex2Dbl(ByVal s As String) As Double
   Dim Lng As MyDblDLong, Dbl As MyDblDouble, tmp As String
   tmp = Right(String(16, "0") & s, 16)
   Lng.L1 = CLng("&H0" & Right(tmp, 8))
   Lng.L2 = CLng("&H0" & Left(tmp, 8))
   LSet Dbl = Lng
   Hex2Dbl = Dbl.d
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_Hex2Dbl()
   assert 1, Hex2Dbl("3FF0000000000000")
   assert 1.12345, Hex2Dbl("3FF1F9A6B50B0F28")
   assert 1.12345E+99, Hex2Dbl("54806FB414ADEF8E")
   assert -9.87654321E+99, Hex2Dbl("D4B20FE0BCD31B4B")
End Sub
