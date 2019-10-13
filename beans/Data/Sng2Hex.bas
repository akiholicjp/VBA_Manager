' VBA: Import _TypeSng.bas

Function Sng2Hex(ByVal s As Single) As String
   Dim Lng As MySngDLong, Sng As MySngSingle
   Sng.s = s
   LSet Lng = Sng
   Sng2Hex = Right("00000000" & Hex(Lng.L1), 8)
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_Sng2Hex()
   assert "3F800000", Sng2Hex(1)
   assert "3F8FCD36", Sng2Hex(1.12345)
   assert "5F1BE8F6", Sng2Hex(1.12345E+19)
   assert "E0AB54AA", Sng2Hex(-9.87654321E+19)
End Sub
