' VBA: Import _TypeSng.bas

Function Hex2Sng(ByVal s As String) As Single
   Dim Lng As MySngDLong, Sng As MySngSingle
   Lng.L1 = CLng("&H0" & s)
   LSet Sng = Lng
   Hex2Sng = Sng.s
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_Hex2Sng()
   assert 1, Hex2Sng("3F800000")
   assert 1.12345, Hex2Sng("3F8FCD36")
   assert 1.12345E+19, Hex2Sng("5F1BE8F6")
   assert -9.87654321E+19, Hex2Sng("E0AB54AA")
End Sub
