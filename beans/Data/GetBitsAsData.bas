' VBA: Import GetBits.bas
' VBA: Import Bytes2Data.bas

Function GetBitsAsData(data As Variant, ByVal iWord As Integer, ByVal iBit As Integer, ByVal iLen As Integer, Optional ByVal f2Comp As Boolean = False) As Double
   GetBitsAsData = Bytes2Data(GetBits(data, iWord, iBit, iLen), iLen, f2Comp)
End Function
