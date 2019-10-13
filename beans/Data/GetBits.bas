Function GetBits(data As Variant, ByVal iWord As Long, ByVal iBit As Long, ByVal iLen As Long) As Byte()
   Dim iShift As Long, iTmp As Long
   Dim iptr As Long, optr As Long

   iShift = 2 ^ (7 - ((iLen + iBit - 1) Mod 8))
   iptr = iWord + Int((iLen + iBit - 1) / 8) + LBound(data)
   optr = Int((iLen - 1) / 8)

   ReDim vRet(optr) As Byte

   Do While iLen + iBit > 8
      iTmp = data(iptr) + data(iptr - 1) * &H100
      vRet(optr) = Int(iTmp / iShift) And &HFF
      iLen = iLen - 8
      iptr = iptr - 1
      optr = optr - 1
   Loop

   If iLen > 0 Then vRet(optr) = Int(data(iptr) / iShift) And (2 ^ iLen - 1)
   GetBits = vRet
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetBits()
   Dim data(2) As Byte
   Dim rdata() As Byte
   data(0) = 1: data(1) = 2: data(2) = 3
   rdata = GetBits(data, 0, 0, 8): assert 1, rdata(0)
   rdata = GetBits(data, 1, 0, 8): assert 2, rdata(0)
   rdata = GetBits(data, 2, 0, 8): assert 3, rdata(0)

   data(0) = 255: data(1) = 129: data(2) = 191
   rdata = GetBits(data, 0, 0, 8): assert 255, rdata(0)
   rdata = GetBits(data, 1, 0, 8): assert 129, rdata(0)
   rdata = GetBits(data, 2, 0, 8): assert 191, rdata(0)

   data(0) = 255: data(1) = 129: data(2) = 191
   rdata = GetBits(data, 1, 0, 16): assert 129, rdata(0): assert 191, rdata(1)

   data(0) = 255: data(1) = 128: data(2) = 254
   rdata = GetBits(data, 0, 1, 7): assert 127, rdata(0)
   rdata = GetBits(data, 0, 2, 6): assert 63, rdata(0)
   rdata = GetBits(data, 0, 3, 5): assert 31, rdata(0)
   rdata = GetBits(data, 0, 4, 4): assert 15, rdata(0)
   rdata = GetBits(data, 0, 5, 3): assert 7, rdata(0)
   rdata = GetBits(data, 0, 6, 2): assert 3, rdata(0)
   rdata = GetBits(data, 0, 7, 1): assert 1, rdata(0)

   rdata = GetBits(data, 0, 1, 8): assert 255, rdata(0)
   rdata = GetBits(data, 0, 2, 8): assert 254, rdata(0)

   rdata = GetBits(data, 2, 1, 7): assert 126, rdata(0)
   rdata = GetBits(data, 2, 2, 6): assert 62, rdata(0)
   rdata = GetBits(data, 2, 3, 5): assert 30, rdata(0)
   rdata = GetBits(data, 2, 4, 4): assert 14, rdata(0)
   rdata = GetBits(data, 2, 5, 3): assert 6, rdata(0)
   rdata = GetBits(data, 2, 6, 2): assert 2, rdata(0)
   rdata = GetBits(data, 2, 7, 1): assert 0, rdata(0)
End Sub
