Function HexGetBitsStr(ByVal str As String, ByVal iWord As Long, ByVal iBit As Long, ByVal iLen As Long, Optional ByVal sDelim As String = "") As String
   Dim i As Long
   Dim v As Variant
   Dim aTmp() As Byte
   Dim iShift As Long
   Dim iptr As Long, optr As Long
   Dim s As String

   If Len(str) Mod 2 <> 0 Then
      HexGetBitsStr = WorksheetFunction.NA
      Exit Function
   End If
   i = Len(str) / 2

   ReDim aRet(i - 1) As Byte
   Do While i > 0
      aRet(i - 1) = CByte("&H" & Mid(str, i * 2 - 1, 2))
      i = i - 1
   Loop

   iShift = 2 ^ (7 - ((iLen + iBit - 1) Mod 8))
   iptr = iWord + Int((iLen + iBit - 1) / 8)
   optr = Int((iLen - 1) / 8)

   s = ""
   Do While iLen + iBit > 8
      i = aRet(iptr) + aRet(iptr - 1) * &H100
      If s = "" Then
         s = Right("0" & Hex(Int(i / iShift) And &HFF), 2)
      Else
         s = Right("0" & Hex(Int(i / iShift) And &HFF), 2) & sDelim & s
      End If
      iLen = iLen - 8
      iptr = iptr - 1
      optr = optr - 1
   Loop

   If iLen > 0 Then
      If s = "" Then
         s = Right("0" & Hex(Int(aRet(iptr) / iShift) And &HFF), 2)
      Else
         s = Right("0" & Hex(Int(aRet(iptr) / iShift) And &HFF), 2) & sDelim & s
      End If
   End If

   HexGetBitsStr = s
End Function
