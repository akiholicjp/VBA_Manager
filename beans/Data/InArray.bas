Function InArray(ByRef vTarget As Variant, ByRef ary As Variant, Optional ByVal From As Long = 0, Optional ByVal None As Long = -1) As Long
   Dim i As Long
   Dim iRet As Long
   iRet = None
   For i = LBound(ary) To UBound(ary)
      If IsObject(ary(i)) Then
         If vTarget Is ary(i) Then iRet = i - LBound(ary) + From
      Else
         If vTarget = ary(i) Then iRet = i - LBound(ary) + From
      End If
      If iRet >= From Then Exit For
   Next i
   InArray = iRet
End Function
