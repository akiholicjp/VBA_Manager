Function InCollection(ByRef vTarget As Variant, ByRef obj As Variant, Optional ByVal From As Long = 1, Optional ByVal None As Long = 0) As Long
   Dim i As Long
   Dim iRet As Long
   iRet = None
   For i = 1 To obj.Count
      If IsObject(obj(i)) Then
         If vTarget Is obj(i) Then iRet = i - 1 + From
      Else
         If vTarget = obj(i) Then iRet = i - 1 + From
      End If
      If iRet >= From Then Exit For
   Next i
   InCollection = iRet
End Function
