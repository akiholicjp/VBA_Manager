Function GetEachFromCollection(o As Variant, Optional vVal As Variant) As Variant
   Static oCol As Object
   Static i As Long
   If o Is Nothing Then GoTo ERR_PROC
   If Not oCol Is o Then
      Set oCol = o
      i = 1
   End If
   If i > oCol.Count Then GoTo ERR_PROC
   If IsObject(oCol(i)) Then
      Set vVal = oCol(i)
   Else
      Let vVal = oCol(i)
   End If
   i = i + 1
   If IsObject(vVal) Then
      Set GetEachFromCollection = vVal
   Else
      Let GetEachFromCollection = vVal
   End If
   Exit Function
ERR_PROC:
   Set oCol = Nothing
   vVal = Null
   GetEachFromCollection = Null
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetEachFromCollection()
   Dim o As New Collection
   o.Add 1
   o.Add "ABC"
   o.Add 2.0
   assert 1, GetEachFromCollection(o)
   assert "ABC", GetEachFromCollection(o)
   assert 2.0, GetEachFromCollection(o)
   assert null, GetEachFromCollection(o)
End Sub
