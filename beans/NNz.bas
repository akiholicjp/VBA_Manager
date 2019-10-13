Function NNz(ByRef v As Variant, Optional ByVal vv As Variant = "") As Variant
   If IsNull(v) Or IsEmpty(v) Then
      If IsObject(vv) Then
         Set NNz = vv
      Else
         NNz = vv
      End If
   Else
      If IsObject(v) Then
         Set NNz = v
      Else
         NNz = v
      End If
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_NNz()
   Dim a As Variant
   assert "", NNz(a)
   a = Null
   assert "", NNz(a)
   assert "AA", NNz(a, "AA")
   assert Nothing, NNz(a, Nothing)
   a = "A"
   assert "A", NNz(a, 1)
   a = Empty
   assert "TEST", NNz(a, "TEST")
End Sub
