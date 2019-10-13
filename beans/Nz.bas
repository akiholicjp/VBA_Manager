Function Nz(ByRef v As Variant, Optional ByVal vv As Variant = "") As Variant
   If IsNull(v) Then
      If IsObject(vv) Then
         Set Nz = vv
      Else
         Nz = vv
      End If
   Else
      If IsObject(v) Then
         Set Nz = v
      Else
         Nz = v
      End If
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_Nz()
   Dim a As Variant
   assert Empty, Nz(a)
   a = Null
   assert "", Nz(a)
   assert "AA", Nz(a, "AA")
   assert Nothing, Nz(a, Nothing)
   a = "A"
   assert "A", Nz(a, 1)
   a = Empty
   assert Empty, Nz(a, "TEST")
End Sub
