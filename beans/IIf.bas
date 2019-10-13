Function IIf(ByVal b As Boolean, vt As Variant, vf As Variant) As Variant
   If b Then
      If IsObject(vt) Then
         Set IIf = vt
      Else
         IIf = vt
      End If
   Else
      If IsObject(vf) Then
         Set IIf = vf
      Else
         IIf = vf
      End If
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_IIf()
   assert 1, IIf(True, 1, 2)
   assert 2, IIf(False, 1, 2)
   assert "TEST", IIf(True, "TEST", Nothing)
   assert Nothing, IIf(False, "TEST", Nothing)
End Sub
