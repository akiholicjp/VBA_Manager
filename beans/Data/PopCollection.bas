Function PopCollection(ByRef o As Object) As Variant
   With o
      If .Count > 0 Then
         If IsObject(.Item(.Count)) Then
            Set PopCollection = .Item(.Count)
         Else
            PopCollection = .Item(.Count)
         End If
         .Remove .Count
      Else
         PopCollection = Null
      End If
   End With
End Function

' =================== VBA: TEST: Begin ===================

' VBA: Import PushCollection.bas

Public Sub xUnitTest_beans_PopCollection()
   Dim o As Object
   assert "A", PushCollection(o, "A")
   assert "B", PushCollection(o, "B")
   assert "B", PopCollection(o)
   assert "A", PopCollection(o)
   assert Null, PopCollection(o)
End Sub
