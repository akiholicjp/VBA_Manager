Function PushCollection(ByRef o As Object, newItem As Variant) As Variant
   If o Is Nothing Then Set o = New Collection
   With o
      .Add newItem
      If IsObject(.Item(.Count)) Then
         Set PushCollection = .Item(.Count)
      Else
         PushCollection = .Item(.Count)
      End If
   End With
End Function

' =================== VBA: TEST: Begin ===================

' VBA: Import PopCollection.bas

Public Sub xUnitTest_beans_PushCollection()
   Dim o As Object
   assert "A", PushCollection(o, "A")
   assert "B", PushCollection(o, "B")
   assert "B", PopCollection(o)
   assert "A", PopCollection(o)
   assert Null, PopCollection(o)
End Sub
