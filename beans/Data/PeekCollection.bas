Function PeekCollection(ByRef o As Object) As Variant
   With o
      If .Count > 0 Then
         If IsObject(.Item(.Count)) Then
            Set PeekCollection = .Item(.Count)
         Else
            PeekCollection = .Item(.Count)
         End If
      Else
         PeekCollection = Null
      End If
   End With
End Function

' =================== VBA: TEST: Begin ===================

' VBA: Import PopCollection.bas
' VBA: Import PushCollection.bas

Public Sub xUnitTest_beans_PeekCollection()
   Dim o As New Collection
   assert Null, PeekCollection(o)
   assert "A", PushCollection(o, "A")
   assert "A", PeekCollection(o)
   assert "B", PushCollection(o, "B")
   assert "B", PeekCollection(o)
   assert "B", PopCollection(o)
   assert "A", PeekCollection(o)
   assert "A", PopCollection(o)
   assert Null, PeekCollection(o)
   assert Null, PopCollection(o)
End Sub
