' VBA: Import ArrayToCollection.bas
Function ToCollection(ByRef target As Variant) As Collection
   If IsArray(target) Then
      Set ToCollection = ArrayToCollection(target)
   ElseIf IsObject(target) Then
      If TypeName(target) = "Collection" Then
         Set ToCollection = target
      Else
         Set ToCollection = New Collection
         ToCollection.Add target
      End If
   Else
      Set ToCollection = New Collection
      ToCollection.Add target
   End If
End Function
