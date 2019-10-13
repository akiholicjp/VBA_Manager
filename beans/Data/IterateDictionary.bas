Function IterateDictionary(o As Variant) As Collection
   Dim v As Variant
   Set IterateDictionary = New Collection
   With IterateDictionary
      For Each v In o.Keys
         .Add Array(v, o(v))
      Next v
   End With
End Function
