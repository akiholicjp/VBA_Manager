' VBA: Import CollectionToArray.bas
Function ToArray(ByRef target As Variant) As Variant
   Dim ary() As Variant
   If IsArray(target) Then
      ToArray = target
   ElseIf IsObject(target) Then
      If TypeName(target) = "Collection" Then
         ToArray = CollectionToArray(target)
      Else
         ReDim ary(0 To 0)
         Set ary(0) = target
         ToArray = ary
      End If
   Else
      ReDim ary(0 To 0)
      ary(0) = target
      ToArray = ary
   End If
End Function