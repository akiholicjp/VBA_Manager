Function ArrayToCollection(ByRef ary As Variant) As Collection
   Dim v As Variant
   Set ArrayToCollection = New Collection
   For Each v In ary
      ArrayToCollection.Add v
   Next v
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ArrayToCollection()
   assert "[1%,""ABC"",5#]", Dump(ArrayToCollection(Array(1, "ABC", 5.0)))
End Sub
