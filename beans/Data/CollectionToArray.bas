Function CollectionToArray(ByRef obj As Variant) As Variant
   Dim i As Long
   Dim aryVals() As Variant
   ReDim aryVals(0 To obj.Count - 1)
   For i = 1 To obj.Count
      If IsObject(obj(i)) Then
         Set aryVals(i - 1) = obj(i)
      Else
         Let aryVals(i - 1) = obj(i)
      End If
   Next i
   CollectionToArray = aryVals
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_CollectionToArray()
   Dim o As New Collection
   With o
      .Add 1
      .Add "ABC"
      .Add 2.0
   End With
   assert "(1%,""ABC"",2#)", Dump(CollectionToArray(o))
End Sub
