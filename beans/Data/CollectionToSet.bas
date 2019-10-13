' VBA: Import ../NewDic.bas
' VBA: Import ArrayToCollection.bas
Function CollectionToSet(ByRef obj As Variant) As Variant
   Dim oSet As Object
   Set oSet = NewDic()
   For Each v In obj
      If IsObject(v) Then
         If Not oSet.Exists(ObjPtr(v)) Then oSet.Add Key:=ObjPtr(v), Item:=v
      Else
         If Not oSet.Exists(v) Then oSet.Add Key:=v, Item:=v
      End If
   Next v
   Set CollectionToSet = ArrayToCollection(oSet.Items())
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_CollectionToSet()
   Dim o As New Collection
   With o
      .Add 1
      .Add "ABC"
      .Add 2.0
      .Add 1
      .Add "ABC"
   End With
   assert "[1%,""ABC"",2#]", Dump(CollectionToSet(o))
End Sub
