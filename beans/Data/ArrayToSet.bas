' VBA: Import ../NewDic.bas
' VBA: Import ArrayToCollection.bas
Function ArrayToSet(ByRef ary As Variant, Optional ByVal IgnoreError As Boolean = True) As Variant
   Dim oSet As Object
   Dim v As Variant
   Set oSet = NewDic()
   For Each v In ary
      If IsError(v) And IgnoreError Then
         ' Ignore
      ElseIf IsObject(v) Then
         If Not oSet.Exists(ObjPtr(v)) Then oSet.Add Key:=ObjPtr(v), Item:=v
      Else
         If Not oSet.Exists(v) Then oSet.Add Key:=v, Item:=v
      End If
   Next v
   Set ArrayToSet = ArrayToCollection(oSet.Items())
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ArrayToSet()
   Dim o As New Collection
   With o
      .Add 1
      .Add "ABC"
      .Add 2.0
      .Add 1
      .Add "ABC"
   End With
   assert "[1%,""ABC"",2#]", Dump(ArrayToSet(o))
End Sub
