Function CloneDictionary(ByRef oDic As Object) As Object
   If oDic Is Nothing Then
      Set CloneDictionary = Nothing
      Exit Function
   ElseIf TypeName(oDic) <> "Dictionary" Then
      Set CloneDictionary = Nothing
      Exit Function
   End If
   Dim oDicNew As Object
   Dim v As Variant
   Set oDicNew = CreateObject("Scripting.Dictionary")

   oDicNew.CompareMode = oDic.CompareMode
   For Each v in oDic.Keys
      oDicNew.Add v, oDic(v)
   Next

   Set CloneDictionary = oDicNew
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_CloneDictionary()
   assert Nothing, CloneDictionary(Nothing)
   assert Nothing, CloneDictionary(New Collection)

   Dim oDic1 As Object, oDic2 As Object
   Set oDic1 = CreateObject("Scripting.Dictionary")
   assert ObjPtr(oDic1), ObjPtr(oDic1)
   assert Dump(oDic1), Dump(oDic1)
   oDic1.Add "A", "B"
   assert ObjPtr(oDic1), ObjPtr(oDic1)
   assert Dump(oDic1), Dump(oDic1)

   Set oDic2 = CloneDictionary(oDic1)
   assertNe ObjPtr(oDic2), ObjPtr(oDic1)
   assert Dump(oDic2), Dump(oDic1)

   oDic1.Add "B", "C"
   Set oDic2 = CloneDictionary(oDic1)
   assert Dump(oDic2), Dump(oDic1)

   oDic1.Add "C", "D"
   assertNe Dump(oDic2), Dump(oDic1)
End Sub
