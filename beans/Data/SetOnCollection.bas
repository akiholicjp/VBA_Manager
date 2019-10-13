Sub SetOnCollection(col As Collection, key As Variant, val As Variant)
   On Error Resume Next
   col.Remove key
   On Error GoTo 0
   col.Add Item:=val, Key:=key
End Sub

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_SetOnCollection()
   Dim oCol As New Collection
   Call SetOnCollection(oCol, "AAA", "BBB")
   assert "BBB", oCol("AAA")
   Call SetOnCollection(oCol, "AAA", "CCC")
   assert "CCC", oCol("AAA")

   Set oCol = New Collection
   oCol.Add Key:="XXX", Item:="ZZZ"
   assert "ZZZ", oCol("XXX")
   Call SetOnCollection(oCol, "XXX", "YYY")
   assert "YYY", oCol("XXX")
End Sub
