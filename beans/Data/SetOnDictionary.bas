Sub SetOnDictionary(dic As Object, key As Variant, val As Variant)
   ' If dic.Exists(key) Then dic.Remove key
   ' dic.Add key:=key, item:=val
   dic(key) = val
End Sub

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_SetOnDictionary()
   Dim oDic As Object
   Set oDic = CreateObject("Scripting.Dictionary")

   Call SetOnDictionary(oDic, "AAA", "BBB")
   assert "BBB", oDic("AAA")
   Call SetOnDictionary(oDic, "AAA", "CCC")
   assert "CCC", oDic("AAA")

   Set oDic = CreateObject("Scripting.Dictionary")
   oDic.Add Key:="XXX", Item:="ZZZ"
   assert "ZZZ", oDic("XXX")
   Call SetOnDictionary(oDic, "XXX", "YYY")
   assert "YYY", oDic("XXX")
End Sub
