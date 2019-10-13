Function GetEachFromDictionary(o As Variant, Optional vKey As Variant, Optional vVal As Variant) As Variant
   Static oDic As Object
   Static i As Long

   If o Is Nothing Then GoTo ERR_PROC
   If Not oDic Is o Then
      Set oDic = o
      i = 1
   End If
   If i > oDic.Count Then GoTo ERR_PROC
   vKey = oDic.Keys()(i - 1)
   If IsObject(vVal) Then
      Set vVal = oDic(vKey)
   Else
      Let vVal = oDic(vKey)
   End If
   i = i + 1
   GetEachFromDictionary = Array(vKey, vVal)
   Exit Function
ERR_PROC:
   Set oDic = Nothing
   vKey = Null
   vVal = Null
   GetEachFromDictionary = Null
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetEachFromDictionary()
   Dim o As Object
   Set o = CreateObject("Scripting.Dictionary")
   o.Add Key:=1, Item:=2
   o.Add Key:="ABC", Item:=3.0
   o.Add Key:=4.0, Item:="cde"
   assert Array(1, 2), GetEachFromDictionary(o)
   assert Array("ABC", 3.0), GetEachFromDictionary(o)
   assert Array(4.0, "cde"), GetEachFromDictionary(o)
   assert null, GetEachFromDictionary(o)
End Sub
