' KeyとValueのペアでDictionaryを生成したものを返す。先頭がDictionaryだった場合は、それにKeyとValueのペアを追加する。
Function DicProp(ParamArray Target() As Variant) As Object
   Dim oDic As Object
   Dim iNum As Long, i As Long
   Dim vKey As Variant
   Dim bKey As Boolean
   iNum = UBound(Target) - LBound(Target) + 1
   i = LBound(Target)
   If iNum > 0 Then
      If IsObject(Target(i)) Then
         Set oDic = Target(i)
         i = i + 1
      End If
   End If
   If oDic Is Nothing Then
      Set oDic = CreateObject("Scripting.Dictionary")
   End If
   bKey = True
   Do While i < iNum
      If bKey Then
         If IsObject(Target(i)) Then
            Set vkey = Target(i)
         Else
            vkey = Target(i)
         End If
         bKey = False
      Else
         oDic.Add Key:=vKey, Item:=Target(i)
         bKey = True
      End If
      i = i + 1
   Loop
   Set DicProp = oDic
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_DicProp()
   Dim oDic As Object
   Dim oDic2 As Object

   Set oDic = DicProp()
   Set oDic2 = CreateObject("Scripting.Dictionary")
   assert Dump(oDic2), Dump(oDic)

   Set oDic = DicProp(oDic, "AAA", "BBB", "CCC", "DDD")
   oDic2.Add Key:="AAA", Item:="BBB"
   oDic2.Add Key:="CCC", Item:="DDD"
   assert Dump(oDic2), Dump(oDic)

   Call DicProp(oDic, "EEE", "FFF")
   oDic2.Add Key:="EEE", Item:="FFF"
   assert Dump(oDic2), Dump(oDic)

   Set oDic = DicProp("AAA", "BBB", "CCC", "DDD")
   Set oDic2 = CreateObject("Scripting.Dictionary")
   oDic2.Add Key:="AAA", Item:="BBB"
   oDic2.Add Key:="CCC", Item:="DDD"
   assert Dump(oDic2), Dump(oDic)

   Call DicProp(oDic, 999, "ZZZ")
   oDic2.Add Key:=999, Item:="ZZZ"
   assert Dump(oDic2), Dump(oDic)
End Sub
