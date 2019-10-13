Function SelectValueFromList(ByVal sList As String, ByVal sKey As String, Optional ByVal sDelimList As String = ",", Optional ByVal sDelimKey As String = ":", Optional vOnNothing As Variant = "") As Variant
   Dim v As Variant, s As String
   For Each v In Split(sList, sDelimList)
      v = Trim(v)
      s = Trim(sKey & sDelimKey)
      If v Like s & "*" Then
         SelectValueFromList = Trim(Mid(v, Len(s) + 1))
         Exit Function
      End If
   Next v
   If VarType(vOnNothing) = 10 Then
      SelectValueFromList = CVErr(xlErrNA)
   Else
      SelectValueFromList = vOnNothing
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_SelectValueFromList()
   assert "3", SelectValueFromList("a:1,b:2,c:3,d:4", "c", sDelimList:=",", sDelimKey:=":")
   assert "3", SelectValueFromList("a: 1, b: 2, c: 3, d: 4", "c", sDelimList:=",", sDelimKey:=":")
   assert "", SelectValueFromList("a:1,b:2,c:3,d:4", "z", sDelimList:=",", sDelimKey:=":", vOnNothing:="")
End Sub
