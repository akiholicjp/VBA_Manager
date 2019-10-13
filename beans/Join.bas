Function Join(vList As Variant, Optional ByVal sDelim As String = " ") As String
   Dim s As String, v As Variant
   s = ""
   For Each v In vList
      If s = "" Then
         s = v
      Else
         s = s & sDelim & v
      End If
   Next v
   Join = s
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_Join()
   assert "A BC DEF", Join(Array("A", "BC", "DEF"))
   assert "ABCDEF", Join(Array("A", "BC", "DEF"), "")
   assert "A" & vbCrLf & "BC" & vbCrLf & "DEF", Join(Array("A", "BC", "DEF"), vbCrLf)
   Dim c As New Collection
   c.Add "X"
   c.Add "YY"
   c.Add "ZZZ"
   assert "X YY ZZZ", Join(c)
   assert "XYYZZZ", Join(c, "")
   assert "X" & vbCrLf & "YY" & vbCrLf & "ZZZ", Join(c, vbCrLf)
End Sub
