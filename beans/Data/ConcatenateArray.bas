Function ConcatenateArray(Target As Variant) As String
   Dim s As String, ss As String, v As Variant
   On Error Resume Next
   s = ""

   If IsObject(Target) Then
      ' Objectの場合、単独変換、又は、Iteratorの成功した方が残る。両方だめなら初期値。
      s = CStr(Target)
      s = CStr(Target.Text)
      For Each v In Target
         ss = ""
         ss = CStr(v)
         ss = CStr(v.Text)
         s = s & ss
      Next v
   ElseIf IsArray(Target) Then
      For Each v In Target
         ss = ""
         ss = CStr(v)
         ss = CStr(v.Text)
         s = s & ss
      Next v
   Else
      s = CStr(Target)
   End If
   ConcatenateArray = s
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ConcatenateArray()
   Dim oCol As Object
   Dim s As String
   assert "", ConcatenateArray(oCol)
   Set oCol = New Collection
   assert "", ConcatenateArray(oCol)
   oCol.Add "A"
   assert "A", ConcatenateArray(oCol)
   oCol.Add "B"
   assert "AB", ConcatenateArray(oCol)
   oCol.Add "C"
   assert "ABC", ConcatenateArray(oCol)
   oCol.Remove 2
   assert "AC", ConcatenateArray(oCol)
   oCol.Remove 2
   assert "A", ConcatenateArray(oCol)
   oCol.Remove 1
   assert "", ConcatenateArray(oCol)

   Dim v(2) As Long
   v(0) = 1
   v(1) = 22
   v(2) = 33
   assert "12233", ConcatenateArray(v)

   Dim w(2) As Variant
   w(0) = "A"
   w(1) = "BB"
   w(2) = "CCC"
   assert "ABBCCC", ConcatenateArray(w)

   assert "X", ConcatenateArray("X")

   assert "1", ConcatenateArray(1)

End Sub
