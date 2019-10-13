Function ConcatenateRange(ParamArray Target() As Variant) As String
   Dim sRet As String, s As String, ss As String, v As Variant, vv As Variant
   On Error Resume Next
   sRet = ""
   For Each v In Target
      ' Objectの場合、単独変換、又は、Iteratorの成功した方が残る。両方だめなら初期値。
      s = ""
      If IsArray(v) Then
         For Each vv In v
            ss = ""
            ss = CStr(vv)
            ss = CStr(vv.Text)
            s = s & ss
         Next vv
      ElseIf IsObject(v) Then
         s = CStr(v)
         s = CStr(v.Text)
         For Each vv In v
            ss = ""
            ss = CStr(vv)
            ss = CStr(vv.Text)
            s = s & ss
         Next vv
      Else
         s = CStr(v)
         s = CStr(v.Text)
      End If
      sRet = sRet & s
   Next v
   ConcatenateRange = sRet
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_ConcatenateRange()
   Dim oCol As Object
   Dim s As String
   assert "", ConcatenateRange(oCol)
   Set oCol = New Collection
   assert "", ConcatenateRange(oCol)
   oCol.Add "A"
   assert "A", ConcatenateRange(oCol)
   oCol.Add "B"
   assert "AB", ConcatenateRange(oCol)
   oCol.Add "C"
   assert "ABC", ConcatenateRange(oCol)
   oCol.Remove 2
   assert "AC", ConcatenateRange(oCol)
   oCol.Remove 2
   assert "A", ConcatenateRange(oCol)
   oCol.Remove 1
   assert "", ConcatenateRange(oCol)

   Dim v(2) As Long
   v(0) = 1
   v(1) = 22
   v(2) = 33
   assert "12233", ConcatenateRange(v)

   Dim w(2) As Variant
   w(0) = "A"
   w(1) = "BB"
   w(2) = "CCC"
   assert "ABBCCC", ConcatenateRange(w)

   assert "X", ConcatenateRange("X")

   assert "1", ConcatenateRange(1)

End Sub
