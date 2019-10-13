Function CCLng(v As Variant, Optional ByVal OnErrLng As Long = -1, Optional ByRef Ret As Long) As Long
   If VarType(v) <= vbNull Or VarType(v) >= vbObject Then
         CCLng = OnErrLng
   ElseIf VarType(v) = vbString Then
      If IsNumeric(v) Then
         CCLng = CLng(v)
      Else
         CCLng = OnErrLng
      End If
   Else
      CCLng = CLng(v)
   End If
   Ret = CCLng
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_CCLng()
   Dim r As Long
   assert 123, CCLng("123")
   assert 123, CCLng("123.4")
   assert -1, CCLng("123.4x", -1)
   assert 0, CCLng("123.4x", 0)
   assert 123, CCLng("123.4", -1)
   assert 123, CCLng("123.4", 0)
   assert -1, CCLng("123.4x", -1)
   assert 0, CCLng("123.4x", 0)
   assert 123, CCLng("123.4", OnErrLng:=-1)
   assert 123, CCLng("123.4", OnErrLng:=0)
   assert -1, CCLng("123.4x", OnErrLng:=-1)
   assert 0, CCLng("123.4x", OnErrLng:=0)

   r = 0
   assert 123, CCLng("123", Ret:=r)
   assert 123, r
   r = 999
   assert 123, CCLng("123.4", Ret:=r)
   assert 123, r
   r = 999
   assert -1, CCLng("123.4x", -1, Ret:=r)
   assert -1, r
   r = 999
   assert 0, CCLng("123.4x", 0, Ret:=r)
   assert 0, r
   r = 999
   assert 123, CCLng("123.4", -1, r)
   assert 123, r
   r = 999
   assert 123, CCLng("123.4", 0, r)
   assert 123, r
   r = 999
   assert -1, CCLng("123.4x", -1, r)
   assert -1, r
   r = 999
   assert 0, CCLng("123.4x", 0, r)
   assert 0, r
   r = 999
   assert 123, CCLng("123.4", OnErrLng:=-1, Ret:=r)
   assert 123, r
   r = 999
   assert 123, CCLng("123.4", OnErrLng:=0, Ret:=r)
   assert 123, r
   r = 999
End Sub
