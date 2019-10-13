Function CCDbl(v As Variant, Optional ByVal OnErrDbl As Double = -1, Optional ByRef Ret As Double) As Double
   If VarType(v) <= vbNull Or VarType(v) >= vbObject Then
      CCDbl = OnErrDbl
   ElseIf VarType(v) = vbString Then
      If IsNumeric(v) Then
         CCDbl = CDbl(v)
      Else
         CCDbl = OnErrDbl
      End If
   Else
      CCDbl = CDbl(v)
   End If
   Ret = CCDbl
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_CCDbl()
   Dim r As Double
   assert 123, CCDbl("123")
   assert 123.4, CCDbl("123.4")
   assert -1, CCDbl("123x")
   assert -1, CCDbl("123.4x")
   assert 123.4, CCDbl("123.4", -1)
   assert 123.4, CCDbl("123.4", 0)
   assert -1, CCDbl("123.4x", -1)
   assert 0, CCDbl("123.4x", 0)
   assert 123.4, CCDbl("123.4", OnErrDbl:=-1)
   assert 123.4, CCDbl("123.4", OnErrDbl:=0)
   assert -1, CCDbl("123.4x", OnErrDbl:=-1)
   assert 0, CCDbl("123.4x", OnErrDbl:=0)

   r = 0
   assert 123, CCDbl("123", Ret:=r)
   assert 123, r
   r = 999
   assert 123.4, CCDbl("123.4", Ret:=r)
   assert 123.4, r
   r = 999
   assert -1, CCDbl("123.4x", -1, Ret:=r)
   assert -1, r
   r = 999
   assert 0, CCDbl("123.4x", 0, Ret:=r)
   assert 0, r
   r = 999
   assert 123.4, CCDbl("123.4", -1, r)
   assert 123.4, r
   r = 999
   assert 123.4, CCDbl("123.4", 0, r)
   assert 123.4, r
   r = 999
   assert -1, CCDbl("123.4x", -1, r)
   assert -1, r
   r = 999
   assert 0, CCDbl("123.4x", 0, r)
   assert 0, r
   r = 999
   assert 123.4, CCDbl("123.4", OnErrDbl:=-1, Ret:=r)
   assert 123.4, r
   r = 999
   assert 123.4, CCDbl("123.4", OnErrDbl:=0, Ret:=r)
   assert 123.4, r
   r = 999
End Sub
