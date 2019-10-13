Function UniRegExp(ByVal Pattern As Variant, Optional bSet As Boolean = False, Optional ByVal bGlobal As Boolean = True, Optional ByVal MultiLine As Boolean = True, Optional ByVal IgnoreCase As Boolean = False) As Object
   Static G_RegExp As Object
   If G_RegExp Is Nothing Then
      Set G_RegExp = CreateObject("VBScript.RegExp")
      With G_RegExp
         .Global = bGlobal
         .MultiLine = MultiLine
         .IgnoreCase = IgnoreCase
      End With
   ElseIf bSet Then
      With G_RegExp
         .Global = bGlobal
         .MultiLine = MultiLine
         .IgnoreCase = IgnoreCase
      End With
   End If
   If Not IsNull(Pattern) Then G_RegExp.Pattern = Pattern
   Set UniRegExp = G_RegExp
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_UniRegExp()
   Dim o1 As Object, o2 As Object
   Set o1 = UniRegExp(Null)
   assert "IRegExp2", TypeName(o1)
   Set o2 = UniRegExp(Null)
   assert "IRegExp2", TypeName(o2)
   assert ObjPtr(o1), ObjPtr(o2)
End Sub
