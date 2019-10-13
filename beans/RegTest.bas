' VBA: Import UniRegExp.bas
Function RegTest(ByVal sTest As String, ByVal sPattern As String, Optional ByRef RegExp As Variant) As Boolean
   If IsMissing(RegExp) Then
      RegTest = UniRegExp(sPattern).Test(sTest)
   Else
      RegExp.Pattern = sPattern
      RegTest = RegExp.Test(sTest)
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_RegTest()
   assert True, RegTest("TEST", "T..T")
   assert False, RegTest("TEST", "T.T")
End Sub
