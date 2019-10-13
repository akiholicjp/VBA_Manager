' VBA: Import UniRegExp.bas
Function RegReplace(ByVal strTarget As String, Pattern As Variant, Replace As Variant, Optional ByRef RegExp As Variant) As String
   Dim i As Long
   Dim oReg As Object
   If IsMissing(RegExp) Then
      Set oReg = UniRegExp(Null)
   Else
      Set oReg = RegExp
   End If
   With oReg
      If IsArray(Pattern) And IsArray(Replace) Then
         If UBound(Pattern) - LBound(Pattern) = UBound(Replace) - LBound(Replace) Then
            For i = LBound(Pattern) To UBound(Pattern)
               .Pattern = Pattern(i)
               strTarget = .Replace(strTarget, Replace(i))
            Next
         End If
      ElseIf IsObject(Pattern) And IsObject(Replace) Then
         If Pattern.Count = Replace.Count Then
            For i = 1 To Pattern.Count
               .Pattern = Pattern(i)
               strTarget = .Replace(strTarget, Replace(i))
            Next
         End If
      Else
         .Pattern = CStr(Pattern)
         strTarget = .Replace(strTarget, Replace)
      End If
   End With
   RegReplace = strTarget
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_RegReplace()
   assert "T", RegReplace("TEST", "T.S", "")
   assert "TEST", RegReplace("TEST", "^E.S", "")
   assert "TESX", RegReplace("TEST", "T$", "X")
End Sub

