Function StripLeadingSlashes(ByVal sText As String) As String
   Dim i As Long
   Dim s As String

   StripLeadingSlashes = ""
   sText = Trim(sText)
   For i = 1 To Len(sText)
      s = Mid(sText, i, 1)
      If s <> "\" And s <> "/" Then
         StripLeadingSlashes = Mid(sText, i)
         Exit For
      End If
   Next i
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_StripLeadingSlashes()
   assert "AAA/BBB", StripLeadingSlashes("/AAA/BBB")
   assert "AAA/BBB", StripLeadingSlashes("AAA/BBB")
   assert "AAA\BBB", StripLeadingSlashes("\AAA\BBB")
   assert "AAA\BBB", StripLeadingSlashes("AAA\BBB")
   assert "", StripLeadingSlashes("/")
   assert "", StripLeadingSlashes("\")
End Sub
