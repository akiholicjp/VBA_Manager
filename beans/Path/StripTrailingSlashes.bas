Function StripTrailingSlashes(ByVal sText As String) As String
   Dim i As Long
   Dim s As String

   StripTrailingSlashes = ""
   sText = Trim(sText)
   For i = Len(sText) To 1 Step -1
      s = Mid(sText, i, 1)
      If s <> "\" And s <> "/" Then
         StripTrailingSlashes = Mid(sText, 1, i)
         Exit For
      End If
   Next i
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_StripTrailingSlashes()
   assert ThisWorkbook.Path, StripTrailingSlashes(ThisWorkbook.Path)
   assert ThisWorkbook.Path, StripTrailingSlashes(ThisWorkbook.Path & "/")
   assert ThisWorkbook.Path, StripTrailingSlashes(ThisWorkbook.Path & "\")
   assert ".", StripTrailingSlashes(".")
   assert ".", StripTrailingSlashes("./")
   assert ".", StripTrailingSlashes(".\")
End Sub
