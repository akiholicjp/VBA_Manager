Function GetCellComment(oCell As Range) As String
   If oCell.Comment Is Nothing Then
      GetCellComment = ""
   Else
      GetCellComment = oCell.Comment.Text
   End If
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetCellComment()
   assert "", GetCellComment(ThisWorkbook.Worksheets("GetCellComment").Range("A1"))
   assert "ABCTEST", GetCellComment(ThisWorkbook.Worksheets("GetCellComment").Range("A2"))
End Sub
