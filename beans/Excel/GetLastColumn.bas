Function GetLastColumn(ByVal Row As Long, Optional ByVal oSheet As Worksheet) As Object
   If oSheet Is Nothing Then Set oSheet = ActiveSheet
   With oSheet
      Set GetLastColumn = .Cells(Row, .Columns.Count).End(xlToLeft)
   End With
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetLastColumn()
   assert "$C$4", GetLastColumn(4, ThisWorkbook.Worksheets("GetLastCell")).Address
End Sub
