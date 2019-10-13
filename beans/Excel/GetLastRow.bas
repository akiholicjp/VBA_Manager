Function GetLastRow(ByVal Column As Long, Optional ByVal oSheet As Worksheet) As Object
   If oSheet Is Nothing Then Set oSheet = ActiveSheet
   With oSheet
      Set GetLastRow = .Cells(.Rows.Count, Column).End(xlUp)
   End With
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetLastRow()
   assert "$C$4", GetLastRow(3, ThisWorkbook.Worksheets("GetLastCell")).Address
End Sub
