Function GetLastCell(Optional ByVal oSheet As Worksheet) As Range
   If oSheet Is Nothing Then Set oSheet = ActiveSheet
   Set GetLastCell = oSheet.Cells.SpecialCells(xlCellTypeLastCell)
'   Set GetLastCell = oSheet.UsedRange.Rows(oSheet.UsedRange.Rows.Count)
End Function

' =================== VBA: TEST: Begin ===================

Public Sub xUnitTest_beans_GetLastCell()
   assert "$E$4", GetLastCell(ThisWorkbook.Worksheets("GetLastCell")).Address
End Sub
